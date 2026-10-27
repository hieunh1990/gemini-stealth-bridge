# Hướng dẫn Tích hợp Client — Gemini Stealth Bridge & BIDV Minutes Service (`guidegsp.md`)

Tài liệu này cung cấp hướng dẫn kỹ thuật chi tiết giúp các lập trình viên, các hệ thống client, microservices, AI agents hoặc bot nghiệp vụ dễ dàng tích hợp và gọi API vào hệ sinh thái **Gemini Stealth Bridge (GSB)** và **Dịch vụ Tổng hợp Biên bản Cuộc họp & Thông báo Kết luận chuẩn Thể thức BIDV**.

---

## 1. Tổng quan Kiến trúc Dịch vụ

Hệ thống được đóng gói và vận hành dưới dạng Docker Compose gồm 2 service microservice chạy song song:

```
                      ┌───────────────────────────────────────────────┐
                      │                 CLIENT APPS                   │
                      │  (Web Apps, Mobile, Bots, Agent, CLI, Scripts)│
                      └───────┬───────────────────────────────┬───────┘
                              │                               │
            1. Gọi Biên bản BIDV (.docx/.pdf)                 │ 2. Gọi LLM trực tiếp
            POST http://100.116.58.118:20135/api/v1/generate  │ POST http://100.116.58.118:20133/v1/chat/completions
                              │                               │
                              ▼                               │
              ┌───────────────────────────────┐               │
              │     bidv-minutes-service      │               │
              │       (FastAPI Engine)        │               │
              │         Port: 20135           │               │
              │  - Trích xuất & sinh thể thức │               │
              │  - Headless LibreOffice PDF   │               │
              └───────────────┬───────────────┘               │
                              │ (nội bộ docker network)       │
                              ▼                               ▼
              ┌───────────────────────────────────────────────┐
              │             gemini-stealth-bridge             │
              │                (GSB Daemon)                   │
              │                 Port: 20133                   │
              │   Proxy Google Gemini Web (Argon, Pro, Flash) │
              └───────────────────────────────────────────────┘
```

### Thông tin Kết nối Mạng Tailscale & Localhost

> **Địa chỉ IP Tailscale của máy chủ:** `100.116.58.118` (Tên máy: `pop-os-1`)  
> Mọi máy tính, laptop, điện thoại hoặc thiết bị trong cùng mạng Tailscale đều có thể kết nối trực tiếp không cần cấu hình port-forward hay NAT.

| Dịch vụ | Địa chỉ qua Tailscale IP | Địa chỉ Localhost | Docker Internal | Mô tả |
|---|---|---|---|---|
| **BIDV Minutes API (JSON/Upload)** | `http://100.116.58.118:20135` | `http://127.0.0.1:20135` | `http://bidv-minutes:20135` | REST API sinh file `.docx` và `.pdf` chuẩn thể thức BIDV |
| **BIDV Minutes Swagger Docs** | `http://100.116.58.118:20135/docs` | `http://127.0.0.1:20135/docs` | - | Giao diện tài liệu API tương tác trực tiếp |
| **BIDV Minutes Web UI** | `http://100.116.58.118:20135/` | `http://127.0.0.1:20135/` | - | Giao diện Web kéo thả file cho người dùng cuối |
| **GSB Core API (OpenAI)** | `http://100.116.58.118:20133/v1` | `http://127.0.0.1:20133/v1` | `http://gsb:20133/v1` | Cổng OpenAI-compatible (kết nối trực tiếp Gemini) |
| **GSB Web Control Plane** | `http://100.116.58.118:20133/ui` | `http://127.0.0.1:20133/ui` | - | Quản lý tài khoản Gemini & xem hạn mức Quota |

---

## 2. Tích hợp Dịch vụ Biên bản BIDV (`bidv-minutes-service`)

Dịch vụ này nhận đầu vào là bản ghi chép cuộc họp (hoặc transcript ghi âm), sử dụng **Gemini Argon** qua GSB để tổng hợp thành Biên bản cuộc họp hoặc Thông báo kết luận theo **100% chuẩn quy định hành chính BIDV**, tự động kết xuất ra cả 2 định dạng **Word (`.docx`)** và **PDF (`.pdf`)**.

### 2.1. API `POST /api/v1/generate` (Nhận JSON Body)

Đây là endpoint chuẩn nhất dành cho các client backend, AI agent, hệ thống tự động hóa.

- **URL qua Tailscale:** `http://100.116.58.118:20135/api/v1/generate`  
  *(hoặc `http://127.0.0.1:20135/api/v1/generate` nếu gọi trên cùng máy)*
- **Method:** `POST`
- **Headers:** `Content-Type: application/json`

#### Tham số Request (JSON)

| Trường | Kiểu dữ liệu | Bắt buộc | Mặc định | Ý nghĩa |
|---|---|---|---|---|
| `transcript` | `string` | **Có** | - | Toàn bộ văn bản ghi chép cuộc họp, phân vai hoặc transcript |
| `doc_type` | `string` | Không | `"BB"` | Loại văn bản: `"BB"` (Biên bản cuộc họp) hoặc `"TBKL"` (Thông báo kết luận) |
| `meeting_title` | `string` | Không | `""` | Tiêu đề hoặc trích yếu phiên họp hiển thị dưới tên văn bản |
| `doc_number` | `string` | Không | `"2710"` | Số hiệu văn bản (Ví dụ: `2710` -> `2710/BB-BIDV`) |
| `signer_name` | `string` | Không | `"HOÀNG VIỆT HÙNG"` | Họ và tên người ký (Lãnh đạo chủ trì) |
| `signer_title` | `string` | Không | `"ĐẠI DIỆN ĐỒNG CHỦ TRÌ..."` | Chức danh người ký |
| `recipients` | `array[str]` | Không | `null` | Danh sách Nơi nhận (nếu để trống hệ thống sẽ tự điền chuẩn) |
| `model` | `string` | Không | `"gemini-argon"` | Model AI xử lý: `"gemini-argon"` (khuyên dùng), `"gemini-pro"`, `"gemini-flash"` |
| `include_base64` | `boolean` | Không | `false` | Nếu `true`, trả về trực tiếp chuỗi base64 của file `.docx` và `.pdf` |

#### Ví dụ Request (cURL)
```bash
curl -X POST "http://100.116.58.118:20135/api/v1/generate" \
  -H "Content-Type: application/json" \
  -d '{
    "transcript": "Ngày 27/10/2026, tại Hội trường BIDV, đồng chí Hoàng Việt Hùng - Tổng Giám đốc BIDV chủ trì cuộc họp cấp cao song phương với Tencent. Hai bên thống nhất: 1. Thí điểm WeChat Pay trong Q1/2027 do Khối Bán lẻ và Ban CNTT đầu mối. 2. Xây dựng Private Cloud hoàn thành trước tháng 12/2026 do Ban CNTT và Trung tâm CNTT thực hiện.",
    "doc_type": "BB",
    "meeting_title": "Phiên họp Cấp cao Song phương BIDV - Tencent",
    "doc_number": "2710",
    "signer_name": "HOÀNG VIỆT HÙNG",
    "signer_title": "ĐẠI DIỆN ĐỒNG CHỦ TRÌ PHIÊN HỌP\nTỔNG GIÁM ĐỐC BIDV",
    "model": "gemini-argon",
    "include_base64": false
  }'
```

#### Cấu trúc Response (JSON)
```json
{
  "status": "success",
  "message": "Đã tạo biên bản thành công theo đúng thể thức chuẩn BIDV",
  "model_used": "gemini-argon",
  "doc_type": "BB",
  "docx_download_url": "http://100.116.58.118:20135/api/download/Bien_ban_hop_BIDV_20261027_094500_abc123.docx",
  "pdf_download_url": "http://100.116.58.118:20135/api/download/Bien_ban_hop_BIDV_20261027_094500_abc123.pdf",
  "docx_filename": "Bien_ban_hop_BIDV_20261027_094500_abc123.docx",
  "pdf_filename": "Bien_ban_hop_BIDV_20261027_094500_abc123.pdf",
  "text_content": "Văn bản nội dung đầy đủ được AI tổng hợp...",
  "preview_text": "Trích đoạn xem trước...",
  "docx_base64": null,
  "pdf_base64": null
}
```

---

### 2.2. API `POST /api/v1/upload` (Upload File Tệp tin)

Dành cho các ứng dụng có sẵn tệp tài liệu `.docx`, `.txt`, `.md` cần gửi trực tiếp.

- **URL:** `http://100.116.58.118:20135/api/v1/upload`
- **Method:** `POST`
- **Headers:** `Content-Type: multipart/form-data`

```bash
curl -X POST "http://100.116.58.118:20135/api/v1/upload" \
  -F "file=@/duong/dan/ghi_chep_hop.docx" \
  -F "doc_type=BB" \
  -F "model=gemini-argon" \
  -F "meeting_title=Phiên họp Cấp cao Song phương BIDV - Tencent"
```

---

### 2.3. API `GET /api/download/{filename}` (Tải tệp kết quả)

Client dùng đường dẫn `docx_download_url` hoặc `pdf_download_url` nhận được từ API sinh biên bản để tải tệp về:

```bash
# Tải file Word
curl -O "http://100.116.58.118:20135/api/download/Bien_ban_hop_BIDV_20261027_094500_abc123.docx"

# Tải file PDF
curl -O "http://100.116.58.118:20135/api/download/Bien_ban_hop_BIDV_20261027_094500_abc123.pdf"
```

---

### 2.4. API `GET /api/health` & `GET /api/v1/models`

- **Kiểm tra trạng thái kết nối:**
  ```bash
  curl -s http://100.116.58.118:20135/api/health
  ```
  *Phản hồi:*
  ```json
  {
    "status": "ok",
    "service": "bidv-minutes",
    "gsb_connected": true,
    "gsb_status": "connected",
    "available_models": ["gemini-flash-lite", "gemini-flash", "gemini-pro", "gemini-argon"],
    "default_model": "gemini-argon",
    "output_dir": "/app/outputs"
  }
  ```

---

## 3. Tích hợp Trực tiếp với Gemini Stealth Bridge (`gsb`)

Ngoài dịch vụ tạo biên bản BIDV, client có thể sử dụng trực tiếp **GSB** như một máy chủ OpenAI chuẩn (`OpenAI-compatible HTTP Daemon`) để tương tác với Gemini Web.

### 3.1. Thông số kết nối GSB
- **Base URL (qua Tailscale):** `http://100.116.58.118:20133/v1`
- **Base URL (Localhost):** `http://127.0.0.1:20133/v1`
- **API Key:** `public` (hoặc giá trị trong cấu hình `BRIDGE_API_KEY`)
- **Headers:** `Authorization: Bearer public`

### 3.2. Danh sách Model đề xuất
- **`gemini-argon`**: Model cao cấp nhất với năng lực tư duy, tổng hợp chính sách và phân tích đa chiều cực kỳ sắc bén (rất phù hợp cho văn bản chỉ đạo cấp cao).
- **`gemini-pro`**: Model cân bằng giữa tốc độ và độ sâu phân tích.
- **`gemini-flash`** & **`gemini-flash-lite`**: Tốc độ phản hồi cực nhanh, độ trễ thấp.

### 3.3. Ví dụ gọi Chat Completions (OpenAI SDK Python)
```python
from openai import OpenAI

# Kết nối qua Tailscale IP
client = OpenAI(
    base_url="http://100.116.58.118:20133/v1",
    api_key="public"
)

response = client.chat.completions.create(
    model="gemini-argon",
    messages=[
        {"role": "system", "content": "Bạn là chuyên gia phân tích tài chính BIDV."},
        {"role": "user", "content": "Tóm tắt xu hướng ứng dụng GenAI trong Core Banking năm 2026."}
    ]
)

print(response.choices[0].message.content)
```

---

## 4. Thư viện Code Mẫu Tích hợp Client (BIDV Minutes API)

### 4.1. Python Client (với `requests`)

```python
import requests

SERVER_IP = "100.116.58.118"  # Tailscale IP máy chủ

def create_bidv_minutes(transcript_text, doc_type="BB", title=""):
    url = f"http://{SERVER_IP}:20135/api/v1/generate"
    payload = {
        "transcript": transcript_text,
        "doc_type": doc_type,
        "meeting_title": title,
        "model": "gemini-argon",
        "include_base64": False
    }
    
    resp = requests.post(url, json=payload, timeout=120)
    resp.raise_for_status()
    data = resp.json()
    
    print(f"Trạng thái: {data['status']}")
    print(f"Link tải Word: {data['docx_download_url']}")
    print(f"Link tải PDF:  {data['pdf_download_url']}")
    return data

# Thử nghiệm
if __name__ == "__main__":
    sample = "Cuộc họp Ban Chỉ đạo Chuyển đổi số BIDV ngày 27/10/2026..."
    create_bidv_minutes(sample, doc_type="BB", title="Họp Ban Chỉ đạo CĐS")
```

### 4.2. Node.js / TypeScript Client

```typescript
const SERVER_IP = "100.116.58.118"; // Tailscale IP máy chủ

interface GenerateResponse {
  status: string;
  docx_download_url: string;
  pdf_download_url: string;
  docx_filename: string;
  pdf_filename: string;
  text_content: string;
}

async function generateMinutes(transcript: string): Promise<GenerateResponse> {
  const res = await fetch(`http://${SERVER_IP}:20135/api/v1/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      transcript: transcript,
      doc_type: "BB",
      model: "gemini-argon"
    })
  });

  if (!res.ok) {
    throw new Error(`API Error: ${res.statusText}`);
  }

  const data: GenerateResponse = await res.json();
  console.log("Word Document:", data.docx_download_url);
  console.log("PDF Document:", data.pdf_download_url);
  return data;
}
```

### 4.3. Golang Client

```go
package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"time"
)

type RequestPayload struct {
	Transcript   string `json:"transcript"`
	DocType      string `json:"doc_type"`
	MeetingTitle string `json:"meeting_title"`
	Model        string `json:"model"`
}

type ResponsePayload struct {
	Status          string `json:"status"`
	DocxDownloadURL string `json:"docx_download_url"`
	PdfDownloadURL  string `json:"pdf_download_url"`
}

func main() {
	payload := RequestPayload{
		Transcript:   "Bản ghi chép nội dung họp...",
		DocType:      "BB",
		MeetingTitle: "Phiên họp Cấp cao BIDV",
		Model:        "gemini-argon",
	}

	body, _ := json.Marshal(payload)
	client := &http.Client{Timeout: 120 * time.Second}

	resp, err := client.Post("http://100.116.58.118:20135/api/v1/generate", "application/json", bytes.NewBuffer(body))
	if err != nil {
		panic(err)
	}
	defer resp.Body.Close()

	respBytes, _ := io.ReadAll(resp.Body)
	var res ResponsePayload
	json.Unmarshal(respBytes, &res)

	fmt.Printf("File Word: %s\nFile PDF: %s\n", res.DocxDownloadURL, res.PdfDownloadURL)
}
```

### 4.4. C# / .NET Client (`HttpClient`)

```csharp
using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

class Program
{
    static async Task Main(string[] args)
    {
        var client = new HttpClient { Timeout = TimeSpan.FromMinutes(3) };
        var payload = new
        {
            transcript = "Nội dung ghi chép cuộc họp lãnh đạo...",
            doc_type = "BB",
            model = "gemini-argon"
        };

        var content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
        var response = await client.PostAsync("http://100.116.58.118:20135/api/v1/generate", content);

        if (response.IsSuccessStatusCode)
        {
            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            Console.WriteLine("Word: " + doc.RootElement.GetProperty("docx_download_url").GetString());
            Console.WriteLine("PDF:  " + doc.RootElement.GetProperty("pdf_download_url").GetString());
        }
    }
}
```

---

## 5. Quy chuẩn Thể thức Hành chính BIDV Được Tự Động Hóa

Dịch vụ áp dụng nghiêm ngặt các tiêu chuẩn của **`skill bidv-meeting-minutes`**:

1. **Khổ giấy & Căn lề:** Khổ Letter (8.5 x 11.0 inch). Căn lề: Trên 0.75", Dưới 0.75", Trái 1.0", Phải 0.75".
2. **Phông chữ:** `Times New Roman` xuyên suốt toàn bộ văn bản; Cỡ chữ thân bài `13pt`; Khoảng cách dòng `1.15 line`, cách đoạn `6pt`.
3. **Bảng Header & Footer không viền (Borderless):**
   - **Header:** Cột trái ghi "NGÂN HÀNG TMCP ĐẦU TƯ VÀ PHÁT TRIỂN VIỆT NAM" + Số hiệu văn bản; Cột phải ghi "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" + Địa danh ngày tháng.
   - **Footer:** Cột trái là mục "Nơi nhận:" (`11pt - 12pt`); Cột phải là Khung chữ ký với cờ `cantSplit` (chống bị cắt rời sang trang tiếp theo).
4. **Cấu trúc Thân bài chuẩn:**
   - **2 đoạn dẫn nhập:** Nêu đầy đủ địa điểm, người chủ trì, thành phần tham dự và câu dẫn chỉ đạo.
   - **Mục 1. Bối cảnh cuộc họp:** Đúng 3 đoạn văn xuôi, đoạn 3 tóm tắt các điểm nghẽn trọng tâm dạng `(i)...(ii)...(iii)...(iv)...(v)...`
   - **Mục 2. Chỉ đạo triển khai:**
     - `2.1. Định hướng và chỉ đạo chung`: 4-5 nguyên tắc, mỗi nguyên tắc mở đầu bằng cụm động từ mạnh in đậm có dấu hai chấm (`**1. Đặt an toàn hệ thống lên hàng đầu:** ...`).
     - `2.2. Nội dung triển khai cụ thể: nội dung - đơn vị thực hiện - thời hạn`: Chia nhỏ theo từng dự án/chủ đề, mỗi đầu việc bắt buộc có 3 gạch đầu dòng rõ ràng:
       - *- Nội dung chỉ đạo và thống nhất:* ...
       - *- Đơn vị thực hiện:* Giao Ban/Khối...
       - *- Thời hạn hoàn thành:* *in nghiêng mốc thời gian*
5. **Làm sạch mã hiển thị:** Tự động loại bỏ toàn bộ thẻ HTML thô (`<table>`, `<tr>`, `<td>`), chuyển đổi toàn bộ cú pháp Markdown (`**in đậm**`, `*in nghiêng*`) sang Native Word Runs (`run.font.bold`, `run.font.italic`).

---

## 6. Vận hành & Cấu hình Docker

Thư mục dự án: `/home/hieu/Documents/gsb`

### Lệnh Vận hành Cơ bản
```bash
# Khởi động dịch vụ ở chế độ chạy nền
docker compose up -d

# Xem log thời gian thực của service biên bản
docker compose logs -f bidv-minutes

# Xem log thời gian thực của GSB
docker compose logs -f gsb

# Khởi động lại hoặc build lại sau khi cập nhật
docker compose up -d --build

# Dừng toàn bộ cụm
docker compose down
```

### Quản trị Tệp Tin Xuất Bản
Tất cả các tệp `.docx` và `.pdf` sau khi sinh ra qua API sẽ được lưu vĩnh viễn trên máy chủ tại thư mục host:
👉 [`/home/hieu/Documents/gsb/data/outputs`](file:///home/hieu/Documents/gsb/data/outputs)
