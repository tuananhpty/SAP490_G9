const fs = require('fs');
const path = require('path');
const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    HeadingLevel,
    AlignmentType,
    BorderStyle,
    Table,
    TableRow,
    TableCell,
    WidthType,
    ShadingType
} = require('docx');

async function buildDocx() {
    const primaryColor = "003366";   // Dark Navy
    const secondaryColor = "005599"; // Ocean Blue
    const darkText = "222222";       // Soft Black
    const accentBg = "F0F7FF";       // Soft Blue Highlight
    const qaBg = "FFF8E6";           // Soft Gold Highlight
    const techBg = "F4F6F8";         // Soft Gray Code

    const doc = new Document({
        styles: {
            default: {
                document: {
                    run: {
                        font: "Calibri",
                        size: 24, // 12pt
                        color: darkText,
                    },
                    paragraph: {
                        spacing: { line: 360, before: 120, after: 120 }, // 1.5 line spacing
                    },
                },
            },
        },
        sections: [
            {
                properties: {
                    page: {
                        margin: {
                            top: 1440,    // 1 inch
                            bottom: 1440,
                            left: 1440,
                            right: 1440,
                        },
                    },
                },
                children: [
                    // Title Block
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 240, after: 120 },
                        children: [
                            new TextRun({
                                text: "KỊCH BẢN THUYẾT TRÌNH BẢO VỆ KHÓA LUẬN TỐT NGHIỆP",
                                bold: true,
                                size: 36, // 18pt
                                color: primaryColor,
                            }),
                        ],
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 0, after: 360 },
                        children: [
                            new TextRun({
                                text: "Đề tài: Hệ thống Cổng Thông tin Tự phục vụ Nhân sự & Quản lý Kỹ năng (ESS & SkillCert) tích hợp SAP S/4HANA trên nền tảng SAP CAP Framework",
                                italics: true,
                                size: 24,
                                color: "555555",
                            }),
                        ],
                    }),

                    // Info Table
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        shading: { fill: accentBg, type: ShadingType.CLEAR },
                                        borders: {
                                            top: { style: BorderStyle.SINGLE, size: 6, color: primaryColor },
                                            bottom: { style: BorderStyle.SINGLE, size: 6, color: primaryColor },
                                            left: { style: BorderStyle.SINGLE, size: 6, color: primaryColor },
                                            right: { style: BorderStyle.SINGLE, size: 6, color: primaryColor },
                                        },
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({ text: "• Trình độ nói: ", bold: true }),
                                                    new TextRun("Sinh viên đại học trình độ lập trình ở mức khá - trung bình."),
                                                ],
                                            }),
                                            new Paragraph({
                                                children: [
                                                    new TextRun({ text: "• Phong cách: ", bold: true }),
                                                    new TextRun("Tự nhiên, khiêm tốn nhưng tự tin, tập trung vào Luồng Nghiệp vụ & Giải pháp Kiến trúc."),
                                                ],
                                            }),
                                            new Paragraph({
                                                children: [
                                                    new TextRun({ text: "• Thời lượng tổng: ", bold: true }),
                                                    new TextRun("15 - 20 phút (Bao gồm Báo cáo Slide, Live Demo & Phản biện)."),
                                                ],
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),

                    new Paragraph({ spacing: { before: 240, after: 120 } }),

                    // Heading 1: Muc tieu & chien luoc
                    new Paragraph({
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 360, after: 120 },
                        children: [
                            new TextRun({
                                text: "📍 PHẦN 1: MỞ ĐẦU & ĐẶT VẤN ĐỀ (2 - 3 PHÚT)",
                                bold: true,
                                size: 28,
                                color: primaryColor,
                            }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Hướng dẫn thao tác Slide: ", bold: true, italics: true }),
                            new TextRun("Mở Slide 1 (Bìa) -> Chuyển Slide 2 (Đặt vấn đề & Thực trạng doanh nghiệp)."),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "🎙️ Lời thoại mẫu (Phát biểu trực tiếp):", bold: true, color: secondaryColor }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: `"Kính chào Thầy/Cô trong Hội đồng và toàn thể các bạn sinh viên. Em tên là [Tên của bạn]. Sau đây em xin đại diện nhóm trình bày báo cáo khóa luận tốt nghiệp với đề tài: 'Xây dựng Hệ thống Cổng Thông tin Tự phục vụ Nhân sự và Quản lý Kỹ năng (ESS & SkillCert) tích hợp SAP S/4HANA sử dụng SAP CAP Framework'.\n\nThưa Thầy/Cô, trong các doanh nghiệp lớn áp dụng giải pháp ERP của SAP, việc quản lý nhân sự truyền thống đang gặp phải 3 rào cản lớn:\n1. Giao diện SAP tiêu chuẩn (GUI / Fiori tiêu chuẩn) khá phức tạp và khó thao tác đối với nhân viên khối sản xuất/văn phòng không chuyên về SAP.\n2. Áp lực công việc cho bộ phận HR rất lớn khi phải nhập tay thủ công mọi thay đổi nhỏ như số tài khoản ngân hàng, địa chỉ hay đơn xin nghỉ phép vào SAP Infotype.\n3. Nguy cơ nghẽn hệ thống SAP ERP nếu hàng ngàn nhân viên cùng truy cập chấm công hoặc gửi đơn cùng lúc vào giờ cao điểm.\n\nChính vì lý do đó, dự án của em được nghiên cứu và phát triển nhằm mục tiêu cung cấp một Cổng thông tin Web (Web Portal) giao diện Fiori hiện đại, tối ưu cho nhân viên tự phục vụ, đồng thời đóng vai trò là một lớp đệm (Backend-For-Frontend - BFF) xử lý logic và đồng bộ dữ liệu an toàn với hệ thống SAP S/4HANA ở phía sau."`,
                                italics: true,
                            }),
                        ],
                    }),

                    // Section 2: Architecture
                    new Paragraph({
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 360, after: 120 },
                        children: [
                            new TextRun({
                                text: "🏗️ PHẦN 2: KIẾN TRÚC HỆ THỐNG & CÔNG NGHỆ (3 - 4 PHÚT)",
                                bold: true,
                                size: 28,
                                color: primaryColor,
                            }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Sơ đồ luồng dữ liệu kiến trúc (Text Diagram):", bold: true }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: `[ Trình duyệt Web / Client ] (HTML5, Custom Fiori Design, Socket.IO)
            │
            ▼ (HTTPS / OData V4 / REST API)
[ SAP CAP Node.js Server - Lớp BFF Trung Trung ]
            ├── CSDL Cục bộ SQLite (Profile Requests, Locks, Notifications, Snapshots)
            └── SAP Cloud SDK (OData V4/V2 Adapter)
                     │
                     ▼ (Basic Auth / RFC / REST)
[ SAP S/4HANA Backend System ]
            ├── ABAP Staging Tables (Bảng lưu tạm dữ liệu)
            └── ABAP Background Job (Chạy ngầm lúc 05:00 sáng ghi Infotypes)`,
                                font: "Consolas",
                                size: 20,
                            }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "🎙️ Lời thoại giải thích kiến trúc:", bold: true, color: secondaryColor }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: `"Để giải quyết bài toán trên, hệ thống của em được thiết kế theo mô hình 3 lớp chính:\n• Lớp Frontend: Dùng HTML5, CSS và JavaScript thuần kết hợp phong cách thiết kế SAP Fiori Launchpad để tạo cảm giác thân thiện, quen thuộc.\n• Lớp Middle-tier (BFF Service): Sử dụng SAP CAP Framework (Node.js). Lớp này xử lý logic nghiệp vụ, quản lý đăng nhập (Google OAuth 2.0 mapping với Mã nhân viên SAP PERNR), lưu trữ trạng thái đơn từ tạm thời ở cơ sở dữ liệu SQLite cục bộ và gửi thông báo Real-time qua Socket.IO.\n• Lớp Backend ERP: Hệ thống SAP S/4HANA đóng vai trò là kho dữ liệu gốc (Single Source of Truth). Hệ thống Web giao tiếp với SAP thông qua các dịch vụ OData V4/V2 do SAP Cloud SDK hỗ trợ và các chương trình ABAP Background Job chạy ngầm.\n\nLý do em chọn kiến trúc này là để đảm bảo nguyên tắc Decoupling (Tách biệt ứng dụng): Web Portal phản hồi cực nhanh cho người dùng mà không bắt hệ thống SAP Core phải xử lý liên tục từng click chuột."`,
                                italics: true,
                            }),
                        ],
                    }),

                    // Section 3: Core Features
                    new Paragraph({
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 360, after: 120 },
                        children: [
                            new TextRun({
                                text: "⚡ PHẦN 3: CÁC PHÂN HỆ NGHỆVỤ CỐT LÕI (5 - 7 PHÚT)",
                                bold: true,
                                size: 28,
                                color: primaryColor,
                            }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Dự án gồm 5 phân hệ chức năng chính được hoàn thiện đầy đủ end-to-end:",
                            }),
                        ],
                    }),

                    // Feature 1
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        children: [
                            new TextRun({ text: "1. Phân hệ Hồ sơ cá nhân & Quy trình phê duyệt (My Profile Workflow)", bold: true, color: secondaryColor }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "• Mô tả: Cho phép nhân viên nộp yêu cầu cập nhật các thông tin nhạy cảm (Tài khoản ngân hàng, CCCD, Địa chỉ, Người phụ thuộc).\n• Kỹ thuật nổi bật: Sử dụng cơ chế Khóa trường (Field Locking) chống nộp đơn trùng lặp; Mã hóa Hash & Idempotency Key tránh nộp đơn 2 lần do lỗi mạng; Phê duyệt 2 bước trước khi đưa vào SAP Staging.",
                            }),
                        ],
                    }),

                    // Feature 2
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        children: [
                            new TextRun({ text: "2. Phân hệ Đăng ký Chấm công & Nghỉ phép (Attendance Service)", bold: true, color: secondaryColor }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "• Mô tả: Đăng ký nghỉ phép, đi công tác, tăng ca.\n• Kỹ thuật nổi bật: Kết nối OData V4 với service ZUI_NXR_ATTREQ_O4; đồng bộ dữ liệu vào bảng Staging của SAP và chạy chương trình ABAP Job (znxr_attreq_sync_job) lúc 05:00 sáng để tạo bản ghi Infotype PA2002.",
                            }),
                        ],
                    }),

                    // Feature 3
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        children: [
                            new TextRun({ text: "3. Phân hệ Upload Chấm công hàng loạt (HR Worktime Bulk Upload)", bold: true, color: secondaryColor }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "• Mô tả: Dành riêng cho bộ phận HR tải lên tập tin chấm công hàng ngàn dòng (Giờ vào/ra, IOT, IWA).\n• Kỹ thuật nổi bật: Kiểm tra tính hợp lệ dữ liệu (Validate WorkDate, Pernr, Format) trước khi batch upload vào SAP Staging, không ghi trực tiếp PA2002 để tránh quá tải.",
                            }),
                        ],
                    }),

                    // Feature 4 & 5
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        children: [
                            new TextRun({ text: "4. Phân hệ Quản lý Kỹ năng & Chứng chỉ (Skill & Qualification Catalog)", bold: true, color: secondaryColor }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "• Mô tả: Tra cứu danh mục kỹ năng, đăng ký chứng chỉ mới, phê duyệt cấp chứng chỉ cho nhân viên.",
                            }),
                        ],
                    }),
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        children: [
                            new TextRun({ text: "5. Phân hệ Xác thực & Thông báo thời gian thực (Auth & Realtime Notification)", bold: true, color: secondaryColor }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "• Mô tả: Đăng nhập Google OAuth 2.0 kết hợp Passport.js tự động ánh xạ sang mã PERNR SAP qua bảng ProfileIdentityLinks. Tích hợp Socket.IO đẩy thông báo tức thì khi đơn duyệt.",
                            }),
                        ],
                    }),

                    // Technical Highlights
                    new Paragraph({
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 360, after: 120 },
                        children: [
                            new TextRun({
                                text: "🛡️ PHẦN 4: CÁC ĐIỂM KỸ THUẬT ĐÁNG CHÚ Ý & THÁCH THỨC",
                                bold: true,
                                size: 28,
                                color: primaryColor,
                            }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "Trình bày 4 cơ chế kỹ thuật cốt lõi giúp đề tài đạt điểm giỏi chuyên môn:",
                            }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "1. Cơ chế Idempotency Key & Payload Hash: ", bold: true }),
                            new TextRun("Mỗi yêu cầu thay đổi profile tạo ra một UUID duy nhất. Khi người dùng nhấp Submit liên tục, server kiểm tra hash payload và UUID trong bảng ProfileChangeRequests để ngăn chặn trùng lặp bản ghi."),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "2. Cơ chế Field Locking (Khóa trường): ", bold: true }),
                            new TextRun("Sử dụng bảng ProfileFieldLocks. Khi trường 'Tài khoản ngân hàng' đang nằm trong đơn pending, hệ thống ngăn tạo thêm đơn mới chứa trường này cho đến khi đơn cũ hoàn tất xử lý."),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "3. Kiến trúc Staging Table & ABAP 05:00 Sync Job: ", bold: true }),
                            new TextRun("Chuyển toàn bộ thao tác ghi nặng từ thời gian thực sang xử lý lô (Batch processing) ngầm ngoài giờ làm việc thông qua chương trình znxr_attreq_sync_job.prog.abap."),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "4. Identity Linking (Ánh xạ Định danh): ", bold: true }),
                            new TextRun("Gắn kết tài khoản Social Login (Google OAuth) với dữ liệu nhân sự SAP (Pernr) mà không cần thay đổi cấu trúc bảng người dùng của SAP."),
                        ],
                    }),

                    // Live Demo Guide
                    new Paragraph({
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 360, after: 120 },
                        children: [
                            new TextRun({
                                text: "💻 PHẦN 5: KỊCH BẢN THAO TÁC LIVE DEMO CHI TIẾT TỪNG BƯỚC",
                                bold: true,
                                size: 28,
                                color: primaryColor,
                            }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "• Bước 1: ", bold: true }),
                            new TextRun("Mở 2 cửa sổ trình duyệt song song (1 tab Trình duyệt ẩn danh đại diện Nhân viên, 1 tab đại diện Quản lý HR)."),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "• Bước 2: ", bold: true }),
                            new TextRun("Nhân viên đăng nhập Google -> Hệ thống chào đón 'Xin chào [Tên Nhân Viên], Mã PERNR: 00001002'."),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "• Bước 3: ", bold: true }),
                            new TextRun("Nhân viên vào trang Profile -> Sửa 'Số tài khoản ngân hàng' -> Nhấn Submit. Hệ thống thông báo 'Đã gửi đơn thành công. Trường dữ liệu đã được khóa'."),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "• Bước 4: ", bold: true }),
                            new TextRun("Sang tab HR -> Thấy chuông thông báo Socket.IO nhảy đỏ -> HR mở danh sách đơn -> Kiểm tra thông tin cũ/mới -> Nhấn Phê duyệt (Approve)."),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({ text: "• Bước 5: ", bold: true }),
                            new TextRun("Trở lại tab Nhân viên -> Nhận thông báo Real-time đơn đã phê duyệt -> Thông tin Profile tự động cập nhật bản ghi mới."),
                        ],
                    }),

                    // Q&A Section
                    new Paragraph({
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 360, after: 120 },
                        children: [
                            new TextRun({
                                text: "❓ PHẦN 6: BỘ CÂU HỎI PHẢN BIỆN DỰ ĐOÁN & ĐÁP ÁN MẪU TỐT NHẤT",
                                bold: true,
                                size: 28,
                                color: primaryColor,
                            }),
                        ],
                    }),

                    // QA 1
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        children: [
                            new TextRun({ text: "Câu 1: Tại sao không ghi trực tiếp dữ liệu từ Web vào SAP Infotype luôn mà phải qua Staging Table với Background Job 05:00 sáng?", bold: true, color: primaryColor }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: `💡 ĐÁP ÁN MẪU: "Báo cáo Thầy/Cô, ban đầu em cũng cân nhắc gọi BAPI/OData ghi trực tiếp vào SAP. Tuy nhiên khi phân tích nghiệp vụ thực tế, việc ghi thẳng vào các Infotype (như PA2002) đòi hỏi SAP phải thực hiện lock bảng và chạy logic kiểm tra rất nặng. Nếu vào giờ cao điểm có hàng trăm nhân viên cùng nộp đơn hoặc HR upload file hàng ngàn dòng, SAP Core sẽ bị đơ/chậm.\n\nVì vậy em thiết kế luồng đẩy dữ liệu vào Staging Table (bảng tạm) trước. Sau đó chương trình ABAP Job (znxr_attreq_sync_job) chạy tự động vào 05:00 sáng (thời điểm SAP rảnh nhất) để ghi theo lô. Mô hình này vừa giúp ứng dụng Web phản hồi tức thì vừa bảo vệ an toàn cho hệ thống SAP ERP."`,
                                italics: true,
                            }),
                        ],
                    }),

                    // QA 2
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        children: [
                            new TextRun({ text: "Câu 2: Idempotency Key trong dự án hoạt động thế nào?", bold: true, color: primaryColor }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: `💡 ĐÁP ÁN MẪU: "Dạ thưa Thầy/Cô, khi người dùng nhấp Submit trên giao diện Web, nếu đường truyền mạng lag họ có thể bấm liên tục nhiều lần. Nếu không xử lý, hệ thống sẽ tạo ra các bản ghi đơn trùng lặp.\n\nEm giải quyết bằng cách tạo một chuỗi UUID gọi là IdempotencyKey gắn với phiên chỉnh sửa. Khi request gửi lên backend CAP, hệ thống kiểm tra trong bảng ProfileChangeRequests: nếu IdempotencyKey này đã tồn tại, backend trả về ngay kết quả cũ chứ không thực thi lại logic ghi database."`,
                                italics: true,
                            }),
                        ],
                    }),

                    // QA 3
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        children: [
                            new TextRun({ text: "Câu 3: Tại sao ứng dụng lại dùng CSDL SQLite cho lớp CAP backend mà không dùng thẳng SAP HANA?", bold: true, color: primaryColor }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: `💡 ĐÁP ÁN MẪU: "Dạ thưa Thầy/Cô, trong môi trường phát triển (Development/Testing), việc dùng SQLite cục bộ giúp nhóm triển khai nhanh, thử nghiệm độc lập mà không phụ thuộc vào hạ tầng đám mây SAP HANA Cloud đắt đỏ. Khung SAP CAP hỗ trợ cơ chế ngắt kết nối linh hoạt (Database Agnostic), khi đưa lên sản xuất (Production), nhóm chỉ cần đổi file cấu hình package.json sang kind: 'hana' là toàn bộ Schema CDS sẽ được tự động biên dịch và triển khai lên SAP HANA Cloud HDI Container mà không cần sửa code logic."`,
                                italics: true,
                            }),
                        ],
                    }),

                    // QA 4
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        children: [
                            new TextRun({ text: "Câu 4: Cơ chế Khóa trường (Field Locking) ở phân hệ Profile hoạt động ra sao?", bold: true, color: primaryColor }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: `💡 ĐÁP ÁN MẪU: "Dạ, khi nhân viên nộp đơn xin đổi 'Số tài khoản ngân hàng' và đơn ở trạng thái Chờ duyệt (01), bảng ProfileFieldLocks sẽ ghi nhận trường BankKey của nhân viên đó bị Lock bởi Request ID hiện tại. Nếu nhân viên mở tab mới cố tình tạo thêm đơn chứa trường BankKey, hệ thống CAP kiểm tra bảng Lock và chặn ngay lập tức, tránh xung đột dữ liệu đè lên nhau."`,
                                italics: true,
                            }),
                        ],
                    }),

                    // QA 5
                    new Paragraph({
                        heading: HeadingLevel.HEADING_2,
                        children: [
                            new TextRun({ text: "Câu 5: Phần mã nguồn này em tự viết hay tham khảo ở đâu?", bold: true, color: primaryColor }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: `💡 ĐÁP ÁN MẪU: "Dạ báo cáo Thầy/Cô, về khung hệ thống em sử dụng tài liệu chuẩn của SAP CAP và SAP Cloud SDK. Phần định nghĩa OData kết nối SAP em tham khảo các mẫu tích hợp chuẩn. Còn toàn bộ phần logic nghiệp vụ nâng cao như: Field Locking, Idempotency Hash, Staging Workflow, tích hợp Socket.IO và bộ giao diện Fiori trên Express/Node.js là do nhóm em tự thiết kế và lập trình hoàn toàn ạ."`,
                                italics: true,
                            }),
                        ],
                    }),

                    // Conclusion
                    new Paragraph({
                        heading: HeadingLevel.HEADING_1,
                        spacing: { before: 360, after: 120 },
                        children: [
                            new TextRun({
                                text: "🏁 PHẦN 7: KẾT LUẬN & HƯỚNG PHÁT TRIỂN",
                                bold: true,
                                size: 28,
                                color: primaryColor,
                            }),
                        ],
                    }),
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: `"Kính thưa Hội đồng, đề tài của em đã hoàn thành các mục tiêu đề ra về mặt giải pháp kiến trúc và tính năng nghiệp vụ. Do thời gian có hạn, hướng phát triển tiếp theo của nhóm là triển khai ứng dụng lên SAP BTP (Business Technology Platform), tích hợp SAP XSUAA Authentication và chuyển đổi CSDL sang SAP HANA Cloud.\n\nEm xin chân thành cảm ơn Thầy/Cô đã lắng nghe. Em rất mong nhận được những nhận xét và câu hỏi góp ý từ Hội đồng!"`,
                                italics: true,
                            }),
                        ],
                    }),
                ],
            },
        ],
    });

    const outputPath = path.join(process.cwd(), "Kich_Ban_Thuyet_Trinh_Bao_Ve_Khoa_Luan.docx");
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(outputPath, buffer);
    console.log("Successfully created docx at:", outputPath);
}

buildDocx().catch(err => {
    console.error("Error creating docx:", err);
    process.exit(1);
});
