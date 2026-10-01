# Thành phần bên thứ ba

Frontend: React/ReactDOM (MIT), TypeScript (Apache-2.0), Vite (MIT), Office.js từ CDN Microsoft theo điều khoản Microsoft. Backend: Python (PSF), FastAPI/Starlette/Uvicorn/Pydantic và các dependencies; cryptography/cffi, openpyxl, python-docx/lxml, reportlab/Pillow, pypdf. License và metadata được giữ trong runtime hoặc thư mục vendor của bản bàn giao khi có trong bản phân phối.

`tests/microsoft-sideload-template.xlsx` lấy từ [OfficeDev/Office-Addin-Scripts](https://github.com/OfficeDev/Office-Addin-Scripts), giấy phép MIT. Workbook nghiệm thu chỉ thay ID/version của webextension và Title nhằm giới hạn runner vào file kiểm thử riêng.

Không đưa tài khoản, khóa API, tài liệu tài chính hay dữ liệu người dùng vào mã nguồn hoặc bộ ZIP.
