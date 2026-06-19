# 04 Gemini OCR Service

## Pham vi

Prompt Gemini, goi model, parse JSON, chuan hoa bang OCR, doc bang, doc phieu coc, fallback model.

## File duoc sua

Da copy san trong `source/`:

- `gk_pilepro/gk_excel.py`
- `gk_pilepro/gk_overrides.py`
- `gk_pilepro/ui/gk_ocr_ui.py`

## Khong tu y sua

- `presence_server.py`
- `*.spec`, tru khi them dependency OCR moi.
- Asset/logo/icon neu khong lien quan OCR.

## Can phoi hop khi

- Doi nut/luong doc anh: hoi `User Client`.
- Doi cau truc table OCR tra ve: hoi `Excel/Data Files`.
- Doi luu history/log OCR: hoi `Local App Data`.

## Kiem tra

```bat
.\.venv\Scripts\python.exe -m py_compile app.py gk_pilepro\gk_excel.py gk_pilepro\gk_overrides.py
```
