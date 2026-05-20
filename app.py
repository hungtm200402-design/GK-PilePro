# -*- coding: utf-8 -*-
import os
import sys
import hashlib
import uuid
from datetime import datetime
import copy, re, json, traceback, copy, difflib, glob
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Alignment, Protection
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.formula.translate import Translator

APP_TITLE = "GK PilePro"
APP_LOGO_PNG = Path("assets") / "tool_kl_logo.png"
APP_TASKBAR_PNG = Path("assets") / "tool_kl_taskbar.png"
APP_ICON_ICO = Path("assets") / "tool_kl.ico"
TEMPLATE_PRESETS = {"Bảng bất kỳ - tự nhận cột": {}}
CANONICAL_TEMPLATE_COLUMNS = {}

def rounded_icon_image(image, size=(128, 128), radius_ratio=0.18):
    image = image.convert("RGBA").resize(size, Image.Resampling.LANCZOS)
    mask = Image.new("L", size, 0)
    from PIL import ImageDraw
    radius = max(1, int(min(size) * radius_ratio))
    ImageDraw.Draw(mask).rounded_rectangle(
        (0, 0, size[0] - 1, size[1] - 1),
        radius=radius,
        fill=255,
    )
    alpha = image.getchannel("A")
    image.putalpha(Image.composite(alpha, Image.new("L", size, 0), mask))
    return image

UI_BG = "#f4f7fb"
UI_SURFACE = "#ffffff"
UI_SURFACE_2 = "#eef3f8"
UI_BORDER = "#d6dee8"
UI_TEXT = "#162033"
UI_MUTED = "#667085"
UI_PRIMARY = "#1f6feb"
UI_PRIMARY_ACTIVE = "#1557bd"
UI_SUCCESS = "#17803d"
UI_SUCCESS_ACTIVE = "#126931"
UI_WARN = "#b7791f"

class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command, width=15, variant="default"):
        self.command = command
        self.variant = variant
        self.text = text
        self.colors = {
            "default": (UI_SURFACE, UI_TEXT, "#e8eef5", UI_BORDER),
            "primary": (UI_PRIMARY, "#ffffff", UI_PRIMARY_ACTIVE, UI_PRIMARY),
            "success": (UI_SUCCESS, "#ffffff", UI_SUCCESS_ACTIVE, UI_SUCCESS),
            "soft": ("#f8fbff", UI_PRIMARY, "#e8f2ff", "#b8d8ff"),
            "warn": ("#fffaf0", "#c06b00", "#fff0ce", "#ffd38a"),
        }
        self.bg_color, self.fg_color, self.hover_color, self.border_color = self.colors.get(variant, self.colors["default"])
        self.pixel_width = max(94, int(width * 9.2))
        super().__init__(
            parent,
            width=self.pixel_width,
            height=42,
            bg=parent.cget("bg") if hasattr(parent, "cget") else UI_SURFACE,
            bd=0,
            highlightthickness=0,
            cursor="hand2",
        )
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", lambda _e: self._draw(self.hover_color))
        self.bind("<Leave>", lambda _e: self._draw(self.bg_color))
        self._draw(self.bg_color)

    def _draw(self, fill):
        self.delete("all")
        try:
            self.create_round_rect(1, 1, self.pixel_width - 1, 41, radius=9, fill=fill, outline=self.border_color)
        except Exception:
            self.create_rectangle(1, 1, self.pixel_width - 1, 41, fill=fill, outline=self.border_color)
        font = ("Segoe UI", 9, "bold" if self.variant in {"primary", "success"} else "normal")
        self.create_text(self.pixel_width // 2, 21, text=self.text, fill=self.fg_color, font=font)

    def create_round_rect(self, x1, y1, x2, y2, radius=9, **kwargs):
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        return self.create_polygon(points, smooth=True, splinesteps=24, **kwargs)

    def _click(self, _event=None):
        if callable(self.command):
            self.command()

def ui_button(parent, text, command, width=15, variant="default"):
    return RoundedButton(parent, text, command, width=width, variant=variant)

class RoundedMappingLabel(tk.Canvas):
    def __init__(
        self,
        parent,
        text,
        bg_color,
        border_color,
        text_color=UI_TEXT,
        width=118,
        height=30,
        radius=7,
    ):
        super().__init__(
            parent,
            width=width,
            height=height,
            bg=parent.cget("bg") if hasattr(parent, "cget") else UI_SURFACE,
            bd=0,
            highlightthickness=0,
        )
        self.text = text
        self.bg_color = bg_color
        self.border_color = border_color
        self.text_color = text_color
        self.height_px = height
        self.radius = radius
        self.bind("<Configure>", lambda _e: self._draw())
        self._draw()

    def _round_rect(self, x1, y1, x2, y2, radius=7, **kwargs):
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        return self.create_polygon(points, smooth=True, splinesteps=18, **kwargs)

    def _draw(self):
        self.delete("all")
        w = max(20, self.winfo_width())
        self._round_rect(1, 1, w - 1, self.height_px - 1, self.radius, fill=self.bg_color, outline=self.border_color)
        self.create_text(
            9,
            self.height_px // 2,
            text=self.text,
            fill=self.text_color,
            font=("Segoe UI", 8),
            anchor="w",
        )


class RoundedMappingDropdown(tk.Canvas):
    def __init__(
        self,
        parent,
        values,
        variable,
        bg_color,
        border_color,
        width=170,
        height=30,
        radius=7,
    ):
        super().__init__(
            parent,
            width=width,
            height=height,
            bg=parent.cget("bg") if hasattr(parent, "cget") else UI_SURFACE,
            bd=0,
            highlightthickness=0,
            cursor="hand2",
        )
        self.values = list(values or [])
        self.variable = variable or tk.StringVar()
        self.bg_color = bg_color
        self.border_color = border_color
        self.height_px = height
        self.radius = radius
        self._combo_callbacks = []
        self.menu = tk.Menu(self, tearoff=0)
        self._rebuild_menu()
        self.variable.trace_add("write", lambda *_: self._draw())
        self.bind("<Configure>", lambda _e: self._draw())
        self.bind("<Button-1>", self._open_menu)
        self._draw()

    def bind(self, sequence=None, func=None, add=None):
        if sequence == "<<ComboboxSelected>>" and callable(func):
            self._combo_callbacks.append(func)
            return str(len(self._combo_callbacks))
        return super().bind(sequence, func, add)

    def _rebuild_menu(self):
        self.menu.delete(0, "end")
        for value in self.values:
            self.menu.add_command(label=value, command=lambda v=value: self._select(v))

    def set_values(self, values):
        self.values = list(values or [])
        self._rebuild_menu()
        if self.values and self.variable.get() not in self.values:
            self.variable.set(self.values[0])
        elif not self.values:
            self.variable.set("")

    def __setitem__(self, key, value):
        if key == "values":
            self.set_values(value)
        else:
            super().__setitem__(key, value)

    def __getitem__(self, key):
        if key == "values":
            return tuple(self.values)
        return super().__getitem__(key)

    def current(self, index=None):
        if index is None:
            try:
                return self.values.index(self.variable.get())
            except ValueError:
                return -1
        try:
            index = int(index)
            self.variable.set(self.values[index] if 0 <= index < len(self.values) else "")
        except Exception:
            self.variable.set("")

    def get(self):
        return self.variable.get()

    def _round_rect(self, x1, y1, x2, y2, radius=7, **kwargs):
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        return self.create_polygon(points, smooth=True, splinesteps=18, **kwargs)

    def _select(self, value):
        self.variable.set(value)
        for callback in list(self._combo_callbacks):
            try:
                callback(None)
            except Exception:
                pass

    def _open_menu(self, event=None):
        try:
            self.menu.tk_popup(self.winfo_rootx(), self.winfo_rooty() + self.winfo_height())
        finally:
            try:
                self.menu.grab_release()
            except Exception:
                pass

    def _fit_text(self, text, max_width):
        text = str(text or "")
        if max_width <= 20:
            return ""
        font = ("Segoe UI", 8)
        item = self.create_text(0, -100, text=text, font=font, anchor="w")
        if self.bbox(item)[2] <= max_width:
            self.delete(item)
            return text
        self.delete(item)
        suffix = "..."
        out = text
        while out:
            trial = out[:-1].rstrip() + suffix
            item = self.create_text(0, -100, text=trial, font=font, anchor="w")
            width = self.bbox(item)[2]
            self.delete(item)
            if width <= max_width:
                return trial
            out = out[:-1]
        return suffix

    def _draw(self):
        self.delete("all")
        w = max(40, self.winfo_width())
        self._round_rect(1, 1, w - 1, self.height_px - 1, self.radius, fill=self.bg_color, outline=self.border_color)
        arrow_x = w - 16
        self.create_line(arrow_x - 8, 6, arrow_x - 8, self.height_px - 6, fill=self.border_color)
        self.create_polygon(
            arrow_x - 4,
            self.height_px // 2 - 2,
            arrow_x + 4,
            self.height_px // 2 - 2,
            arrow_x,
            self.height_px // 2 + 3,
            fill=UI_MUTED,
            outline=UI_MUTED,
        )
        text = self._fit_text(self.variable.get(), max(10, w - 34))
        self.create_text(10, self.height_px // 2, text=text, fill=UI_TEXT, font=("Segoe UI", 8), anchor="w")

# Các marker dùng để nhận diện dòng TỔNG/TOTAL.
# Giữ dạng không dấu vì hàm norm() đã bỏ dấu tiếng Việt.
TOTAL_MARKERS = {
    "tong", "total", "cong", "sum",
}

SYNONYM_GROUPS = [
    {"stt", "no", "so thu tu", "tt"},
    {"ngay", "date", "ngay thi cong", "ngay ep"},
    {"ten coc", "pile name", "pile no", "ten coc pile name", "ten coc pile no", "pile"},
    {"loai coc", "type of pile", "pile type", "loai coc type of pile"},
    {"vi tri", "location", "vi tri location"},
    {"d1", "đ1", "1st", "to hop coc 1", "to hop 1", "cot to hop 1"},
    {"d2", "đ2", "2nd", "to hop coc 2", "to hop 2", "cot to hop 2"},
    {"d3", "đ3", "3rd", "to hop coc 3", "to hop 3", "cot to hop 3"},
    {"d4", "đ4", "4th", "to hop coc 4", "to hop 4", "cot to hop 4"},
    {"d5", "đ5", "5th", "to hop coc 5", "to hop 5", "cot to hop 5"},
    {"d6", "đ6", "6th", "to hop coc 6", "to hop 6", "cot to hop 6"},
    {"chieu dai to hop", "tong to hop", "total", "total m", "length of pile", "chieu dai coc", "chieu dai to hop m"},
    {"chieu sau ep thuc te", "pressing depth", "thuc te", "reality", "actual depth"},
    {"tai trong dung ep", "luc ep khi dung", "pressing load", "jacking stopping load", "load", "tan"},
    {"bat dau", "start", "thoi gian bat dau", "time start"},
    {"ket thuc", "finish", "end", "thoi gian ket thuc", "time finish"},
    {"dau coc thiet ke", "design pile head", "design"},
    {"mat dat tu nhien", "natural ground", "mat dat tu nhien m"},
    {"dau coc thuc te", "actual pile head", "actual"},
    {"ghi chu", "note", "remark", "remarks"},
]

DEFAULT_MODEL = "gemini-3.1-flash-lite"
FALLBACK_MODELS = ["gemini-3.1-flash-lite", "gemini-2.5-flash-lite", "gemini-2.5-flash"]
APPROVAL_SECRET = "GK_PILEPRO_APPROVAL_V1"

def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def resource_dir():
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return app_dir()

def resource_path(*parts):
    return resource_dir().joinpath(*parts)

def current_user_role_labels():
    if getattr(sys, "frozen", False):
        exe_name = Path(sys.executable).stem.lower()
        if "admin" in exe_name:
            return "Admin", "Quản trị viên"
    return "Thành viên", "Người dùng"

def is_admin_build():
    if not getattr(sys, "frozen", False):
        return False
    return "admin" in Path(sys.executable).stem.lower()

def approval_path():
    return app_dir() / "gk_pilepro_approval.json"

def admin_approved_machines_path():
    return app_dir() / "gk_pilepro_approved_machines.json"

def code_versions_path():
    return app_dir() / "gk_pilepro_code_versions.json"

def legacy_revoked_machines_path():
    return app_dir() / "gk_pilepro_revoked_machines.json"

def get_machine_code():
    raw = f"{os.environ.get('COMPUTERNAME', '')}|{os.environ.get('USERNAME', '')}|{uuid.getnode()}"
    digest = hashlib.sha256(raw.encode("utf-8")).hexdigest().upper()
    return "-".join([digest[i:i + 5] for i in range(0, 20, 5)])

def make_approval_code(machine_code, version=None):
    clean = re.sub(r"[^A-Z0-9]", "", str(machine_code or "").upper())
    if version is None:
        version = machine_approval_version(machine_code)
    if int(version or 1) <= 1:
        digest = hashlib.sha256(f"{APPROVAL_SECRET}:{clean}".encode("utf-8")).hexdigest().upper()
    else:
        digest = hashlib.sha256(f"{APPROVAL_SECRET}:{clean}:V{int(version or 1)}".encode("utf-8")).hexdigest().upper()
    return "-".join([digest[i:i + 4] for i in range(0, 16, 4)])

def load_revoked_machines():
    try:
        data = json.loads(code_versions_path().read_text(encoding="utf-8"))
        if isinstance(data, list):
            return data
    except Exception:
        pass
    try:
        data = json.loads(legacy_revoked_machines_path().read_text(encoding="utf-8"))
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []

def save_revoked_machines(items):
    code_versions_path().write_text(
        json.dumps(items or [], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

def machine_approval_version(machine_code):
    machine_code = str(machine_code or "").strip().upper()
    if not machine_code:
        return 1
    for item in load_revoked_machines():
        if str(item.get("machine_code", "")).strip().upper() == machine_code:
            try:
                return max(1, int(item.get("approval_version") or item.get("version") or 2))
            except Exception:
                return 2
    return 1

def invalidate_machine_approval_code(machine_code):
    machine_code = str(machine_code or "").strip().upper()
    if not machine_code:
        return 1
    items = load_revoked_machines()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for item in items:
        if str(item.get("machine_code", "")).strip().upper() == machine_code:
            try:
                next_version = max(2, int(item.get("approval_version") or item.get("version") or 1) + 1)
            except Exception:
                next_version = 2
            item["approval_version"] = next_version
            item["invalidated_at"] = now
            item.pop("approval_code", None)
            item.pop("revoked_at", None)
            save_revoked_machines(items)
            return next_version
    items.append({
        "machine_code": machine_code,
        "approval_version": 2,
        "invalidated_at": now,
    })
    save_revoked_machines(items)
    return 2

def is_machine_approved():
    machine_code = get_machine_code()
    current_version = machine_approval_version(machine_code)
    approved_list_path = admin_approved_machines_path()
    if approved_list_path.exists():
        approved_items = load_admin_approved_machines()
        if not any(item.get("machine_code") == machine_code for item in approved_items):
            return False
    try:
        data = json.loads(approval_path().read_text(encoding="utf-8"))
        saved_version = int(data.get("approval_version") or 1)
        return (
            data.get("machine_code") == machine_code
            and saved_version == current_version
            and data.get("approval_code") == make_approval_code(machine_code, current_version)
        )
    except Exception:
        return False

def save_machine_approval(approval_code):
    machine_code = get_machine_code()
    current_version = machine_approval_version(machine_code)
    if str(approval_code or "").strip().upper() != make_approval_code(machine_code, current_version):
        return False
    approval_path().write_text(
        json.dumps(
            {
                "machine_code": machine_code,
                "approval_code": make_approval_code(machine_code, current_version),
                "approval_version": current_version,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    try:
        remember_admin_approved_machine(machine_code)
    except Exception:
        pass
    return True

def load_admin_approved_machines():
    try:
        data = json.loads(admin_approved_machines_path().read_text(encoding="utf-8"))
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []

def save_admin_approved_machines(items):
    admin_approved_machines_path().write_text(
        json.dumps(items or [], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

def remember_admin_approved_machine(machine_code):
    machine_code = str(machine_code or "").strip().upper()
    if not machine_code:
        return None
    approval_version = machine_approval_version(machine_code)
    approval_code = make_approval_code(machine_code, approval_version)
    items = load_admin_approved_machines()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    found = False
    for item in items:
        if item.get("machine_code") == machine_code:
            item["approval_code"] = approval_code
            item["approval_version"] = approval_version
            item["approved_at"] = now
            found = True
            break
    if not found:
        items.append({
            "machine_code": machine_code,
            "approval_code": approval_code,
            "approval_version": approval_version,
            "approved_at": now,
        })
    save_admin_approved_machines(items)
    return approval_code

def delete_admin_approved_machine(machine_code):
    machine_code = str(machine_code or "").strip().upper()
    invalidate_machine_approval_code(machine_code)
    items = [x for x in load_admin_approved_machines() if x.get("machine_code") != machine_code]
    save_admin_approved_machines(items)
    try:
        data = json.loads(approval_path().read_text(encoding="utf-8"))
        if str(data.get("machine_code", "")).strip().upper() == machine_code:
            approval_path().unlink(missing_ok=True)
    except Exception:
        pass
    return items

def import_local_approval_to_admin_list():
    try:
        data = json.loads(approval_path().read_text(encoding="utf-8"))
        machine_code = data.get("machine_code", "")
        approval_code = data.get("approval_code", "")
        approval_version = int(data.get("approval_version") or 1)
        if (
            machine_code
            and approval_version == machine_approval_version(machine_code)
            and approval_code == make_approval_code(machine_code, approval_version)
        ):
            remember_admin_approved_machine(machine_code)
    except Exception:
        pass


def env_path():
    external = app_dir() / ".env"
    if external.exists():
        return external
    bundled = resource_dir() / ".env"
    if bundled.exists():
        return bundled
    return external

def user_settings_path():
    return app_dir() / "tool_kl_settings.json"

def load_env_values():
    load_dotenv(env_path())
    values = {
        "GEMINI_API_KEY": os.getenv("GEMINI_API_KEY", ""),
        "GEMINI_MODEL": os.getenv("GEMINI_MODEL", DEFAULT_MODEL),
    }
    try:
        settings = json.loads(user_settings_path().read_text(encoding="utf-8"))
        if settings.get("GEMINI_MODEL"):
            values["GEMINI_MODEL"] = settings["GEMINI_MODEL"]
    except Exception:
        pass
    return values

def save_env(api_key, model):
    if getattr(sys, "frozen", False):
        user_settings_path().write_text(
            json.dumps({"GEMINI_MODEL": (model.strip() or DEFAULT_MODEL)}, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
        return
    env_path().write_text(
        f"GEMINI_API_KEY={api_key.strip()}\nGEMINI_MODEL={(model.strip() or DEFAULT_MODEL)}\n",
        encoding="utf-8"
    )

def clean_text(v):
    s = str(v or "").strip()
    s = s.replace("−","-").replace("–","-").replace("—","-")
    if re.fullmatch(r"-?\d+\.\d+", s):
        s = s.replace(".", ",")
    return s

def norm(s):
    s = str(s or "").strip().lower()
    vietnamese_map = str.maketrans(
        "àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ",
        "aaaaaaaaaaaaaaaaaeeeeeeeeeeeiiiiiooooooooooooooooouuuuuuuuuuuyyyyyd"
    )
    s = s.translate(vietnamese_map)
    s = re.sub(r"[\n\r\t/()\-_:;,\.]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def extract_json(text):
    s = str(text or "").strip()
    s = re.sub(r"^```json\s*", "", s, flags=re.I)
    s = re.sub(r"^```\s*", "", s)
    s = re.sub(r"\s*```$", "", s)
    try:
        return json.loads(s)
    except Exception:
        pass
    m = re.search(r"\{[\s\S]*\}", s)
    if m:
        return json.loads(m.group(0))
    m = re.search(r"\[[\s\S]*\]", s)
    if m:
        return {"tables":[{"title":"","columns":[],"rows":json.loads(m.group(0))}]}
    raise ValueError("Không parse được JSON từ Gemini.")

def build_prompt():
    return """
Bạn là công cụ đọc bảng từ ảnh để nhập Excel. Nhiệm vụ là đọc THẬT CHUẨN dữ liệu trong bảng.

NGUYÊN TẮC BẮT BUỘC:
1. Chỉ đọc dữ liệu nằm TRONG BẢNG chính.
2. Không đọc chữ ngoài bảng, chữ ký, tên người ký, tiêu đề dưới cuối trang.
3. Không tự bịa dữ liệu.
4. Ô nào trống thì trả về "".
5. Ô nào không chắc thì trả về "" hoặc giữ đúng phần đọc được, không đoán.
6. Giữ đúng thứ tự cột từ trái sang phải.
7. Giữ đúng thứ tự dòng từ trên xuống dưới.
8. Không gộp nhiều ô thành một ô.
9. Không tách một ô thành nhiều ô nếu ảnh chỉ có một ô.
10. Không bỏ dòng dữ liệu có STT.
11. Không lấy các dòng trống phía dưới bảng.
12. Không lấy dòng chữ ký, đại diện chủ đầu tư, đại diện đơn vị thi công.

CÁCH ĐỌC BẢNG:
- Đọc theo đường kẻ bảng, không đọc theo chữ rời rạc.
- Xác định header trước, sau đó đọc từng dòng dữ liệu.
- Nếu header nhiều tầng thì ghép tên cột cho rõ nghĩa.
- Nếu một header cha có nhiều cột con thì phải tách từng cột con riêng.

QUY TẮC TỔ HỢP CỌC:
- Nếu ảnh có cột "Tổ hợp cọc" và dưới đó có nhiều cột con:
  + cột con thứ 1 từ trái sang phải = D1
  + cột con thứ 2 = D2
  + cột con thứ 3 = D3
  + cột con thứ 4 = D4
  + cột con thứ 5 = D5
  + cột con thứ 6 = D6
- Ví dụ trong ảnh có "Tổ hợp cọc" gồm 2 ô: 6 | 10
  thì trả về D1 = 6, D2 = 10.
- Không được gộp thành "6 10".
- Nếu ảnh đã ghi sẵn D1/D2 hoặc 1st/2nd thì giữ theo tên đó.

QUY TẮC SỐ LIỆU:
- Giữ dấu phẩy thập phân nếu ảnh dùng dấu phẩy: 14,5; 16,20; -1,20.
- Giữ dấu cộng/trừ: +1,5; -1,20.
- Giữ giờ theo ảnh: 15h00, 15h30, 16h00.
- Giữ ngày theo ảnh: 11/05 hoặc 11/05/2026.
- Không đổi 90 thành 9.0.
- Không đổi D300 thành D30 hoặc 0300.
- Không đổi tên cọc/tim cọc, ví dụ 22, 21, C262, RBH1-C46.

QUY TẮC DỪNG DÒNG:
- Chỉ lấy các dòng có dữ liệu thật.
- Nếu STT từ 1 đến 12 có dữ liệu, còn 13 trở xuống trống thì chỉ trả 12 dòng.
- Không trả các dòng trống 13,14,15...
- Không lấy nét gạch chéo/ký tên làm dữ liệu.

CỘT THƯỜNG GẶP CẦN GIỮ ĐÚNG:
- STT
- Ngày tháng
- Giờ ép hoặc Thời gian bắt đầu/kết thúc
- Tên tim cọc hoặc Tên cọc
- Loại cọc
- D1, D2, D3, D4, D5, D6
- Chiều dài cọc
- Chiều dài ép
- Ép âm dương
- Lực ép
- Ghi chú

BẮT BUỘC SOÁT LỖI THEO Ô:
- Đọc từng ô theo đường kẻ, không suy từ dòng bên trên nếu ô đó trống.
- Dữ liệu nào viết trong ảnh thì giữ y nguyên ở JSON: ngày, giờ, tên cọc, loại cọc, dấu +/-, dấu phẩy, ghi chú.
- Nếu một ô khó đọc, để trống đúng ô đó; không sửa cả dòng, không tự đoán.
- Sau khi đọc xong, tự kiểm tra lại từng cột: STT, ngày, giờ, tên cọc, loại cọc, D1, D2, chiều dài, ép âm dương, lực ép, ghi chú.
- Không tự tính tổng từ ảnh; tổng sẽ do Excel tính lại theo file mẫu.

KIỂM TRA TRƯỚC KHI TRẢ JSON:
- Số ô mỗi dòng phải bằng số cột.
- Nếu thiếu ô thì điền "" để đủ cột.
- Không được lệch cột.
- Dữ liệu ở cột nào phải đúng cột đó.
- Tên cọc không được nhảy sang Loại cọc.
- D1/D2 không được nhảy sang Chiều dài.
- Ghi chú không được nhảy sang Lực ép.


QUY TẮC TỔ HỢP CỌC CỰC KỲ QUAN TRỌNG:
- Nếu header là "Tổ hợp cọc" và trong mỗi dòng có 2 số nằm dưới vùng đó, ví dụ 6 | 10:
  + số thứ nhất phải là D1
  + số thứ hai phải là D2
- Không được trả một cột chung tên "Tổ hợp cọc".
- Không được làm lệch cột sau nó.
- Ví dụ đúng:
  Loại cọc=D300, D1=6, D2=10, Chiều dài cọc=16, Chiều dài ép=14,5, Ép âm dương=+1,5, Lực ép=90, Ghi chú=cắt cọc.

Output JSON thuần, không markdown, không giải thích:
{
  "tables": [
    {
      "title": "tên bảng nếu đọc được, không có thì để rỗng",
      "columns": ["cột 1", "cột 2", "cột 3"],
      "rows": [
        ["ô11", "ô12", "ô13"],
        ["ô21", "ô22", "ô23"]
      ]
    }
  ]
}
"""


def build_prompt_phieu_coc(excel_columns=None):
    """
    Prompt đọc phiếu cọc.
    - Nếu có excel_columns: chỉ trả 1 bảng duy nhất với cột ĐÚNG Y HỆT Excel.
    - Nếu không có: trả cả "Thông tin phiếu" lẫn "Danh sách cọc".
    """
    if excel_columns:
        # Lọc bỏ cột STT/No vì tool tự nối STT
        skip_keys = {"stt", "no", "so thu tu", "tt"}
        cols_to_fill = [c for c in excel_columns if norm(c) not in skip_keys]
        col_list = json.dumps(cols_to_fill, ensure_ascii=False)
        n_cols = len(cols_to_fill)
        example_row = json.dumps(["" for _ in cols_to_fill], ensure_ascii=False)
        # Tạo hướng dẫn đặc biệt cho cột dạng "Độ dài X"
        import re as _re
        do_dai_cols = []
        for c in cols_to_fill:
            m = _re.search(r"(?:d[oô]\s*d[aà]i|chieu\s*dai|length|dai)\s*(\d+[\.,]?\d*)", norm(c))
            if m:
                do_dai_cols.append((c, m.group(1).replace(",", ".")))

        do_dai_hint = ""
        if do_dai_cols:
            do_dai_lines = "\n".join(
                f"  - Cột \"{c}\": số lượng cọc có chiều dài = {dl}m trong phiếu"
                for c, dl in do_dai_cols
            )
            do_dai_hint = f"""
QUY TẮC ĐẶC BIỆT – CỘT ĐỘ DÀI:
Excel có các cột theo từng chiều dài cọc:
{do_dai_lines}

Cách điền:
- Xem trong phiếu có bao nhiêu cọc chiều dài Xm → điền số lượng đó vào cột "Độ dài X".
- Ví dụ: phiếu ghi "D500 - 6m - 10 cây" thì cột "Độ dài 6" = 10, các cột khác = "".
- Nếu phiếu không có cọc độ dài đó thì để "".
- KHÔNG cộng dồn, KHÔNG bịa số lượng.
"""

        col_guide = "\n".join(
            f"  - Cột {i+1}: \"{c}\" → tìm thông tin tương ứng trong phiếu, để \"\" nếu không có"
            for i, c in enumerate(cols_to_fill)
        )
        return f"""
Bạn là công cụ đọc PHIẾU CỌC từ ảnh để điền vào file Excel.

FILE EXCEL CÓ {n_cols} CỘT SAU – bạn PHẢI trả đúng {n_cols} cột này, ĐÚNG THỨ TỰ, ĐÚNG TÊN:
{col_list}

HƯỚNG DẪN TỪNG CỘT:
{col_guide}
{do_dai_hint}
NGUYÊN TẮC:
1. Đọc phiếu cọc trong ảnh (phiếu xuất, nhập, giao cọc).
2. Mỗi DÒNG trong phiếu → 1 row trong JSON.
3. Mỗi row phải có đúng {n_cols} ô tương ứng với {n_cols} cột ở trên theo đúng thứ tự.
4. Ô không có thông tin → "".
5. Không bịa dữ liệu.
6. Giữ nguyên: ngày tháng, số lượng, chiều dài, loại cọc, mã cọc.
7. Không đọc dòng Tổng/Cộng làm dòng dữ liệu.
8. Không lấy chữ ký.

QUY TẮC SỐ:
- Giữ dấu phẩy thập phân: 14,5 giữ là 14,5.
- Giữ ngày như ảnh: 11/05/2026.
- D300 giữ là D300, PHC500 giữ là PHC500.

KIỂM TRA: mỗi row phải có đúng {n_cols} ô.

Output JSON thuần, không markdown:
{{
  "tables": [
    {{
      "title": "Dữ liệu phiếu cọc",
      "columns": {col_list},
      "rows": [
        {example_row}
      ]
    }}
  ]
}}
"""
    else:
        # Không có Excel: trả cả 2 bảng
        return """
Bạn là công cụ đọc PHIẾU CỌC từ ảnh để nhập vào Excel.

NGUYÊN TẮC:
1. Đọc toàn bộ thông tin: phần form chung và bảng danh sách cọc.
2. Không bịa dữ liệu.
3. Giữ ngày tháng, số liệu đúng như ảnh.
4. Không lấy dòng tổng/cộng làm dữ liệu.
5. Không lấy chữ ký.

Cột thường gặp: STT, Ngày xuất, Số HĐ, Loại cọc, Chiều dài (m), Số lượng, Khối lượng, Ghi chú.

Output JSON thuần:
{
  "tables": [
    {
      "title": "Thông tin phiếu",
      "columns": ["Trường", "Giá trị"],
      "rows": [["Số phiếu","..."],["Ngày lập","..."]]
    },
    {
      "title": "Danh sách cọc",
      "columns": ["STT","Ngày xuất","Loại cọc","Chiều dài (m)","Số lượng","Ghi chú"],
      "rows": [["1","11/05/2026","D300","6","10",""]]
    }
  ]
}
"""


def call_gemini_phieu_coc(image_path, api_key, model_name, excel_columns=None):
    """
    Gọi Gemini đọc phiếu cọc từ ảnh.
    Nếu truyền excel_columns, Gemini sẽ trả về dữ liệu khớp với cột Excel.
    """
    from google import genai
    from google.genai import types
    import time

    client = genai.Client(api_key=api_key)
    image = Image.open(image_path)

    preferred = (model_name or DEFAULT_MODEL).strip()
    models_to_try = []
    for m in [preferred] + FALLBACK_MODELS:
        if m and m not in models_to_try:
            models_to_try.append(m)

    last_error = None
    tried = []
    prompt = build_prompt_phieu_coc(excel_columns)

    for model in models_to_try:
        tried.append(model)
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=[prompt, image],
                    config=types.GenerateContentConfig(
                        temperature=0,
                        response_mime_type="application/json",
                    ),
                )

                raw = response.text or ""
                data = extract_json(raw)
                if "tables" not in data:
                    raise ValueError("Gemini không trả về khóa tables.")

                tables = []
                for t in data.get("tables", []):
                    title = clean_text(t.get("title", ""))
                    columns = [clean_text(c) for c in t.get("columns", [])]
                    rows = t.get("rows", [])

                    norm_rows = []
                    if rows and isinstance(rows[0], dict):
                        if not columns:
                            columns = list(rows[0].keys())
                        for r in rows:
                            norm_rows.append([clean_text(r.get(c, "")) for c in columns])
                    else:
                        for r in rows:
                            if isinstance(r, list):
                                rr = [clean_text(x) for x in r]
                                if columns:
                                    rr = rr[:len(columns)] + [""] * max(0, len(columns) - len(rr))
                                norm_rows.append(rr)

                    if columns and norm_rows:
                        tables.append({"title": title, "columns": columns, "rows": norm_rows})

                raw_with_meta = f"MODEL_USED={model}\nTRIED_MODELS={tried}\nLOAI=PHIEU_COC\n\n{raw}"
                return tables, raw_with_meta

            except Exception as e:
                last_error = e
                msg = str(e)
                if "503" in msg or "UNAVAILABLE" in msg or "high demand" in msg:
                    time.sleep(3 + attempt * 2)
                    continue
                if any(x in msg for x in [
                    "429", "RESOURCE_EXHAUSTED", "Quota",
                    "INVALID_ARGUMENT", "API key not valid",
                    "not found", "404"
                ]):
                    break
                raise

    raise RuntimeError(
        "Không gọi được Gemini (phiếu cọc) sau khi thử các model: "
        + ", ".join(tried)
        + "\nLỗi cuối: "
        + repr(last_error)
    )


def call_gemini(image_path, api_key, model_name):
    """
    Gọi Gemini đọc bảng từ ảnh.
    Có retry khi 503 quá tải và fallback model.
    """
    from google import genai
    from google.genai import types
    import time

    client = genai.Client(api_key=api_key)
    image = Image.open(image_path)

    preferred = (model_name or DEFAULT_MODEL).strip()
    models_to_try = []
    for m in [preferred] + FALLBACK_MODELS:
        if m and m not in models_to_try:
            models_to_try.append(m)

    last_error = None
    tried = []

    for model in models_to_try:
        tried.append(model)
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=[build_prompt(), image],
                    config=types.GenerateContentConfig(
                        temperature=0,
                        response_mime_type="application/json",
                    ),
                )

                raw = response.text or ""
                data = extract_json(raw)
                if "tables" not in data:
                    raise ValueError("Gemini không trả về khóa tables.")

                tables = []
                for t in data.get("tables", []):
                    title = clean_text(t.get("title", ""))
                    columns = [clean_text(c) for c in t.get("columns", [])]
                    rows = t.get("rows", [])

                    norm_rows = []
                    if rows and isinstance(rows[0], dict):
                        if not columns:
                            columns = list(rows[0].keys())
                        for r in rows:
                            norm_rows.append([clean_text(r.get(c, "")) for c in columns])
                    else:
                        for r in rows:
                            if isinstance(r, list):
                                rr = [clean_text(x) for x in r]
                                if columns:
                                    rr = rr[:len(columns)] + [""] * max(0, len(columns) - len(rr))
                                norm_rows.append(rr)

                    if columns and norm_rows:
                        tables.append({"title": title, "columns": columns, "rows": norm_rows})

                raw_with_meta = f"MODEL_USED={model}\nTRIED_MODELS={tried}\n\n{raw}"
                return tables, raw_with_meta

            except Exception as e:
                last_error = e
                msg = str(e)

                # Gemini quá tải tạm thời
                if "503" in msg or "UNAVAILABLE" in msg or "high demand" in msg:
                    time.sleep(3 + attempt * 2)
                    continue

                # Quota/rate/model/key: chuyển model hoặc báo lỗi cuối
                if any(x in msg for x in [
                    "429", "RESOURCE_EXHAUSTED", "Quota",
                    "INVALID_ARGUMENT", "API key not valid",
                    "not found", "404"
                ]):
                    break

                raise

    raise RuntimeError(
        "Không gọi được Gemini sau khi thử các model: "
        + ", ".join(tried)
        + "\nLỗi cuối: "
        + repr(last_error)
    )



def copy_style_row(ws, src_row, dst_row, max_col):
    """
    Copy format/style từ dòng mẫu sang dòng mới.
    Không copy dữ liệu cũ.
    """
    for c in range(1, max_col + 1):
        src = ws.cell(src_row, c)
        dst = ws.cell(dst_row, c)

        if dst.__class__.__name__ == "MergedCell":
            continue

        try:
            if src.has_style:
                dst._style = copy.copy(src._style)

            dst.font = copy.copy(src.font)
            dst.fill = copy.copy(src.fill)
            dst.border = copy.copy(src.border)
            dst.alignment = copy.copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy.copy(src.protection)
        except Exception:
            pass

    try:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
        ws.row_dimensions[dst_row].hidden = False
        ws.row_dimensions[dst_row].collapsed = False
    except Exception:
        pass


def copy_row_dimension(ws, src_row, dst_row):
    """
    Copy chiều cao dòng và bỏ ẩn dòng.
    Sửa lỗi dòng mới bị dính/sát nhau do row height nhỏ hoặc hidden.
    """
    try:
        src_dim = ws.row_dimensions[src_row]
        dst_dim = ws.row_dimensions[dst_row]
        dst_dim.height = src_dim.height
        dst_dim.hidden = False
        dst_dim.outlineLevel = src_dim.outlineLevel
        dst_dim.collapsed = False
    except Exception:
        pass


def find_no_col_from_headers(excel_headers):
    for col_idx, name in excel_headers:
        if is_no_header(name):
            return col_idx
    return None

def ensure_no_column_in_mapping(source_cols, mapping, excel_headers):
    """
    Cột No/STT trong Excel không lấy từ ảnh nữa.
    Tool tự nối STT từ số cuối. Vì vậy nguồn STT trong ảnh sẽ bỏ qua.
    """
    no_col = find_no_col_from_headers(excel_headers)
    if not no_col:
        return mapping
    out = list(mapping)
    for i, src in enumerate(source_cols):
        if is_no_header(src):
            out[i] = None
    return out


def find_best_source_for_target(target_name, source_cols):
    tn = norm(target_name)
    best_i = None
    best_score = -1
    for i, src in enumerate(source_cols):
        sn = norm(src)
        score = 0
        if sn == tn:
            score = 100
        elif sn in tn or tn in sn:
            score = 90
        elif group_of(src) == group_of(target_name) and len(group_of(src)) > 1:
            score = 95
        else:
            import difflib
            score = difflib.SequenceMatcher(None, sn, tn).ratio() * 70
        if score > best_score:
            best_score = score
            best_i = i
    return best_i if best_score >= 40 else None

def normalize_table_for_template(table, template_name):
    """
    Nếu chọn preset có khung cố định, ép bảng Gemini về đúng thứ tự cột chuẩn của mẫu đó.
    """
    canonical = CANONICAL_TEMPLATE_COLUMNS.get(template_name)
    if not canonical or not table:
        return table

    src_cols = table.get("columns", [])
    rows = table.get("rows", [])
    idx_map = [find_best_source_for_target(t, src_cols) for t in canonical]

    norm_rows = []
    for row in rows:
        out = []
        for idx in idx_map:
            if idx is None or idx >= len(row):
                out.append("")
            else:
                out.append(row[idx])
        norm_rows.append(out)

    return {
        "title": table.get("title", ""),
        "columns": canonical,
        "rows": norm_rows
    }

def preset_map_columns(source_cols, excel_headers, preset_name):
    preset = TEMPLATE_PRESETS.get(preset_name, {})
    if not preset:
        return auto_map_columns(source_cols, excel_headers)

    excel_norm_by_idx = [(i, col_idx, norm(name), name) for i, (col_idx, name) in enumerate(excel_headers)]
    result = []

    for src in source_cols:
        srcn = norm(src)
        targets = None
        for key, target_list in preset.items():
            kn = norm(key)
            if srcn == kn or srcn in kn or kn in srcn:
                targets = target_list
                break

        if targets == []:
            result.append(None)
            continue

        best = None
        best_score = -1
        if targets:
            target_norms = [norm(t) for t in targets]
            for i, col_idx, exn, name in excel_norm_by_idx:
                score = 0
                for tn in target_norms:
                    if exn == tn:
                        score = max(score, 100)
                    elif tn in exn or exn in tn:
                        score = max(score, 85)
                    else:
                        score = max(score, difflib.SequenceMatcher(None, tn, exn).ratio() * 70)
                if score > best_score:
                    best_score = score
                    best = i
            result.append(best if best_score >= 45 else None)
        else:
            result.append(None)

    return result

def is_total_marker_text(v):
    s = norm(v)
    if not s:
        return False
    return any(m in s for m in TOTAL_MARKERS)

def find_total_row(ws, header_row):
    """
    Tìm dòng TỔNG thật, tránh nhầm header như 'Tổng tổ hợp'.
    Nhận:
    - Dòng có chữ TỔNG/TOTAL đứng riêng ở vài ô đầu.
    - Dòng tổng kết dạng "KL CỌC NHẬP VỀ", "CỘNG", "TỔNG CỘNG" ngoài header.
    Không nhận các header cột như 'Tổng tổ hợp', 'Tổng số m cọc...'
    """
    start_row = max(1, header_row + 1)
    for r in range(start_row, ws.max_row + 1):
        vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, ws.max_column + 1)]
        non_empty = [v for v in vals if v]
        if not non_empty:
            continue

        first_cells = " ".join(vals[:5])
        first_norm = norm(first_cells)
        row_norm = norm(" ".join(vals))

        # TỔNG thật thường nằm ở vài cột đầu và là chữ riêng
        if re.search(r"(^|\s)(tong|total)(\s|$)", first_norm):
            # loại các header dạng "tong to hop", "tong so m coc" nằm trong header
            if any(x in first_norm for x in ["tong to hop", "tong so m", "tong so met"]):
                continue
            return r

        # Nhận dạng dòng tổng kết đặc thù: "KL CỌC NHẬP VỀ", "CỘNG"
        kl_markers = ["kl coc", "kl cọc", "khoi luong coc", "cong tong", "tong cong"]
        if any(k in row_norm for k in kl_markers):
            return r

        # fallback: dòng ít ô, có tổng/total rõ ràng
        if len(non_empty) <= max(4, ws.max_column // 5):
            if re.search(r"(^|\s)(tong|total)(\s|$)", row_norm):
                if not any(x in row_norm for x in ["tong to hop", "tong so m", "tong so met"]):
                    return r

    return None

def find_insert_row(ws, header_row, header_cols):
    total_row = find_total_row(ws, header_row)
    if total_row:
        return total_row
    # Fallback: tìm dòng trống đầu tiên sau header
    for r in range(header_row + 1, ws.max_row + 2):
        vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, min(ws.max_column, 10) + 1)]
        if not any(vals):
            return r
    return ws.max_row + 1

def last_number_above(ws, col_idx, before_row, header_row):
    """
    Lấy STT cuối cùng trong vùng trắng phía trên dòng chèn, không lấy vùng xám/chữ ký.
    """
    last = 0
    if not col_idx:
        return last
    for r in range(header_row + 1, before_row):
        if row_has_grey_background(ws, r):
            continue
        v = ws.cell(r, col_idx).value
        try:
            if str(v).strip().isdigit():
                last = int(str(v).strip())
        except Exception:
            pass
    return last

def _is_grey_fill(cell):
    try:
        if not cell.fill or not cell.fill.fill_type:
            return False
        
        # Lấy mã màu RGB
        fg = str(cell.fill.fgColor.rgb or "").upper()
        # openpyxl đôi khi trả về 00000000 hoặc 00FFFFFF cho ô không màu/trắng
        if not fg or fg in {"00000000", "FFFFFFFF", "FF000000", "00FFFFFF"}:
            return False

        # Các mã màu xám phổ biến
        if fg in {"FF808080", "FFC0C0C0", "FFBFBFBF", "FF999999", "FFEAEAEA", "FFF2F2F2"}:
            return True
            
        # Nếu có màu nền (không phải trắng/trong suốt) thì kiểm tra xem có phải vàng/xanh không
        # Ưu tiên coi là vùng non-data (xám) nếu không phải các màu highlight dữ liệu quen thuộc
        if fg not in {"FF00FF00", "FFFF00", "FFFFFF00", "FFCCFFCC"}:
            return True
    except Exception:
        pass
    return False

def row_has_grey_background(ws, r, max_col=None):
    max_col = max_col or ws.max_column
    grey_count = 0
    checked = 0
    for c in range(1, max_col + 1):
        cell = ws.cell(r, c)
        if cell.__class__.__name__ == "MergedCell":
            continue
        checked += 1
        if _is_grey_fill(cell):
            grey_count += 1
    return checked > 0 and grey_count >= max(2, checked // 3)

def row_is_mostly_blank(ws, r, key_cols):
    vals = [str(ws.cell(r, c).value or "").strip() for c in key_cols]
    return not any(vals)

def find_last_data_row_before(ws, before_row, header_row, key_cols):
    """
    Tìm dòng dữ liệu thật ngay trước dòng TỔNG để copy style.
    Bỏ qua dòng trống/merged/tổng phụ/vùng xám chữ ký.
    """
    for r in range(before_row - 1, header_row, -1):
        vals = [str(ws.cell(r, c).value or "").strip() for c in key_cols]
        text = " ".join(vals)
        if any(vals) and not is_total_marker_text(text) and not row_has_grey_background(ws, r):
            return r
    return max(header_row + 1, before_row - 1)

def find_insert_row_in_white_area(ws, header_row, key_cols):
    """
    Tìm vị trí chèn đúng vùng bảng trắng:
    - Ưu tiên dòng TỔNG/TOTAL.
    - Nếu phía trên TỔNG có dòng trống trong vùng trắng thì điền vào dòng trống đó.
    - Không điền xuống vùng xám/chữ ký.
    """
    total_row = find_total_row(ws, header_row)
    if not total_row:
        # Tìm dòng trống đầu tiên sau header
        for r in range(header_row + 1, ws.max_row + 2):
            if row_is_mostly_blank(ws, r, key_cols):
                return r, None, False
        return ws.max_row + 1, None, False

    # Tìm dòng trống gần nhất phía trên dòng TỔNG trong vùng trắng
    for r in range(total_row - 1, header_row, -1):
        if row_has_grey_background(ws, r):
            continue
        if row_is_mostly_blank(ws, r, key_cols):
            return r, total_row, False

    # Không có dòng trống thì chèn ngay trên TỔNG
    return total_row, total_row, True

def is_no_header(name):
    n = norm(name)
    # Chỉ nhận No/STT thật, tránh nhầm Note
    if n in {"no", "stt", "so thu tu", "tt", "stt no", "no stt", "stt no", "no.", "số tt", "số thứ tự"}:
        return True
    if n.startswith("stt ") or n.endswith(" stt") or "stt" == n:
        return True
    if n.startswith("no ") or n.endswith(" no") or "no" == n:
        return True
    return False

def is_row_total_header(name):
    n = norm(name)
    return any(x in n for x in ["chieu dai to hop", "total m", "total", "length of pile"])

def is_summary_sum_header(name):
    n = norm(name)
    return any(x in n for x in [
        "chieu dai to hop", "total m", "total",
        "chieu sau ep thuc te", "pressing depth",
        "chieu dai coc", "length of pile"
    ])

def is_segment_header(name):
    n = norm(name)
    return n in {"d1", "d2", "d3", "d4", "d5", "d6", "1st", "2nd", "3rd", "4th", "5th", "6th", "đ1", "đ2", "đ3", "đ4", "đ5", "đ6"}

def find_first_data_row(ws, header_row, no_col, total_row):
    """
    Tìm dòng dữ liệu đầu tiên để tính tổng.
    Ưu tiên cột STT/No có số, bỏ qua header/blank.
    """
    if no_col:
        for r in range(header_row + 1, total_row):
            v = ws.cell(r, no_col).value
            try:
                if str(v).strip().isdigit():
                    return r
            except Exception:
                pass
    return header_row + 1

def excel_col_letter(col_idx):
    return get_column_letter(col_idx)


def is_formula_value(v):
    return isinstance(v, str) and v.startswith("=")

def translate_formula_to_row(formula, from_cell, to_cell):
    try:
        return Translator(formula, origin=from_cell).translate_formula(to_cell)
    except Exception:
        return formula

def capture_formula_columns(ws, row_idx):
    """
    Lấy tất cả cột có công thức ở một dòng.
    Dùng để biết cột nào cần tự sum/tự công thức, không hard-code.
    """
    cols = []
    if not row_idx:
        return cols
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row_idx, c).value
        if is_formula_value(v):
            cols.append(c)
    return cols

def capture_total_sum_columns(ws, total_row, first_data_row=None, last_data_row=None):
    """
    Dòng TỔNG cần SUM cột nào:
    - Nếu ô dòng TỔNG có công thức: lấy.
    - Nếu ô dòng TỔNG có số và cột đó có số ở vùng dữ liệu: cũng lấy.
    """
    cols = []
    if not total_row:
        return cols
    for c in range(1, ws.max_column + 1):
        v = ws.cell(total_row, c).value
        if is_formula_value(v):
            cols.append(c)
            continue
        sv = str(v or "").strip().replace(".", "").replace(",", ".")
        try:
            if sv != "":
                float(sv)
                found_num = False
                if first_data_row and last_data_row and last_data_row >= first_data_row:
                    for r in range(first_data_row, last_data_row + 1):
                        vv = str(ws.cell(r, c).value or "").strip().replace(".", "").replace(",", ".")
                        try:
                            float(vv)
                            found_num = True
                            break
                        except Exception:
                            pass
                else:
                    found_num = True
                if found_num:
                    cols.append(c)
        except Exception:
            pass
    return sorted(set(cols))

def apply_row_formulas_from_template(ws, template_row, dst_row):
    """
    Dòng mới cần công thức gì thì lấy theo dòng mẫu phía trên.
    Ví dụ:
    A16 = A15+1  -> A17 = A16+1
    L16 = SUM(F16:K16) -> L17 = SUM(F17:K17)
    """
    for c in range(1, ws.max_column + 1):
        src = ws.cell(template_row, c)
        dst = ws.cell(dst_row, c)
        if dst.__class__.__name__ == "MergedCell":
            continue
        if is_formula_value(src.value):
            dst.value = translate_formula_to_row(
                src.value,
                f"{excel_col_letter(c)}{template_row}",
                f"{excel_col_letter(c)}{dst_row}"
            )

def set_total_formulas_by_template(ws, total_row, formula_cols, first_data_row, last_data_row):
    """
    Dòng TỔNG cần sum cột nào thì dựa vào chính file mẫu:
    cột nào ở dòng TỔNG đang có công thức thì đặt SUM lại cho cột đó.
    Không hard-code chỉ 1-2 cột.
    """
    if not total_row or last_data_row < first_data_row:
        return
    for c in formula_cols:
        cell = ws.cell(total_row, c)
        if cell.__class__.__name__ == "MergedCell":
            continue
        col = excel_col_letter(c)
        cell.value = f"=SUM({col}{first_data_row}:{col}{last_data_row})"



def group_of(s):
    ns = norm(s)
    groups = globals().get("SYNONYM_GROUPS", [])
    for g in groups:
        gn = {norm(x) for x in g}
        if ns in gn:
            return gn
    return {ns}

def auto_map_columns(source_cols, excel_headers):
    """
    Auto map động từ cột ảnh -> cột Excel.
    Không dùng mẫu cố định.
    Mỗi cột Excel chỉ nhận 1 cột ảnh để tránh đè dữ liệu.
    """
    result = [None] * len(source_cols)
    used_excel_idx = set()

    excel_names = [h for _, h in excel_headers]
    excel_norms = [norm(h) for h in excel_names]
    candidates = []

    for s_idx, src in enumerate(source_cols):
        srcn = norm(src)
        src_group = group_of(src)

        if is_no_header(src):
            continue

        for e_idx, ex in enumerate(excel_names):
            exn = excel_norms[e_idx]
            ex_group = group_of(ex)
            score = 0.0
            if exn == srcn:
                score = 100
            elif src_group == ex_group and len(src_group) > 1:
                score = 95
            elif srcn in exn or exn in srcn:
                score = 85
            else:
                score = difflib.SequenceMatcher(None, srcn, exn).ratio() * 70

            if srcn in {"d1","d2","d3","d4","d5","d6"} and exn == srcn:
                score = 110
            if srcn in {"bat dau", "gio bat dau", "start"} and ("bat dau" in exn or exn == "start"):
                score = max(score, 105)
            if srcn in {"ket thuc", "gio ket thuc", "end"} and ("ket thuc" in exn or exn == "end"):
                score = max(score, 105)
            candidates.append((score, s_idx, e_idx))

    candidates.sort(reverse=True, key=lambda x: x[0])
    for score, s_idx, e_idx in candidates:
        if score < 45:
            continue
        if result[s_idx] is not None:
            continue
        if e_idx in used_excel_idx:
            continue
        result[s_idx] = e_idx
        used_excel_idx.add(e_idx)

    return result


def find_last_stt_in_white_area(ws, no_col, header_row, total_row):
    """
    Lấy STT cuối cùng trong vùng trắng phía trên dòng TỔNG.
    Không lấy vùng xám/chữ ký.
    """
    if not no_col or not total_row:
        return 0
    last = 0
    for r in range(header_row + 1, total_row):
        if row_has_grey_background(ws, r):
            continue
        v = ws.cell(r, no_col).value
        try:
            if str(v).strip().isdigit():
                last = int(str(v).strip())
        except Exception:
            pass
    return last

def find_first_data_row_for_sum(ws, no_col, header_row, total_row):
    """
    Dòng đầu để SUM: dòng đầu trong vùng trắng có STT số.
    """
    if no_col:
        for r in range(header_row + 1, total_row):
            if row_has_grey_background(ws, r):
                continue
            v = ws.cell(r, no_col).value
            try:
                if str(v).strip().isdigit():
                    return r
            except Exception:
                pass
    return header_row + 1





def get_stt_value(ws, row, col, memo=None, depth=0):
    """
    Đọc STT thật trong cột STT.
    Hỗ trợ:
    - Số trực tiếp: 27
    - Chữ kèm số: "1.", "01", "(1)"
    - Công thức đơn giản: =A16+1, =A16 + 1, =+A16+1
    """
    if memo is None:
        memo = {}
    key = (row, col)
    if key in memo:
        return memo[key]
    if depth > 50:
        return None

    v = ws.cell(row, col).value
    if v is None:
        return None
        
    if isinstance(v, (int, float)):
        memo[key] = int(v)
        return memo[key]

    s = str(v).strip()
    if not s:
        return None

    # Thử parse số trực tiếp
    if s.isdigit():
        memo[key] = int(s)
        return memo[key]

    # Dạng số có dấu chấm/ngoặc: "1.", "(1)", "1/"
    m_num = re.search(r"^\(?(\d+)\)?[\./]?$", s)
    if m_num:
        memo[key] = int(m_num.group(1))
        return memo[key]

    # Dạng số float nhưng hiển thị nguyên
    try:
        f_val = float(s.replace(",", "."))
        if f_val.is_integer():
            memo[key] = int(f_val)
            return memo[key]
    except Exception:
        pass

    # Dạng công thức: =A16+1 hoặc =+A16+1
    if s.startswith("="):
        f = s.replace(" ", "")
        if f.startswith("=+"):
            f = "=" + f[2:]
        
        # Regex cho =A16+1
        m = re.fullmatch(r"=([A-Z]+)(\d+)\+(\d+)", f, flags=re.I)
        if m:
            ref_col_letters = m.group(1).upper()
            ref_row = int(m.group(2))
            plus = int(m.group(3))
            try:
                ref_col = coordinate_to_tuple(ref_col_letters + "1")[1]
            except Exception:
                ref_col = col
            base = get_stt_value(ws, ref_row, ref_col, memo, depth + 1)
            if base is not None:
                memo[key] = base + plus
                return memo[key]

    return None


def find_all_stt_chains(ws, no_col, header_row, total_row):
    """
    Tìm tất cả chuỗi STT liên tục trước dòng TỔNG.
    Đọc được cả STT là công thức =A16+1.
    Hỗ trợ gộp các chuỗi bị đứt đoạn nhẹ (thiếu 1-2 số).
    """
    if not no_col or not total_row:
        return []

    seq = []
    memo = {}
    for r in range(header_row + 1, total_row):
        try:
            if row_has_grey_background(ws, r):
                continue
        except Exception:
            pass

        n = get_stt_value(ws, r, no_col, memo)
        if isinstance(n, int):
            seq.append((r, n))

    if not seq:
        return []

    # Gom nhóm các số liên tục hoặc gần liên tục
    chains = []
    cur = [seq[0]]
    for item in seq[1:]:
        # Nếu là số tiếp theo (+1) hoặc nhảy nhẹ (+2) nhưng dòng cũng gần nhau
        val_gap = item[1] - cur[-1][1]
        row_gap = item[0] - cur[-1][0]
        
        if val_gap == 1 or (0 < val_gap <= 3 and 0 < row_gap <= 3):
            cur.append(item)
        else:
            if cur:
                chains.append(cur)
            cur = [item]
    if cur:
        chains.append(cur)
        
    return chains


def select_longest_stt_chain(ws, no_col, header_row, total_row):
    """
    Chọn chuỗi STT chuẩn:
    - Ưu tiên chuỗi dài nhất.
    - Nếu bằng nhau, ưu tiên chuỗi bắt đầu nhỏ hơn.
    - Nếu vẫn bằng, ưu tiên STT cuối lớn hơn.
    """
    chains = find_all_stt_chains(ws, no_col, header_row, total_row)
    if not chains:
        return None
    return sorted(chains, key=lambda ch: (len(ch), -ch[0][1], ch[-1][1]), reverse=True)[0]

def find_stt_sequence_region(ws, no_col, header_row, total_row):
    """
    Lấy chuỗi STT liên tục dài nhất trước dòng TỔNG.
    Không lấy dòng rác 19,20,21 dưới bảng.
    """
    best = select_longest_stt_chain(ws, no_col, header_row, total_row)
    if not best:
        return None, None, 0
    return best[0][0], best[-1][0], best[-1][1]

def get_row_values_nonempty(ws, row, cols):
    return [str(ws.cell(row, c).value or "").strip() for c in cols]



def find_no_column_smart(ws, excel_headers, header_row, total_row):
    """
    Tìm cột STT/No chắc hơn:
    1. Theo tên header đã đọc.
    2. Quét vùng header 3 dòng đầu xem có chữ STT/No.
    3. Nếu header đọc sai, tự chọn cột có chuỗi số liên tục dài nhất trước dòng TỔNG.
    """
    # 1. Theo header
    for col_idx, name in excel_headers:
        if is_no_header(name):
            return col_idx

    # 2. Quét trực tiếp vài dòng header
    max_r = min(ws.max_row, header_row + 3)
    for c in range(1, min(ws.max_column, 8) + 1):
        texts = []
        for r in range(max(1, header_row - 1), max_r + 1):
            texts.append(str(ws.cell(r, c).value or ""))
        joined = norm(" ".join(texts))
        if is_no_header(joined):
            return c
        # cell A đôi khi là "STT" ở dòng trên và "No" ở dòng dưới
        if "stt" in joined and ("no" in joined or c == 1):
            return c

    # 3. Fallback: cột có chuỗi số liên tục dài nhất trước TỔNG, ưu tiên cột bên trái
    best_col = None
    best_len = 0
    best_last = 0
    scan_max_col = min(ws.max_column, 8)
    for c in range(1, scan_max_col + 1):
        chains = find_all_stt_chains(ws, c, header_row, total_row)
        if not chains:
            continue
        best_chain = sorted(chains, key=lambda ch: (len(ch), ch[-1][1]), reverse=True)[0]
        ln = len(best_chain)
        last = best_chain[-1][1]
        if ln > best_len or (ln == best_len and best_col is not None and c < best_col):
            best_col = c
            best_len = ln
            best_last = last

    if best_col and best_len >= 3:
        return best_col

    return None



def convert_excel_value(value):
    """
    Chuyển dữ liệu OCR dạng số thành số thật để công thức SUM tính được.
    Giữ nguyên ngày, giờ, mã cọc, loại cọc, vị trí.
    """
    if value is None:
        return ""
    s = str(value).strip()
    if s == "":
        return ""

    # giữ nguyên ngày, giờ, mã có chữ
    lower = s.lower()
    if "/" in s or "h" in lower:
        return s
    if any(ch.isalpha() for ch in s):
        return s

    # bỏ khoảng trắng, đổi dấu phẩy thập phân sang dấu chấm
    t = s.replace(" ", "").replace(",", ".")
    # bỏ dấu + dư
    if t.startswith("+"):
        t = t[1:]

    try:
        if t.count(".") <= 1:
            num = float(t)
            if num.is_integer():
                return int(num)
            return num
    except Exception:
        pass

    return s

def force_workbook_recalculate(wb):
    """
    Bắt Excel/WPS tính lại công thức khi mở file.
    """
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
    try:
        wb.calculation_properties.fullCalcOnLoad = True
        wb.calculation_properties.forceFullCalc = True
        wb.calculation_properties.calcMode = "auto"
    except Exception:
        pass



def find_header_row_smart(ws):
    """
    Tìm header thật cho file Excel:
    - Bỏ qua dòng title merge (nhiều cột cùng 1 giá trị như 'BẢNG THỐNG KÊ...').
    - Ưu tiên dòng có nhiều cột riêng biệt: STT, Ngày, Tên cọc, Loại cọc...
    - Không chọn dòng dữ liệu số.
    """
    best_row = 1
    best_score = -999
    max_r = min(ws.max_row, 40)
    max_c = min(ws.max_column, 30)

    for r in range(1, max_r + 1):
        # Lấy giá trị từng ô riêng biệt trong dòng này
        cell_vals = []
        for c in range(1, max_c + 1):
            v = str(ws.cell(r, c).value or "").strip()
            cell_vals.append(v)

        non_empty = [v for v in cell_vals if v]
        if not non_empty:
            continue

        # --- PHẠT NẶNG dòng title merge ---
        # Nếu tất cả ô không rỗng đều có cùng 1 giá trị → dòng merge tiêu đề
        unique_vals = set(non_empty)
        if len(unique_vals) == 1 and len(non_empty) >= 3:
            continue  # bỏ hoàn toàn dòng này

        # Nếu > 70% ô có cùng giá trị → khả năng cao là merge title
        most_common_count = max(non_empty.count(v) for v in unique_vals)
        merge_ratio = most_common_count / len(non_empty) if non_empty else 0
        if merge_ratio > 0.7 and len(non_empty) >= 3:
            continue  # bỏ dòng merge title

        # Tính điểm dựa trên các keyword cột thực
        text = " ".join(cell_vals)
        # Cũng lấy thêm vài dòng tiếp theo để hỗ trợ header đa tầng
        extra = " ".join(
            str(ws.cell(rr, c).value or "")
            for rr in range(r + 1, min(r + 3, ws.max_row) + 1)
            for c in range(1, max_c + 1)
        )
        n = norm(text + " " + extra)
        score = 0

        # Thưởng cho các keyword cột thực
        col_keywords = [
            "stt", "no", "ngay", "date", "ten coc", "pile", "loai coc",
            "vi tri", "d1", "d2", "d3", "d4", "d5", "d6", "tong", "total",
            "so hop dong", "so phieu", "xe van chuyen", "chieu dai", "so luong",
            "khoi luong", "ghi chu", "note", "bat dau", "ket thuc", "luc ep",
            "hop dong", "chung loai", "mui", "ky hieu"
        ]
        for kw in col_keywords:
            if kw in n:
                score += 10

        # Thưởng thêm cho số lượng cột riêng biệt (header đa cột riêng biệt)
        score += min(len(unique_vals), 15) * 3

        # Phạt dòng dữ liệu số nhiều
        nums = len(re.findall(r"\b\d+[,.]?\d*\b", norm(text)))
        if nums > 10:
            score -= 30

        # Phạt dòng có text dài (tiêu đề bảng, không phải header cột)
        if non_empty and max(len(v) for v in non_empty) > 60:
            score -= 20

        if score > best_score:
            best_score = score
            best_row = r

    return best_row



def get_cell_value_with_merge(ws, row, col):
    """
    Lấy giá trị ô, nếu ô nằm trong vùng merge thì lấy ô góc trên-trái.
    """
    v = ws.cell(row, col).value
    if v not in (None, ""):
        return v
    try:
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                return ws.cell(rng.min_row, rng.min_col).value
    except Exception:
        pass
    return v

def detect_header_rows_from_real_cells(ws, header_row, max_depth=4):
    """
    Chỉ lấy header từ chính file Excel, không fallback sang mẫu cố định.
    Header nhiều tầng thì lấy vài dòng liên tiếp có chữ/ô merge.
    """
    rows = [header_row]
    for r in range(header_row + 1, min(ws.max_row, header_row + max_depth) + 1):
        vals = [str(get_cell_value_with_merge(ws, r, c) or "").strip() for c in range(1, ws.max_column + 1)]
        non_empty = [v for v in vals if v]
        if not non_empty:
            continue

        numeric_like = 0
        for v in non_empty:
            vv = v.replace(",", ".").replace("/", "").replace("-", "").replace(":", "").replace("h", "").strip()
            if vv.replace(".", "").isdigit():
                numeric_like += 1

        has_text = any(re.search(r"[A-Za-zÀ-ỹ]", v) for v in non_empty)
        has_header_token = any(norm(v) in {"d1","d2","d3","d4","d5","d6","1st","2nd","3rd","4th","5th","6th"} for v in non_empty)

        if (has_text and numeric_like < max(3, len(non_empty) // 2)) or has_header_token:
            rows.append(r)
        else:
            break

    return sorted(set(rows))


def get_headers_smart(ws, header_row):
    """
    Đọc cột THẬT trong Excel, xử lý header đa tầng merge:
      - Dòng header: merged parent (VD: "Chủng loại PHC")
      - Dòng dưới:   merged sub-label (VD: "D500 - Độ dài")
      - Dòng dưới nữa: số phân biệt (VD: 5, 6, 7, 8, ...)
    → Kết quả: "Độ dài 5", "Độ dài 6"...
    Các cột thường (STT, Ngày xuất...) giữ nguyên tên.
    """
    used = find_used_range(ws)
    if used:
        min_col = used["min_col"]
        max_col = used["max_col"]
    else:
        min_col = 1
        max_col = ws.max_column

    header_rows = detect_header_rows_from_real_cells(ws, header_row)
    last_header_row = max(header_rows)

    # --- Tìm dòng số phân biệt (leaf row): ngay sau last_header_row ---
    # Đây là dòng chứa số nguyên riêng biệt: 5, 6, 7, 8, 9, 10, 11...
    numeric_leaf_row = None
    for r in range(last_header_row + 1, min(ws.max_row, last_header_row + 5) + 1):
        vals = [ws.cell(r, c).value for c in range(min_col, max_col + 1)]
        non_empty = [v for v in vals if v is not None and str(v).strip() != ""]
        if not non_empty:
            continue
        # Đa số là số nguyên/float
        num_count = sum(
            1 for v in non_empty
            if isinstance(v, (int, float)) or
               re.match(r"^\d+(\.\d+)?$", str(v).strip().replace(",", "."))
        )
        
        # Nhận diện đây có phải là dòng dữ liệu thực hay không?
        # Nếu dòng này chứa giá trị ở các cột "định danh" như STT, Ngày, Tên (chữ) thì nó là dòng data, không phải label
        is_data_row = False
        text_count = 0
        for i, v in enumerate(non_empty):
            sv = str(v).strip()
            # Dữ liệu STT, hoặc chữ
            if re.search(r"[a-zA-ZÀ-ỹ]", sv) and len(sv) > 2:
                text_count += 1
        
        # Nếu có chữ hoặc nếu là dòng ngay sát dữ liệu có cấu trúc đầy đủ, bỏ qua
        if text_count >= 2:
            is_data_row = True
            
        # Kiểm tra thêm: cột đầu tiên (thường là STT/Ngày) có giá trị liên tục không?
        # Nếu có, đích thị là data.
        first_vals = [str(vals[c] or "").strip() for c in range(min(3, len(vals)))]
        if any(first_vals):
            is_data_row = True

        if not is_data_row and num_count >= max(2, len(non_empty) * 0.6):
            numeric_leaf_row = r
            break

    # --- Kiểm tra xem cột có nằm trong vùng header merge không ---
    def is_col_in_merged_header(col_idx):
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= last_header_row and rng.max_row >= header_rows[0]:
                if rng.min_col <= col_idx <= rng.max_col and rng.max_col > rng.min_col:
                    return True
        return False

    # --- Xây dựng tên cột ---
    headers = []
    for c in range(min_col, max_col + 1):
        # Đọc tất cả header rows
        parts = []
        for r in header_rows:
            v = str(get_cell_value_with_merge(ws, r, c) or "").strip()
            if v and v not in parts:
                parts.append(v)

        name_from_headers = " / ".join(parts).strip() if parts else ""

        # Nếu có numeric_leaf_row: chỉ dùng nếu cột này nằm trong merged group
        if numeric_leaf_row and is_col_in_merged_header(c):
            leaf_val = ws.cell(numeric_leaf_row, c).value
            if leaf_val is not None and str(leaf_val).strip() != "":
                leaf_str = str(leaf_val).strip()
                # Chỉ dùng leaf nếu là số (chiều dài cọc)
                if re.match(r"^\d+(\.\d+)?$", leaf_str.replace(",", ".")):
                    # Lấy tên label gần nhất (phần cuối) để làm prefix
                    if parts:
                        label = parts[-1]
                        for p in reversed(parts):
                            pn = norm(p)
                            if any(k in pn for k in ["dai", "do dai", "length", "chieu"]):
                                label = p
                                break
                        # Chỉ giữ phần sau dấu "-" hoặc "/" nếu có
                        for sep in [" - ", "/ ", "/"]:
                            if sep in label:
                                label = label.split(sep)[-1].strip()
                    else:
                        label = "Độ dài"
                    name = f"{label} {leaf_str}"
                    headers.append((c, name))
                    continue

        # Cột thường: dùng tên ghép từ header rows
        name = name_from_headers or f"Cột {get_column_letter(c)}"
        headers.append((c, name))

    # --- Giải quyết tên trùng (nếu còn) ---
    seen: dict = {}
    result = []
    for c, name in headers:
        if name not in seen:
            seen[name] = 0
        else:
            seen[name] += 1
            name = f"{name} ({seen[name]})"
        result.append((c, name))

    return result





def choose_best_sheet_profile(profiles):
    """
    Chọn sheet có khả năng là bảng nhập liệu tốt nhất:
    ưu tiên có dòng TỔNG, có cột STT, có chuỗi STT.
    """
    best = None
    best_score = -1
    for p in profiles:
        if p.get("error"):
            continue
        score = 0
        if p.get("total_row"):
            score += 30
        if p.get("stt_col"):
            score += 30
        sc = p.get("selected_chain")
        if sc:
            score += min(40, sc.get("length", 0))
        if len(p.get("headers", [])) >= 6:
            score += 10
        if score > best_score:
            best = p
            best_score = score
    return best



def is_excel_formula(v):
    return isinstance(v, str) and v.startswith("=")

def cell_addr(row, col):
    return f"{get_column_letter(col)}{row}"

def normalize_formula_to_pattern(formula, origin_cell):
    """
    Dùng để gom nhóm công thức giống nhau theo cột.
    Không cần quá phức tạp; chủ yếu giúp hiển thị công thức mẫu.
    """
    try:
        # dịch công thức về hàng 1 cùng cột để so pattern tương đối
        col_letters = re.match(r"([A-Z]+)", origin_cell).group(1)
        return Translator(formula, origin=origin_cell).translate_formula(f"{col_letters}1")
    except Exception:
        return formula

def formula_references(formula):
    """
    Lấy nhanh các tham chiếu ô/vùng trong công thức để mô tả cách làm.
    """
    if not formula:
        return []
    refs = re.findall(r"(?:'[^']+'!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?", str(formula))
    # loại trùng, giữ thứ tự
    out = []
    for r in refs:
        if r not in out:
            out.append(r)
    return out

def read_formula_logic_for_sheet(ws, header_row=None, total_row=None, no_col=None):
    """
    Đọc công thức và cách làm trong sheet:
    - công thức từng dòng dữ liệu
    - công thức dòng TỔNG
    - nhóm công thức theo cột
    - merged cells
    """
    header_row = header_row or find_header_row_smart(ws)
    total_row = total_row or find_total_row(ws, header_row)
    headers = get_headers_smart(ws, header_row)
    header_by_col = {c: name for c, name in headers}
    no_col = no_col or (find_no_column_smart(ws, headers, header_row, total_row) if total_row else None)

    formulas = []
    formula_cols = {}
    max_rows_scan = ws.max_row

    for r in range(1, max_rows_scan + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if is_excel_formula(v):
                addr = cell_addr(r, c)
                item = {
                    "cell": addr,
                    "row": r,
                    "col": c,
                    "col_letter": get_column_letter(c),
                    "header": header_by_col.get(c, ""),
                    "formula": v,
                    "references": formula_references(v),
                    "pattern": normalize_formula_to_pattern(v, addr),
                    "role": "normal"
                }
                if total_row and r == total_row:
                    item["role"] = "total_row"
                elif total_row and header_row < r < total_row:
                    item["role"] = "data_row"
                formulas.append(item)
                formula_cols.setdefault(c, []).append(item)

    # Tìm công thức mẫu cho từng cột
    formula_columns = []
    for c, items in sorted(formula_cols.items()):
        patterns = {}
        for it in items:
            patterns[it["pattern"]] = patterns.get(it["pattern"], 0) + 1
        common_pattern = sorted(patterns.items(), key=lambda x: x[1], reverse=True)[0][0] if patterns else ""
        sample = items[0]
        formula_columns.append({
            "col": get_column_letter(c),
            "col_index": c,
            "header": header_by_col.get(c, ""),
            "count": len(items),
            "sample_cell": sample["cell"],
            "sample_formula": sample["formula"],
            "common_pattern": common_pattern,
            "roles": sorted(set(it["role"] for it in items)),
        })

    total_formulas = [it for it in formulas if it["role"] == "total_row"]
    data_formulas = [it for it in formulas if it["role"] == "data_row"]

    merged_ranges = [str(rng) for rng in ws.merged_cells.ranges]

    # Mô tả ngắn "cách làm"
    rules = []
    for it in total_formulas:
        rules.append(f"TỔNG {it['col_letter']} ({it.get('header','')}): {it['formula']}")
    # lấy tối đa 20 công thức dòng dữ liệu mẫu
    seen_cols = set()
    for it in data_formulas:
        if it["col"] in seen_cols:
            continue
        seen_cols.add(it["col"])
        rules.append(f"Dòng dữ liệu cột {it['col_letter']} ({it.get('header','')}): mẫu {it['sample_cell'] if 'sample_cell' in it else it['cell']} = {it['formula']}")
        if len(seen_cols) >= 20:
            break

    return {
        "sheet": ws.title,
        "header_row": header_row,
        "total_row": total_row,
        "stt_col": get_column_letter(no_col) if no_col else None,
        "formula_count": len(formulas),
        "formula_columns": formula_columns,
        "total_formulas": total_formulas,
        "data_formulas_sample": data_formulas[:80],
        "merged_ranges": merged_ranges[:200],
        "rules_text": rules[:80],
    }

def read_formula_logic_for_workbook(excel_path):
    wb = load_workbook(excel_path, data_only=False)
    result = {
        "file": str(excel_path),
        "sheets": []
    }
    for ws in wb.worksheets:
        try:
            header_row = find_header_row_smart(ws)
            headers = get_headers_smart(ws, header_row)
            total_row = find_total_row(ws, header_row)
            no_col = find_no_column_smart(ws, headers, header_row, total_row) if total_row else None
            logic = read_formula_logic_for_sheet(ws, header_row, total_row, no_col)
            result["sheets"].append(logic)
        except Exception as e:
            result["sheets"].append({"sheet": ws.title, "error": repr(e)})
    return result



def auto_mapping_to_excel_columns(source_cols, excel_headers):
    """
    Trả mapping dạng source_idx -> excel column number.
    Dùng khi người dùng quên bấm Auto map cột.
    """
    auto_idx = auto_map_columns(source_cols, excel_headers)
    auto_idx = ensure_no_column_in_mapping(source_cols, auto_idx, excel_headers)
    out = []
    for idx in auto_idx:
        if idx is None:
            out.append(None)
        else:
            try:
                out.append(excel_headers[idx][0])
            except Exception:
                out.append(None)
    return out



def cell_text(v):
    return str(v or "").strip()

def row_non_empty_count(ws, r):
    return sum(1 for c in range(1, ws.max_column + 1) if cell_text(ws.cell(r, c).value))

def find_used_range(ws):
    min_r, min_c = None, None
    max_r, max_c = 0, 0
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            if cell_text(ws.cell(r, c).value):
                min_r = r if min_r is None else min(min_r, r)
                min_c = c if min_c is None else min(min_c, c)
                max_r = max(max_r, r)
                max_c = max(max_c, c)
    if min_r is None:
        return None
    return {"min_row": min_r, "min_col": min_c, "max_row": max_r, "max_col": max_c}

def infer_sheet_type(text):
    n = norm(text)
    if any(x in n for x in ["ngay xuat", "so hop dong", "so phieu", "xe van chuyen", "tong so m coc", "mui d300", "chung loai phc"]):
        return "Bảng xuất/nhập cọc - hợp đồng/phiếu/xe vận chuyển"
    if any(x in n for x in ["tong hop coc thuc te", "tai trong dung ep", "cao do", "chieu sau ep thuc te", "luc ep"]):
        return "Bảng báo cáo ép cọc / nhật ký ép cọc"
    if any(x in n for x in ["summary construction", "pile press", "pile combination", "pressing load"]):
        return "Bảng tổng hợp khối lượng ép cọc song ngữ"
    if any(x in n for x in ["nghiem thu", "bien ban", "xac nhan"]):
        return "Biên bản/nghiệm thu"
    return "Bảng Excel khác / chưa phân loại chắc"

def build_multiline_headers(ws, header_rows, min_col, max_col):
    headers = []
    for c in range(min_col, max_col + 1):
        parts = []
        for r in header_rows:
            v = cell_text(ws.cell(r, c).value)
            if v and v not in parts:
                parts.append(v)
        name = " / ".join(parts).strip()
        if not name:
            name = f"Cột {get_column_letter(c)}"
        headers.append({"col": get_column_letter(c), "index": c, "name": name})
    return headers

def detect_header_rows_general(ws, used):
    """
    Dò các dòng header của từng sheet, không cố định form.
    Lấy vùng có nhiều chữ/ô gộp trước data.
    """
    if not used:
        return []
    min_r, max_scan = used["min_row"], min(used["max_row"], used["min_row"] + 25)
    best_r = min_r
    best_score = -999

    for r in range(min_r, max_scan + 1):
        vals = [cell_text(ws.cell(r, c).value) for c in range(used["min_col"], used["max_col"] + 1)]
        joined = " ".join(vals)
        n = norm(joined)
        non_empty = sum(1 for v in vals if v)
        alpha = sum(1 for v in vals if re.search(r"[A-Za-zÀ-ỹ]", v))
        nums = sum(1 for v in vals if re.fullmatch(r"[-+]?\d+[,.]?\d*", v.replace(" ", "")))

        score = non_empty * 2 + alpha * 3 - nums * 2
        for kw in ["stt", "ngay", "ten", "loai", "vi tri", "d1", "1st", "tong", "ghi chu",
                   "so hop dong", "so phieu", "xe van chuyen", "chung loai", "tai trong", "cao do"]:
            if kw in n:
                score += 15

        if score > best_score:
            best_score = score
            best_r = r

    # Header có thể nhiều tầng: lấy best_r và 1-2 dòng tiếp theo nếu có chữ, không lấy dòng dữ liệu quá số
    header_rows = [best_r]
    for rr in range(best_r + 1, min(best_r + 4, used["max_row"] + 1)):
        vals = [cell_text(ws.cell(rr, c).value) for c in range(used["min_col"], used["max_col"] + 1)]
        joined = " ".join(vals)
        has_alpha = any(re.search(r"[A-Za-zÀ-ỹ]", v) for v in vals if v)
        numeric_heavy = sum(1 for v in vals if re.fullmatch(r"[-+]?\d+[,.]?\d*", v.replace(" ", ""))) >= max(3, len(vals)//3)
        if has_alpha and not numeric_heavy:
            header_rows.append(rr)
        elif any(norm(v) in {"d1","d2","d3","d4","d5","d6","1st","2nd","3rd","4th","5th"} for v in vals):
            header_rows.append(rr)
    return sorted(set(header_rows))

def detect_data_rows_general(ws, used, header_rows):
    if not used or not header_rows:
        return []
    start = max(header_rows) + 1
    rows = []
    for r in range(start, used["max_row"] + 1):
        vals = [cell_text(ws.cell(r, c).value) for c in range(used["min_col"], used["max_col"] + 1)]
        joined = " ".join(vals)
        if not joined.strip():
            continue
        if is_total_marker_text(joined):
            continue
        # dòng dữ liệu thường có số/ngày/mã hoặc nhiều ô
        if sum(1 for v in vals if v) >= 2:
            rows.append(r)
    return rows

def analyze_sheet_content(ws):
    used = find_used_range(ws)
    if not used:
        return {
            "sheet": ws.title,
            "empty": True,
            "summary": "Sheet trống"
        }

    all_text = []
    for r in range(used["min_row"], used["max_row"] + 1):
        vals = [cell_text(ws.cell(r, c).value) for c in range(used["min_col"], used["max_col"] + 1)]
        if any(vals):
            all_text.append(" ".join(vals))
    joined_all = "\n".join(all_text)

    header_rows = detect_header_rows_general(ws, used)
    headers = build_multiline_headers(ws, header_rows, used["min_col"], used["max_col"]) if header_rows else []

    data_rows = detect_data_rows_general(ws, used, header_rows)
    total_rows = []
    for r in range(used["min_row"], used["max_row"] + 1):
        row_text = " ".join(cell_text(ws.cell(r, c).value) for c in range(used["min_col"], used["max_col"] + 1))
        if is_total_marker_text(row_text):
            total_rows.append(r)

    formula_cells = []
    for r in range(used["min_row"], used["max_row"] + 1):
        for c in range(used["min_col"], used["max_col"] + 1):
            v = ws.cell(r, c).value
            if is_formula_value(v):
                formula_cells.append({
                    "cell": f"{get_column_letter(c)}{r}",
                    "formula": v,
                    "col": get_column_letter(c),
                })

    sample_rows = []
    for r in data_rows[:8]:
        row_data = {}
        for h in headers:
            v = cell_text(ws.cell(r, h["index"]).value)
            if v:
                row_data[h["name"]] = v
        if row_data:
            sample_rows.append({"row": r, "values": row_data})

    merged = [str(x) for x in list(ws.merged_cells.ranges)[:80]]

    return {
        "sheet": ws.title,
        "empty": False,
        "sheet_type": infer_sheet_type(joined_all),
        "used_range": {
            "from": f"{get_column_letter(used['min_col'])}{used['min_row']}",
            "to": f"{get_column_letter(used['max_col'])}{used['max_row']}",
            "min_row": used["min_row"],
            "max_row": used["max_row"],
            "min_col": used["min_col"],
            "max_col": used["max_col"],
        },
        "header_rows": header_rows,
        "headers": headers,
        "data_row_count": len(data_rows),
        "data_rows_first_last": [data_rows[0], data_rows[-1]] if data_rows else None,
        "total_rows": total_rows,
        "formula_count": len(formula_cells),
        "formula_samples": formula_cells[:50],
        "merged_ranges_count": len(ws.merged_cells.ranges),
        "merged_ranges_sample": merged,
        "sample_rows": sample_rows,
        "summary": f"{infer_sheet_type(joined_all)} | {len(headers)} cột | {len(data_rows)} dòng dữ liệu | {len(formula_cells)} ô công thức"
    }

def analyze_workbook_sheets(excel_path):
    wb = load_workbook(excel_path, data_only=False)
    return {
        "file": str(excel_path),
        "sheets": [analyze_sheet_content(ws) for ws in wb.worksheets]
    }



def short_header_name(name, max_len=45):
    """
    Rút gọn tên cột hiển thị:
    - Bỏ tiêu đề bảng dài kiểu BẢNG DIỄN GIẢI...
    - Giữ phần header thật của cột
    - Không ảnh hưởng dữ liệu/mapping gốc
    """
    s = str(name or "").replace("\n", " ").replace("\r", " ")
    parts = [re.sub(r"\s+", " ", p).strip() for p in s.split("/") if str(p).strip()]

    cleaned = []
    for p in parts:
        np = norm(p)
        # bỏ title/slogan dài, không phải tên cột
        if "bang dien giai" in np or "khoi luong thi cong" in np:
            continue
        if "summary construction" in np or "bang tong hop" in np:
            continue
        if len(p) > 55 and ("bang" in np or "cong" in np or "construction" in np):
            continue
        if p not in cleaned:
            cleaned.append(p)

    if not cleaned:
        cleaned = parts[-2:] if len(parts) >= 2 else parts

    # Nếu chỉ có đơn vị M/T/Tim thì giữ kèm cha phía trước
    if len(cleaned) >= 2 and cleaned[-1].lower() in {"m", "t", "tim"}:
        out = f"{cleaned[-2]} ({cleaned[-1]})"
    else:
        out = " / ".join(cleaned[-3:]) if cleaned else str(name or "")

    out = re.sub(r"\s+", " ", out).strip()
    if len(out) > max_len:
        out = out[:max_len-3].rstrip() + "..."
    return out



def find_last_data_row_before_total(ws, header_row, total_row, mapping_cols=None, no_col=None):
    """
    Fallback khi không có chuỗi STT liên tục:
    tìm dòng dữ liệu cuối cùng trước dòng TỔNG dựa vào các cột có dữ liệu/mapping.
    """
    end_row = (total_row - 1) if total_row else ws.max_row
    cols = []
    if mapping_cols:
        cols.extend([c for c in mapping_cols if c])
    if no_col:
        cols.append(no_col)
    cols = sorted(set(cols))
    if not cols:
        cols = list(range(1, ws.max_column + 1))

    last = None
    for r in range(header_row + 1, end_row + 1):
        if row_has_grey_background(ws, r):
            continue
        vals = [str(ws.cell(r, c).value or "").strip() for c in cols if c <= ws.max_column]
        if any(vals):
            # bỏ qua dòng header phụ nếu toàn chữ cột
            row_text = norm(" ".join(vals))
            if row_text and not is_total_marker_text(row_text):
                last = r
    return last

def find_last_stt_number_loose(ws, no_col, header_row, total_row=None):
    """
    Lấy số STT cuối cùng, không yêu cầu liên tục.
    Dùng cho các sheet STT bị đứt hoặc có công thức chưa tính.
    """
    if not no_col:
        return 0
    end_row = (total_row - 1) if total_row else ws.max_row
    last = 0
    memo = {}
    for r in range(header_row + 1, end_row + 1):
        if row_has_grey_background(ws, r):
            continue
        n = None
        try:
            n = get_stt_value(ws, r, no_col, memo)
        except Exception:
            n = None
        if isinstance(n, int):
            last = max(last, n)
            continue
        s = str(ws.cell(r, no_col).value or "").strip()
        if s.isdigit():
            last = max(last, int(s))
    return last




def row_has_big_merge_area(ws, row):
    """
    Tránh ghi vào vùng merge lớn kiểu chữ ký/trang trống.
    """
    try:
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row:
                row_span = rng.max_row - rng.min_row + 1
                col_span = rng.max_col - rng.min_col + 1
                if row_span >= 2 and col_span >= 3:
                    return True
    except Exception:
        pass
    return False


def row_is_empty_for_new_data(ws, row, mapping_cols, no_col):
    """
    Dòng được coi là trống nếu các cột cần ghi dữ liệu đều đang trống.
    Không xét cột STT vì có file đã kẻ sẵn hoặc có công thức STT.
    """
    try:
        row_text = " ".join(str(ws.cell(row, c).value or "") for c in range(1, ws.max_column + 1))
        if is_total_marker_text(row_text):
            return False
    except Exception:
        pass

    if row_has_big_merge_area(ws, row):
        return False

    cols = []
    if mapping_cols:
        cols.extend([c for c in mapping_cols if c])
    cols = sorted(set([c for c in cols if c != no_col]))

    if not cols:
        return False

    for c in cols:
        v = ws.cell(row, c).value
        if v not in (None, ""):
            return False

    return True


def find_blank_rows_before_total(ws, start_after_row, total_row, need_count, mapping_cols, no_col):
    """
    Tìm các dòng trống thật sự nằm trước dòng TỔNG để ghi dữ liệu.
    Ưu tiên giữ nguyên form, không insert/xóa dòng.
    """
    rows = []
    for r in range(start_after_row + 1, total_row):
        if row_is_empty_for_new_data(ws, r, mapping_cols, no_col):
            rows.append(r)
            if len(rows) >= need_count:
                break
    return rows

class MappingEditor(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=UI_SURFACE)
        self.pack(fill="x", pady=(8, 0))

        self.source_bg = "#f3efe7"
        self.target_bg = "#edf6ff"
        self.mapping_border = "#c7d5e8"
        self.combo_style = "Mapping.TCombobox"
        try:
            style = ttk.Style()
            style.configure(
                self.combo_style,
                fieldbackground=self.target_bg,
                background=self.target_bg,
                foreground=UI_TEXT,
                bordercolor=self.mapping_border,
                lightcolor=self.mapping_border,
                darkcolor=self.mapping_border,
                arrowcolor=UI_MUTED,
                padding=2,
            )
            style.map(
                self.combo_style,
                fieldbackground=[("readonly", self.target_bg)],
                background=[("readonly", self.target_bg)],
                selectbackground=[("readonly", "#cfe5ff")],
                selectforeground=[("readonly", UI_TEXT)],
            )
        except Exception:
            pass

        self.mapping_vars = []
        self.excel_headers = []
        self.table_cols = []
        self.display_to_col = {}
        self.col_to_display = {}

        self.title = tk.Label(
            self,
            text="Xác nhận mapping cột",
            font=("Segoe UI", 11, "bold"),
            bg=UI_SURFACE,
            fg=UI_TEXT,
        )
        self.title.pack(anchor="w")

        # Khung cuộn mapping
        self.box = tk.Frame(self, bg=UI_SURFACE)
        self.box.pack(fill="x", expand=False, pady=(6, 0))

        canvas_h = int(getattr(parent, "mapping_canvas_h", 260))
        self.canvas = tk.Canvas(
            self.box,
            height=canvas_h,
            highlightthickness=1,
            highlightbackground=UI_BORDER,
            bg=UI_SURFACE,
            bd=0,
        )
        self.scrollbar = ttk.Scrollbar(self.box, orient="vertical", command=self.canvas.yview)

        self.inner = tk.Frame(self.canvas, bg=UI_SURFACE)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="x", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.inner.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)

    def _on_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event=None):
        try:
            self.canvas.itemconfig(self.inner_id, width=event.width)
        except Exception:
            pass

    def _on_mousewheel(self, event):
        try:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    def _bind_mousewheel(self, event=None):
        try:
            self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        except Exception:
            pass

    def _unbind_mousewheel(self, event=None):
        try:
            self.canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass

    def clear(self):
        for w in self.inner.winfo_children():
            w.destroy()

        self.mapping_vars = []
        self.excel_headers = []
        self.table_cols = []
        self.display_to_col = {}
        self.col_to_display = {}

        try:
            self.canvas.yview_moveto(0)
        except Exception:
            pass

    def _make_excel_choices(self, excel_headers):
        choices = ["(bỏ qua)"]
        used = set()
        self.display_to_col = {}
        self.col_to_display = {}

        for col_idx, name in excel_headers:
            try:
                short = short_header_name(name, 34)
            except Exception:
                short = str(name)

            label = f"{get_column_letter(col_idx)}: {short}"

            base = label
            n = 2
            while label in used:
                label = f"{base} ({n})"
                n += 1

            used.add(label)
            choices.append(label)
            self.display_to_col[label] = col_idx
            self.col_to_display[col_idx] = label

        return choices

    def set_mapping(self, table_cols, excel_headers, auto_map_idx):
        self.clear()
        self.table_cols = table_cols
        self.excel_headers = excel_headers

        excel_choices = self._make_excel_choices(excel_headers)

        for i, src in enumerate(table_cols):
            row = tk.Frame(self.inner, bg=UI_SURFACE)
            row.pack(fill="x", pady=3)

            try:
                src_show = short_header_name(src, 18)
            except Exception:
                src_show = str(src)

            var = tk.StringVar()

            if i >= len(auto_map_idx) or auto_map_idx[i] is None:
                var.set("(bỏ qua)")
            else:
                try:
                    excel_col_idx = excel_headers[auto_map_idx[i]][0]
                    var.set(self.col_to_display.get(excel_col_idx, "(bỏ qua)"))
                except Exception:
                    var.set("(bỏ qua)")

            row.grid_columnconfigure(0, weight=0, minsize=118)
            row.grid_columnconfigure(1, weight=0, minsize=18)
            row.grid_columnconfigure(2, weight=1)

            src_box = RoundedMappingLabel(
                row,
                text=src_show,
                bg_color=self.source_bg,
                border_color="#d8d2c8",
            )
            src_box.grid(row=0, column=0, sticky="ew", padx=(2, 4))
            tk.Label(
                row,
                text="→",
                width=1,
                anchor="center",
                bg=UI_SURFACE,
                fg=UI_MUTED,
            ).grid(row=0, column=1, sticky="w", padx=(0, 3))

            cb = RoundedMappingDropdown(
                row,
                values=excel_choices,
                variable=var,
                bg_color=self.target_bg,
                border_color=self.mapping_border,
            )
            cb.grid(row=0, column=2, sticky="ew", padx=(0, 6))

            self.mapping_vars.append(var)

        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.canvas.yview_moveto(0)

    def get_mapping(self):
        out = []
        for var in self.mapping_vars:
            chosen = var.get()
            if chosen == "(bỏ qua)":
                out.append(None)
            else:
                out.append(self.display_to_col.get(chosen))
        return out




def postprocess_to_hop_coc_d1_d2(tables):
    """
    Fix riêng cho form mới:
    Nếu Gemini đọc header "Tổ hợp cọc" thành 1 cột nhưng dữ liệu thực tế là 2 cột con,
    ví dụ: D300 | 6 | 10 | 16 | 14,5 | +1,5 | 90
    thì đổi thành:
    D300 | D1=6 | D2=10 | Chiều dài cọc=16 | Chiều dài ép=14,5 | Ép âm dương=+1,5 | Lực ép=90

    Không ảnh hưởng các form cũ đã có sẵn D1/D2 hoặc 1st/2nd.
    """
    if not tables:
        return tables

    def _n(s):
        try:
            return norm(s)
        except Exception:
            return str(s or "").lower().strip()

    for t in tables:
        cols = list(t.get("columns", []))
        rows = t.get("rows", [])

        if not cols:
            continue

        norm_cols = [_n(c) for c in cols]

        # Nếu đã có D1/D2 hoặc 1st/2nd rồi thì bỏ qua, giữ cấu trúc cũ
        has_d1 = any(x in {"d1", "đ1", "1st"} for x in norm_cols)
        has_d2 = any(x in {"d2", "đ2", "2nd"} for x in norm_cols)
        if has_d1 and has_d2:
            continue

        # Tìm cột "Tổ hợp cọc"
        idx = None
        for i, nc in enumerate(norm_cols):
            if "to hop coc" in nc or nc == "to hop" or "pile combination" in nc:
                idx = i
                break

        if idx is None:
            continue

        # Nếu ngay sau Tổ hợp cọc là Chiều dài cọc thì khả năng cao Gemini bị thiếu header D2
        next_name = norm_cols[idx + 1] if idx + 1 < len(norm_cols) else ""
        should_split = False

        if "chieu dai coc" in next_name or "length of pile" in next_name:
            should_split = True

        # Hoặc nếu dữ liệu ở cột tổ hợp và cột kế tiếp đều là số ngắn, cũng tách
        if not should_split and idx + 1 < len(cols):
            sample_count = 0
            ok_count = 0
            for r in rows[:12]:
                if idx + 1 >= len(r):
                    continue
                a = str(r[idx]).strip()
                b = str(r[idx + 1]).strip()
                if a or b:
                    sample_count += 1
                    if a.replace(",", ".").replace(".", "").isdigit() and b.replace(",", ".").replace(".", "").isdigit():
                        ok_count += 1
            if sample_count and ok_count >= max(1, sample_count // 2):
                should_split = True

        if not should_split:
            continue

        # Đổi header: Tổ hợp cọc -> D1, chèn D2 ngay sau đó
        new_cols = cols[:]
        new_cols[idx] = "D1"
        new_cols.insert(idx + 1, "D2")

        # Không chèn giá trị vào rows.
        # Vì rows hiện tại đang có: 6,10,16,14.5...
        # Chỉ cần chèn header D2 là các giá trị tự dịch đúng cột.
        fixed_rows = []
        for r in rows:
            rr = list(r)

            # Nếu dòng thiếu ô so với header mới thì pad trống cuối dòng
            if len(rr) < len(new_cols):
                rr = rr + [""] * (len(new_cols) - len(rr))

            fixed_rows.append(rr)

        t["columns"] = new_cols
        t["rows"] = fixed_rows
        t["title"] = t.get("title") or "Bảng đã tách Tổ hợp cọc D1/D2"

    return tables

class TableEditor(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=UI_SURFACE)
        self.pack(fill="both", expand=True)
        self.tables = []
        self.current = 0
        self.active_cell = None

        top = tk.Frame(self, bg=UI_SURFACE)
        top.pack(fill="x", pady=(0, 8))

        tk.Label(top, text="Bảng:", bg=UI_SURFACE, fg=UI_TEXT).pack(side="left")
        self.combo = RoundedMappingDropdown(
            top,
            values=[],
            variable=tk.StringVar(),
            bg_color="#f8fbff",
            border_color="#bcd2ee",
            width=250,
            height=34,
            radius=8,
        )
        self.combo.pack(side="left", padx=6)
        self.combo.bind("<<ComboboxSelected>>", lambda e: self.switch_table())

        ui_button(top, "Thêm dòng", self.add_row, width=11, variant="soft").pack(side="left", padx=3)
        ui_button(top, "Xóa dòng", self.delete_row, width=10).pack(side="left", padx=3)
        ui_button(top, "Sửa ô", self.edit_selected_cell, width=10).pack(side="left", padx=3)
        ui_button(top, "Xóa ô", self.clear_selected_cell, width=9).pack(side="left", padx=3)

        tree_frame = tk.Frame(self, bg=UI_SURFACE, highlightthickness=1, highlightbackground=UI_BORDER)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_frame, show="headings", style="Preview.Treeview")
        self.tree.tag_configure("preview_odd", background="#f7f7f3", foreground="#1f2933")
        self.tree.tag_configure("preview_even", background="#e9eee9", foreground="#1f2933")
        self.v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree_yview, style="Vertical.TScrollbar")
        self.h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self._tree_xview, style="Horizontal.TScrollbar")
        self.preview_grid_lines = []

        self.tree.configure(yscrollcommand=self._tree_yscroll, xscrollcommand=self.h_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.v_scroll.grid(row=0, column=1, sticky="ns")
        self.h_scroll.grid(row=1, column=0, sticky="ew")

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self.edit_cell)
        self.tree.bind("<ButtonRelease-1>", self.remember_active_cell)
        self.tree.bind("<Configure>", lambda _e: self._refresh_preview_grid())
        self.tree.bind("<F2>", lambda e: self.edit_selected_cell())
        self.tree.bind("<Return>", lambda e: self.edit_selected_cell())
        self.tree.bind("<Delete>", lambda e: self.clear_selected_cell())
        self.tree.bind("<MouseWheel>", self._on_tree_mousewheel)
        self.tree.bind("<Shift-MouseWheel>", self._on_tree_shift_mousewheel)

    def _on_tree_mousewheel(self, event):
        try:
            self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
            self._refresh_preview_grid()
            return "break"
        except Exception:
            return None

    def _on_tree_shift_mousewheel(self, event):
        try:
            self.tree.xview_scroll(int(-1 * (event.delta / 120)), "units")
            self._refresh_preview_grid()
            return "break"
        except Exception:
            return None

    def _tree_xview(self, *args):
        self.tree.xview(*args)
        self._refresh_preview_grid()

    def _tree_yview(self, *args):
        self.tree.yview(*args)
        self._refresh_preview_grid()

    def _tree_yscroll(self, first, last):
        self.v_scroll.set(first, last)
        self._refresh_preview_grid()

    def _clear_preview_grid(self):
        for line in getattr(self, "preview_grid_lines", []):
            try:
                line.destroy()
            except Exception:
                pass
        self.preview_grid_lines = []

    def _refresh_preview_grid(self):
        try:
            cols = list(self.tree["columns"])
        except Exception:
            return
        self._clear_preview_grid()
        if not cols:
            return
        try:
            total_width = sum(int(self.tree.column(c, "width")) for c in cols)
            if total_width <= 0:
                return
            x_offset = int(float(self.tree.xview()[0]) * total_width)
            visible_width = self.tree.winfo_width()
            height = self.tree.winfo_height()
            x = 0
            for c in cols[:-1]:
                x += int(self.tree.column(c, "width"))
                line_x = x - x_offset
                if -2 <= line_x <= visible_width + 2:
                    line = tk.Frame(self.tree, bg="#cfc9bc", width=1, bd=0, highlightthickness=0)
                    line.place(x=line_x, y=0, width=1, height=height)
                    line.lift()
                    self.preview_grid_lines.append(line)

            for item in self.tree.get_children():
                bbox = self.tree.bbox(item)
                if not bbox:
                    continue
                _x, y, _w, row_h = bbox
                line_y = y + row_h - 1
                if 0 <= line_y <= height:
                    line = tk.Frame(self.tree, bg="#d8d2c8", height=1, bd=0, highlightthickness=0)
                    line.place(x=0, y=line_y, width=visible_width, height=1)
                    line.lift()
                    self.preview_grid_lines.append(line)
        except Exception:
            self._clear_preview_grid()

    def set_tables(self, tables):
        self.tables = tables or []
        self.current = 0
        names = [t.get("title") or f"Bảng {i+1}" for i, t in enumerate(self.tables)]
        self.combo["values"] = names
        if names:
            self.combo.current(0)
        self.render()

    def switch_table(self):
        self.sync_current_from_tree()
        self.current = self.combo.current()
        self.render()

    def render(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree["columns"] = []

        if not self.tables:
            return

        t = self.tables[self.current]
        cols = t["columns"]
        self.tree["columns"] = cols

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=max(100, min(190, len(str(c)) * 10)), anchor="center", stretch=False)

        for idx, row in enumerate(t["rows"]):
            rr = row[:len(cols)] + [""] * max(0, len(cols) - len(row))
            tag = "preview_even" if idx % 2 else "preview_odd"
            self.tree.insert("", "end", values=rr, tags=(tag,))
        self.tree.after_idle(self._refresh_preview_grid)

    def sync_current_from_tree(self):
        if not self.tables:
            return
        rows = []
        for item in self.tree.get_children():
            rows.append(list(self.tree.item(item, "values")))
        self.tables[self.current]["rows"] = rows

    def get_tables(self):
        self.sync_current_from_tree()
        return self.tables

    def get_current_table(self):
        self.sync_current_from_tree()
        if not self.tables:
            return None
        return self.tables[self.current]

    def add_row(self):
        if not self.tables:
            return
        cols = self.tables[self.current]["columns"]
        idx = len(self.tree.get_children())
        tag = "preview_even" if idx % 2 else "preview_odd"
        self.tree.insert("", "end", values=[""] * len(cols), tags=(tag,))
        self.tree.after_idle(self._refresh_preview_grid)
        self.sync_current_from_tree()

    def delete_row(self):
        for item in self.tree.selection():
            self.tree.delete(item)
        self.sync_current_from_tree()

    def remember_active_cell(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if item and col:
            self.active_cell = (item, col)

    def edit_selected_cell(self):
        if self.active_cell:
            item, col = self.active_cell
        else:
            selected = self.tree.selection()
            if not selected:
                return
            item = selected[0]
            col = "#1"
        self._edit_cell_by_item_col(item, col)

    def clear_selected_cell(self):
        if self.active_cell:
            item, col = self.active_cell
        else:
            selected = self.tree.selection()
            if not selected:
                return
            item = selected[0]
            col = "#1"

        if not item or not col:
            return

        idx = int(col.replace("#", "")) - 1
        vals = list(self.tree.item(item, "values"))
        vals += [""] * (idx + 1 - len(vals))
        vals[idx] = ""
        self.tree.item(item, values=vals)
        self.sync_current_from_tree()

    def edit_cell(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or not col:
            return
        self.active_cell = (item, col)
        self._edit_cell_by_item_col(item, col)

    def _edit_cell_by_item_col(self, item, col):
        idx = int(col.replace("#", "")) - 1
        bbox = self.tree.bbox(item, col)
        if not bbox:
            self.tree.see(item)
            bbox = self.tree.bbox(item, col)
            if not bbox:
                return

        x, y, w, h = bbox
        vals = list(self.tree.item(item, "values"))
        vals += [""] * (idx + 1 - len(vals))
        old = vals[idx]

        ent = tk.Entry(self.tree)
        ent.place(x=x, y=y, width=w, height=h)
        ent.insert(0, old)
        ent.focus()
        ent.select_range(0, "end")

        def save(_=None):
            vals2 = list(self.tree.item(item, "values"))
            vals2 += [""] * (idx + 1 - len(vals2))
            vals2[idx] = ent.get()
            self.tree.item(item, values=vals2)
            ent.destroy()
            self.sync_current_from_tree()

        def cancel(_=None):
            ent.destroy()

        ent.bind("<Return>", save)
        ent.bind("<FocusOut>", save)
        ent.bind("<Escape>", cancel)

class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        self._setup_responsive_metrics()
        self.app_logo_img = None

        env = load_env_values()
        self.api_key_var = tk.StringVar(value=env["GEMINI_API_KEY"])
        self.model_var = tk.StringVar(value=env["GEMINI_MODEL"] or DEFAULT_MODEL)

        self.template_var = tk.StringVar(value="Bảng bất kỳ - tự nhận cột")
        self.image_path = None
        self.excel_path = None
        self.tk_img = None
        self.workbook = None
        self.excel_folder = None
        self.sheet_var = tk.StringVar()
        self.header_row = None
        self.excel_headers = []
        self.tables = []
        self.user_name, self.user_role = current_user_role_labels()
        self.approval_dialog_open = False
        self.member_locked = (not is_admin_build()) and (not is_machine_approved())
        if self.member_locked:
            try:
                self.root.withdraw()
            except Exception:
                pass

        self.build_ui()
        self.root.bind_all("<Control-v>", self.paste_image_from_clipboard)
        self.root.bind_all("<Control-V>", self.paste_image_from_clipboard)
        self.root.after(300, self._check_member_approval_loop)

    def _check_member_approval_loop(self):
        if is_admin_build():
            return
        try:
            approved = is_machine_approved()
            if not approved:
                self.member_locked = True
                try:
                    self.root.withdraw()
                except Exception:
                    pass
                if not self.approval_dialog_open:
                    self.show_member_approval_dialog()
            elif self.member_locked:
                self.member_locked = False
                try:
                    self.root.deiconify()
                    self.root.lift()
                    self.status.config(text="Máy đã được duyệt.")
                except Exception:
                    pass
        finally:
            try:
                self.root.after(3000, self._check_member_approval_loop)
            except Exception:
                pass

    def _center_dialog_on_screen(self, win):
        try:
            win.update_idletasks()
            width = win.winfo_width()
            height = win.winfo_height()
            screen_w = win.winfo_screenwidth()
            screen_h = win.winfo_screenheight()
            x = max(0, (screen_w - width) // 2)
            y = max(0, (screen_h - height) // 2)
            win.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def show_member_approval_dialog(self):
        if self.approval_dialog_open:
            return
        self.approval_dialog_open = True
        machine_code = get_machine_code()
        win = tk.Toplevel(self.root)
        win.title("Yêu cầu duyệt sử dụng")
        win.configure(bg=UI_SURFACE)
        win.resizable(False, False)
        try:
            if self.root.state() != "withdrawn":
                win.transient(self.root)
        except Exception:
            pass
        win.grab_set()
        try:
            win.protocol("WM_DELETE_WINDOW", self.root.destroy)
        except Exception:
            pass
        try:
            win.bind("<Destroy>", lambda _e: setattr(self, "approval_dialog_open", False))
        except Exception:
            pass

        body = tk.Frame(win, bg=UI_SURFACE, padx=22, pady=18)
        body.pack(fill="both", expand=True)
        tk.Label(body, text="Máy này chưa được duyệt", bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(body, text="Gửi mã máy bên dưới cho Admin để nhận mã duyệt.", bg=UI_SURFACE, fg=UI_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 12))

        tk.Label(body, text="Mã máy", bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        machine_var = tk.StringVar(value=machine_code)
        machine_entry = tk.Entry(body, textvariable=machine_var, width=34, relief="solid", bd=1, font=("Segoe UI", 10))
        machine_entry.pack(fill="x", pady=(4, 10))
        machine_entry.configure(state="readonly")

        tk.Label(body, text="Mã duyệt", bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        code_var = tk.StringVar()
        code_entry = tk.Entry(body, textvariable=code_var, width=34, relief="solid", bd=1, font=("Segoe UI", 10))
        code_entry.pack(fill="x", pady=(4, 12))
        code_entry.focus_set()

        actions = tk.Frame(body, bg=UI_SURFACE)
        actions.pack(fill="x")

        def copy_machine():
            self.root.clipboard_clear()
            self.root.clipboard_append(machine_code)
            try:
                self.status.config(text="Đã copy mã máy.")
            except Exception:
                pass

        def approve():
            if save_machine_approval(code_var.get()):
                win.destroy()
                self.member_locked = False
                try:
                    self.root.deiconify()
                    self.root.lift()
                    self.status.config(text="Máy đã được duyệt.")
                except Exception:
                    pass
            else:
                messagebox.showerror("Mã không đúng", "Mã duyệt không hợp lệ. Kiểm tra lại mã Admin gửi.")

        ui_button(actions, "Copy mã máy", copy_machine, width=12, variant="soft").pack(side="left", padx=(0, 8))
        ui_button(actions, "Xác nhận", approve, width=11, variant="success").pack(side="left")
        ui_button(actions, "Thoát", self.root.destroy, width=9).pack(side="right")

        self._center_dialog_on_screen(win)
        try:
            win.lift()
            win.focus_force()
            code_entry.focus_force()
            win.attributes("-topmost", True)
            win.after(700, lambda: win.attributes("-topmost", False))
        except Exception:
            pass
        self.root.wait_window(win)

    def open_admin_approval_panel(self):
        win = tk.Toplevel(self.root)
        win.title("Duyệt máy thành viên")
        win.configure(bg=UI_SURFACE)
        win.geometry("720x560")
        win.minsize(640, 500)

        body = tk.Frame(win, bg=UI_SURFACE, padx=22, pady=18)
        body.pack(fill="both", expand=True)
        tk.Label(body, text="Duyệt máy thành viên", bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(body, text="Nhập mã máy thành viên gửi, sau đó gửi lại mã duyệt cho họ.", bg=UI_SURFACE, fg=UI_MUTED, font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 12))

        tk.Label(body, text="Mã máy", bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        machine_var = tk.StringVar()
        machine_entry = tk.Entry(body, textvariable=machine_var, width=42, relief="solid", bd=1, font=("Segoe UI", 10))
        machine_entry.pack(fill="x", pady=(4, 10))

        tk.Label(body, text="Mã duyệt", bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        approval_var = tk.StringVar()
        approval_entry = tk.Entry(body, textvariable=approval_var, width=42, relief="solid", bd=1, font=("Segoe UI", 10))
        approval_entry.pack(fill="x", pady=(4, 12))
        approval_entry.configure(state="readonly")

        actions = tk.Frame(body, bg=UI_SURFACE)
        actions.pack(fill="x")

        list_header = tk.Frame(body, bg=UI_SURFACE)
        list_header.pack(fill="x", pady=(18, 6))
        tk.Label(list_header, text="Danh sách máy đã duyệt", bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 10, "bold")).pack(side="left")
        list_frame = tk.Frame(body, bg=UI_SURFACE, highlightthickness=1, highlightbackground=UI_BORDER)
        list_frame.pack(fill="both", expand=True)
        approved_tree = ttk.Treeview(
            list_frame,
            columns=("machine", "code", "time"),
            show="headings",
            height=8,
        )
        approved_tree.heading("machine", text="Mã máy")
        approved_tree.heading("code", text="Mã duyệt")
        approved_tree.heading("time", text="Thời gian duyệt")
        approved_tree.column("machine", width=230, anchor="w")
        approved_tree.column("code", width=150, anchor="center")
        approved_tree.column("time", width=150, anchor="center")
        approved_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=approved_tree.yview)
        approved_tree.configure(yscrollcommand=approved_scroll.set)
        approved_tree.grid(row=0, column=0, sticky="nsew")
        approved_scroll.grid(row=0, column=1, sticky="ns")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        list_actions = tk.Frame(body, bg=UI_SURFACE)
        list_actions.pack(fill="x", pady=(8, 0))

        def refresh_list():
            import_local_approval_to_admin_list()
            for item in approved_tree.get_children():
                approved_tree.delete(item)
            for row in load_admin_approved_machines():
                machine = row.get("machine_code", "")
                approved_tree.insert(
                    "",
                    "end",
                    iid=machine,
                    values=(machine, row.get("approval_code", ""), row.get("approved_at", "")),
                )

        def fill_from_selected(_event=None):
            selected = approved_tree.selection()
            if not selected:
                return
            values = approved_tree.item(selected[0], "values")
            if values:
                machine_var.set(values[0])
                approval_entry.configure(state="normal")
                approval_var.set(values[1] if len(values) > 1 else "")
                approval_entry.configure(state="readonly")

        approved_tree.bind("<<TreeviewSelect>>", fill_from_selected)

        def generate():
            machine = str(machine_var.get() or "").strip().upper()
            code = remember_admin_approved_machine(machine_var.get())
            if not code:
                messagebox.showwarning("Thiếu mã máy", "Bạn chưa nhập mã máy cần duyệt.")
                return
            approval_entry.configure(state="normal")
            approval_var.set(code)
            approval_entry.configure(state="readonly")
            refresh_list()

        def copy_code():
            if not approval_var.get():
                generate()
            if not approval_var.get():
                return
            self.root.clipboard_clear()
            self.root.clipboard_append(approval_var.get())
            self.status.config(text="Đã copy mã duyệt.")

        def delete_selected():
            selected = approved_tree.selection()
            if not selected:
                messagebox.showinfo("Chưa chọn máy", "Chọn một máy trong danh sách trước khi xóa.")
                return
            machine = str(selected[0])
            if not messagebox.askyesno(
                "Xóa máy đã duyệt",
                f"Xóa máy này khỏi danh sách đã duyệt?\nMã duyệt cũ trên máy đó sẽ không dùng lại được.\n\n{machine}",
            ):
                return
            delete_admin_approved_machine(machine)
            refresh_list()
            self.status.config(text="Đã xóa máy và làm hết hiệu lực mã duyệt cũ.")

        def open_list_menu(event):
            row_id = approved_tree.identify_row(event.y)
            if row_id:
                approved_tree.selection_set(row_id)
                fill_from_selected()
            menu = tk.Menu(win, tearoff=0)
            menu.add_command(label="Xóa máy này", command=delete_selected)
            menu.tk_popup(event.x_root, event.y_root)

        approved_tree.bind("<Delete>", lambda _e: delete_selected())
        approved_tree.bind("<Button-3>", open_list_menu)

        ui_button(actions, "Tạo mã duyệt", generate, width=13, variant="primary").pack(side="left", padx=(0, 8))
        ui_button(actions, "Copy mã", copy_code, width=10, variant="soft").pack(side="left")
        ui_button(actions, "Đóng", win.destroy, width=9).pack(side="right")
        ui_button(list_header, "Xóa máy", delete_selected, width=10, variant="warn").pack(side="right")
        ui_button(list_actions, "Xóa máy đã chọn", delete_selected, width=15, variant="warn").pack(side="left")
        ui_button(list_actions, "Tải lại danh sách", refresh_list, width=14, variant="soft").pack(side="left", padx=(8, 0))
        machine_entry.focus_set()
        refresh_list()

    def _setup_responsive_metrics(self):
        try:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
        except Exception:
            sw, sh = 1500, 900

        self.screen_w = sw
        self.screen_h = sh
        self.compact_ui = sw < 1500 or sh < 850
        self.tiny_ui = sw < 1280 or sh < 760

        win_w = min(1500, max(1080, int(sw * 0.96)))
        win_h = min(900, max(700, int(sh * 0.92)))
        try:
            self.root.geometry(f"{win_w}x{win_h}")
            self.root.minsize(1024 if self.tiny_ui else 1120, 680 if self.tiny_ui else 720)
        except Exception:
            pass

        if self.tiny_ui:
            self.sidebar_w = 124
            self.main_padx = 10
            self.main_pady = 10
            self.card_padx = 10
            self.card_pady = 8
            self.workspace_mins = (170, 520, 255)
            self.mapping_canvas_h = 220
            self.logo_max = (108, 100)
        elif self.compact_ui:
            self.sidebar_w = 140
            self.main_padx = 14
            self.main_pady = 12
            self.card_padx = 11
            self.card_pady = 9
            self.workspace_mins = (200, 600, 285)
            self.mapping_canvas_h = 240
            self.logo_max = (124, 114)
        else:
            self.sidebar_w = 168
            self.main_padx = 22
            self.main_pady = 18
            self.card_padx = 14
            self.card_pady = 12
            self.workspace_mins = (245, 760, 340)
            self.mapping_canvas_h = 260
            self.logo_max = (146, 136)

    def setup_window_icon(self):
        icon_file = resource_path(*APP_ICON_ICO.parts)
        taskbar_file = resource_path(*APP_TASKBAR_PNG.parts)
        self._apply_window_icon(icon_file, taskbar_file)
        try:
            self.root.after(250, lambda: self._apply_window_icon(icon_file, taskbar_file))
        except Exception:
            pass

    def _apply_window_icon(self, icon_file, taskbar_file):
        try:
            if icon_file.exists():
                self.root.iconbitmap(default=str(icon_file))
                self.root.wm_iconbitmap(default=str(icon_file))
        except Exception:
            pass
        try:
            if taskbar_file.exists():
                base = Image.open(taskbar_file).convert("RGBA")
                icon_imgs = [
                    ImageTk.PhotoImage(rounded_icon_image(base, (16, 16))),
                    ImageTk.PhotoImage(rounded_icon_image(base, (32, 32))),
                    ImageTk.PhotoImage(rounded_icon_image(base, (64, 64))),
                    ImageTk.PhotoImage(rounded_icon_image(base, (128, 128))),
                ]
                self.root.iconphoto(True, *icon_imgs)
                self.window_icon_imgs = icon_imgs
        except Exception:
            pass

    def setup_theme(self):
        self.root.configure(bg=UI_BG)
        try:
            style = ttk.Style(self.root)
            style.theme_use("clam")
            style.configure(".", font=("Segoe UI", 9), background=UI_BG, foreground=UI_TEXT)
            style.configure(
                "TCombobox",
                fieldbackground=UI_SURFACE,
                background=UI_SURFACE,
                foreground=UI_TEXT,
                bordercolor=UI_BORDER,
                lightcolor=UI_BORDER,
                darkcolor=UI_BORDER,
                arrowcolor=UI_MUTED,
                relief="flat",
                padding=(10, 7),
                arrowsize=14,
            )
            style.map(
                "TCombobox",
                fieldbackground=[("readonly", UI_SURFACE), ("focus", "#fbfdff")],
                bordercolor=[("focus", UI_PRIMARY), ("!focus", UI_BORDER)],
                arrowcolor=[("active", UI_PRIMARY), ("!active", UI_MUTED)],
            )
            style.configure(
                "SoftBlue.TCombobox",
                fieldbackground="#f8fbff",
                background="#edf6ff",
                foreground=UI_TEXT,
                bordercolor="#bcd2ee",
                lightcolor="#bcd2ee",
                darkcolor="#bcd2ee",
                arrowcolor="#5f728b",
                relief="flat",
                padding=(10, 7),
                arrowsize=14,
            )
            style.map(
                "SoftBlue.TCombobox",
                fieldbackground=[("readonly", "#f8fbff"), ("focus", "#eef7ff")],
                background=[("readonly", "#edf6ff"), ("active", "#dceeff")],
                bordercolor=[("focus", UI_PRIMARY), ("!focus", "#bcd2ee")],
                arrowcolor=[("active", UI_PRIMARY), ("!active", "#5f728b")],
                selectbackground=[("readonly", "#dbeafe")],
                selectforeground=[("readonly", UI_TEXT)],
            )
            style.configure(
                "TEntry",
                fieldbackground=UI_SURFACE,
                foreground=UI_TEXT,
                bordercolor=UI_BORDER,
                lightcolor=UI_BORDER,
                darkcolor=UI_BORDER,
                relief="flat",
                padding=(10, 7),
            )
            style.map("TEntry", bordercolor=[("focus", UI_PRIMARY), ("!focus", UI_BORDER)])
            style.configure(
                "Treeview",
                background=UI_SURFACE,
                fieldbackground=UI_SURFACE,
                foreground=UI_TEXT,
                rowheight=28,
                bordercolor=UI_BORDER,
                lightcolor=UI_BORDER,
                darkcolor=UI_BORDER,
            )
            style.configure(
                "Treeview.Heading",
                background=UI_SURFACE_2,
                foreground=UI_TEXT,
                font=("Segoe UI", 9, "bold"),
                relief="flat",
                padding=(8, 7),
            )
            style.map("Treeview", background=[("selected", "#dbeafe")], foreground=[("selected", UI_TEXT)])
            style.configure(
                "Preview.Treeview",
                background="#f7f7f3",
                fieldbackground="#f7f7f3",
                foreground="#1f2933",
                rowheight=28,
                bordercolor="#cfd7dd",
                lightcolor="#cfd7dd",
                darkcolor="#cfd7dd",
            )
            style.configure(
                "Preview.Treeview.Heading",
                background="#6b6358",
                foreground="#ffffff",
                font=("Segoe UI", 9, "bold"),
                relief="flat",
                padding=(8, 7),
            )
            style.map(
                "Preview.Treeview",
                background=[("selected", "#cfe5ff")],
                foreground=[("selected", UI_TEXT)],
            )
            style.map(
                "Preview.Treeview.Heading",
                background=[("active", "#5d554c"), ("!active", "#6b6358")],
                foreground=[("active", "#ffffff"), ("!active", "#ffffff")],
            )
            style.configure(
                "Vertical.TScrollbar",
                gripcount=0,
                background="#c8d5e6",
                darkcolor="#c8d5e6",
                lightcolor="#c8d5e6",
                troughcolor="#eef3f8",
                bordercolor="#eef3f8",
                arrowcolor=UI_MUTED,
                relief="flat",
                width=14,
            )
            style.configure(
                "Horizontal.TScrollbar",
                gripcount=0,
                background="#c8d5e6",
                darkcolor="#c8d5e6",
                lightcolor="#c8d5e6",
                troughcolor="#eef3f8",
                bordercolor="#eef3f8",
                arrowcolor=UI_MUTED,
                relief="flat",
                width=14,
            )
            style.map(
                "Vertical.TScrollbar",
                background=[("active", "#9fb2cc"), ("pressed", "#7f96b5")],
                arrowcolor=[("active", UI_PRIMARY), ("!active", UI_MUTED)],
            )
            style.map(
                "Horizontal.TScrollbar",
                background=[("active", "#9fb2cc"), ("pressed", "#7f96b5")],
                arrowcolor=[("active", UI_PRIMARY), ("!active", UI_MUTED)],
            )
        except Exception:
            pass

    def build_ui(self):
        try:
            self.root.state("zoomed")
        except Exception:
            pass
        self.setup_window_icon()
        self.setup_theme()

        def card(parent, padx=None, pady=None):
            if padx is None:
                padx = self.card_padx
            if pady is None:
                pady = self.card_pady
            frame = tk.Frame(
                parent,
                bg=UI_SURFACE,
                padx=padx,
                pady=pady,
                highlightthickness=1,
                highlightbackground="#e6edf5",
            )
            return frame

        def section_title(parent, title, subtitle=None):
            tk.Label(parent, text=title, font=("Segoe UI", 11, "bold"), bg=UI_SURFACE, fg=UI_TEXT).pack(anchor="w")
            if subtitle:
                tk.Label(parent, text=subtitle, font=("Segoe UI", 8), bg=UI_SURFACE, fg=UI_MUTED).pack(anchor="w", pady=(2, 0))

        shell = tk.Frame(self.root, bg=UI_BG)
        shell.pack(fill="both", expand=True)

        sidebar = tk.Frame(
            shell,
            width=self.sidebar_w,
            bg="#f8fbff",
            padx=8 if self.compact_ui else 10,
            pady=10 if self.compact_ui else 14,
            highlightthickness=1,
            highlightbackground="#e7edf6",
        )
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        brand = tk.Frame(sidebar, bg="#f8fbff")
        brand.pack(fill="x", pady=(0, 24))
        logo_file = resource_path(*APP_LOGO_PNG.parts)
        try:
            if logo_file.exists():
                logo_source = Image.open(logo_file).convert("RGBA")
                logo_source.thumbnail(self.logo_max, Image.LANCZOS)
                self.app_logo_img = ImageTk.PhotoImage(logo_source)
                tk.Label(brand, image=self.app_logo_img, bg="#f8fbff").pack(anchor="center")
        except Exception:
            self.app_logo_img = None

        nav_items = [
            ("⌂", "Trang chủ", True),
            ("▦", "Excel", False),
            ("◷", "Lịch sử", False),
            ("▤", "Mẫu mapping", False),
            ("⚙", "Cài đặt", False),
            ("?", "Trợ giúp", False),
            ("i", "Giới thiệu", False),
        ]
        for icon, text, active in nav_items:
            bg = "#eaf2ff" if active else "#f8fbff"
            fg = UI_PRIMARY if active else "#667085"
            row = tk.Frame(sidebar, bg=bg, padx=10, pady=9)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=icon, width=2, bg=bg, fg=fg, font=("Segoe UI", 11)).pack(side="left")
            tk.Label(row, text=text, bg=bg, fg=fg, font=("Segoe UI", 9, "bold" if active else "normal")).pack(side="left", padx=(6, 0))

        sidebar_spacer = tk.Frame(sidebar, bg="#f8fbff")
        sidebar_spacer.pack(fill="both", expand=True)
        user_box = card(sidebar, padx=8, pady=8)
        user_box.configure(bg="#ffffff")
        user_box.pack(fill="x", pady=(8, 0))
        tk.Label(user_box, text=self.user_name, font=("Segoe UI", 9, "bold"), bg=UI_SURFACE, fg=UI_TEXT).pack(anchor="w")
        tk.Label(user_box, text=self.user_role, font=("Segoe UI", 8), bg=UI_SURFACE, fg=UI_MUTED).pack(anchor="w")
        self.status = tk.Label(
            user_box,
            text="● Sẵn sàng",
            anchor="w",
            fg=UI_SUCCESS,
            bg=UI_SURFACE,
            font=("Segoe UI", 8, "bold"),
            wraplength=128,
            justify="left",
        )
        self.status.pack(fill="x", pady=(10, 0))
        if is_admin_build():
            ui_button(user_box, "Duyệt máy", self.open_admin_approval_panel, width=12, variant="warn").pack(fill="x", pady=(10, 0))

        main = tk.Frame(shell, bg=UI_BG, padx=self.main_padx, pady=self.main_pady)
        main.pack(side="left", fill="both", expand=True)

        header = tk.Frame(main, bg=UI_BG)
        header.pack(fill="x")
        title_box = tk.Frame(header, bg=UI_BG)
        title_box.pack(side="left", fill="x", expand=True)
        if APP_TITLE:
            tk.Label(title_box, text=APP_TITLE, font=("Segoe UI", 20 if self.compact_ui else 22, "bold"), bg=UI_BG, fg=UI_TEXT).pack(anchor="w")
        tk.Label(title_box, text="Ứng dụng Phục hồi & Quản lý Dữ liệu Cọc", font=("Segoe UI", 9 if self.compact_ui else 10), bg=UI_BG, fg=UI_MUTED).pack(anchor="w", pady=(3, 0))
        if not self.tiny_ui:
            for text in ("☼", "🔔", "?"):
                tk.Label(header, text=text, bg=UI_SURFACE, fg=UI_TEXT, width=3, height=2, font=("Segoe UI", 10 if self.compact_ui else 11), highlightthickness=1, highlightbackground="#edf2f7").pack(side="left", padx=4 if self.compact_ui else 6)
            tk.Label(header, text="A", bg="#6366f1", fg="#ffffff", width=3, height=2, font=("Segoe UI", 10 if self.compact_ui else 11, "bold")).pack(side="left", padx=(10, 6))
            profile = tk.Frame(header, bg=UI_BG)
            profile.pack(side="left")
            tk.Label(profile, text=self.user_name, font=("Segoe UI", 9, "bold"), bg=UI_BG, fg=UI_TEXT).pack(anchor="w")
            tk.Label(profile, text=self.user_role, font=("Segoe UI", 8), bg=UI_BG, fg=UI_MUTED).pack(anchor="w")

        toolbar = card(main, padx=12 if self.compact_ui else 18, pady=10 if self.compact_ui else 14)
        toolbar.pack(fill="x", pady=(12 if self.compact_ui else 18, 10 if self.compact_ui else 12))
        toolbar.grid_columnconfigure(0, weight=5, uniform="toolbar")
        toolbar.grid_columnconfigure(1, weight=4, uniform="toolbar")
        toolbar.grid_columnconfigure(2, weight=4, uniform="toolbar")
        toolbar.grid_rowconfigure(0, weight=1)

        source = tk.Frame(toolbar, bg=UI_SURFACE, padx=6 if self.compact_ui else 10, pady=4)
        source.grid(row=0, column=0, sticky="nsew", padx=(0, 8 if self.compact_ui else 16))
        section_title(source, "NGUỒN DỮ LIỆU")
        source_buttons = tk.Frame(source, bg=UI_SURFACE)
        source_buttons.pack(anchor="w", pady=(8 if self.compact_ui else 12, 0))
        source_btns = [
            ("Chọn Excel", self.choose_excel, "primary"),
            ("Đọc workbook", self.scan_current_workbook, "default"),
            ("Đọc từng sheet", self.read_each_sheet_content, "default"),
            ("Đọc công thức", self.read_current_excel_formulas, "default"),
            ("Đọc lại Excel", self.refresh_excel_header_info, "soft"),
        ]
        for idx, (text, command, variant) in enumerate(source_btns):
            ui_button(source_buttons, text, command, width=10 if self.compact_ui else 11, variant=variant).grid(
                row=idx // (2 if self.compact_ui else 5),
                column=idx % (2 if self.compact_ui else 5),
                padx=(0, 4 if self.compact_ui else 8),
                pady=2,
                sticky="w",
            )

        process = tk.Frame(toolbar, bg=UI_SURFACE, highlightthickness=1, highlightbackground="#edf2f7", padx=10 if self.compact_ui else 20, pady=4)
        process.grid(row=0, column=1, sticky="nsew", padx=(0, 8 if self.compact_ui else 16))
        section_title(process, "XỬ LÝ DỮ LIỆU")
        process_buttons = tk.Frame(process, bg=UI_SURFACE)
        process_buttons.pack(anchor="center", pady=(8 if self.compact_ui else 12, 0))
        for idx, (text, command, variant) in enumerate([
            ("Đọc bảng", self.run_gemini, "soft"),
            ("Đọc phiếu cọc", self.run_gemini_phieu_coc, "soft"),
            ("Auto map cột", self.build_mapping, "warn"),
        ]):
            ui_button(process_buttons, text, command, width=11 if self.compact_ui else 12, variant=variant).grid(
                row=idx if self.compact_ui else 0,
                column=0 if self.compact_ui else idx,
                padx=3 if self.compact_ui else 5,
                pady=2,
            )

        export = tk.Frame(toolbar, bg=UI_SURFACE, highlightthickness=1, highlightbackground="#edf2f7", padx=10 if self.compact_ui else 20, pady=4)
        export.grid(row=0, column=2, sticky="nsew")
        section_title(export, "XEM & XUẤT")
        export_buttons = tk.Frame(export, bg=UI_SURFACE)
        export_buttons.pack(anchor="center", pady=(8 if self.compact_ui else 12, 0))
        ui_button(export_buttons, "Xem trước ghép", self.preview_excel, width=12 if self.compact_ui else 14, variant="soft").grid(row=0 if not self.compact_ui else 0, column=0, padx=4, pady=2)
        ui_button(export_buttons, "Xuất ra Excel", self.fill_excel, width=12 if self.compact_ui else 14, variant="success").grid(row=0 if not self.compact_ui else 1, column=1 if not self.compact_ui else 0, padx=4, pady=2)

        filters = card(main, padx=14, pady=9)
        filters.pack(fill="x", pady=(0, 12))
        tk.Label(filters, text="Sheet:", bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 9, "bold")).pack(side="left", padx=(4, 8))
        self.sheet_combo = RoundedMappingDropdown(
            filters,
            values=[],
            variable=self.sheet_var,
            bg_color="#f8fbff",
            border_color="#bcd2ee",
            width=180 if self.compact_ui else 220,
            height=34,
            radius=8,
        )
        self.sheet_combo.pack(side="left", padx=(0, 10 if self.compact_ui else 18))
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_excel_header_info())
        tk.Label(filters, text="Chế độ đọc bảng:", bg=UI_SURFACE, fg=UI_MUTED).pack(side="left", padx=(4, 8))
        self.template_combo = RoundedMappingDropdown(
            filters,
            values=["Bảng bất kỳ - tự nhận cột"],
            variable=self.template_var,
            bg_color="#f8fbff",
            border_color="#bcd2ee",
            width=260 if self.compact_ui else 330,
            height=34,
            radius=8,
        )
        self.template_combo["values"] = ["Bảng bất kỳ - tự nhận cột"]
        self.template_combo.pack(side="left")

        workspace = tk.Frame(main, bg=UI_BG)
        workspace.pack(fill="both", expand=True)
        left_min, center_min, right_min = self.workspace_mins
        workspace.grid_columnconfigure(0, weight=1, minsize=left_min)
        workspace.grid_columnconfigure(1, weight=5, minsize=center_min)
        workspace.grid_columnconfigure(2, weight=1, minsize=right_min)
        workspace.grid_rowconfigure(0, weight=1)

        left_col = tk.Frame(workspace, bg=UI_BG)
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        image_card = card(left_col)
        image_card.pack(fill="both", expand=True)
        section_title(image_card, "ẢNH OCR")
        upload_box = tk.Frame(image_card, bg="#fbfdff", highlightthickness=1, highlightbackground="#d7e3f2", padx=12, pady=12)
        upload_box.pack(fill="both", expand=True, pady=(12, 0))
        self.img_label = tk.Label(
            upload_box,
            text="Kéo thả ảnh OCR vào đây\n\nHỗ trợ: .jpg, .png, .jpeg",
            bg="#fbfdff",
            fg=UI_MUTED,
            font=("Segoe UI", 10),
            justify="center",
        )
        self.img_label.pack(fill="both", expand=True)
        self.img_label.bind("<Control-v>", self.paste_image_from_clipboard)
        self.img_label.bind("<Control-V>", self.paste_image_from_clipboard)
        ui_button(upload_box, "Chọn ảnh / Tải lên", self.choose_image, width=18, variant="primary").pack(pady=(10, 0))

        info = card(left_col)
        info.pack(fill="x", pady=(12, 0))
        section_title(info, "THÔNG TIN EXCEL")
        self.excel_info = tk.Text(
            info,
            height=9,
            wrap="word",
            bg="#fbfdff",
            fg=UI_TEXT,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=UI_BORDER,
            font=("Segoe UI", 8),
        )
        self.excel_info.pack(fill="both", expand=True, pady=(10, 0))

        center_col = card(workspace)
        center_col.grid(row=0, column=1, sticky="nsew", padx=(0, 8))
        section_title(center_col, "PREVIEW BẢNG", "Kiểm tra và sửa dữ liệu trước khi đưa vào Excel")
        self.table_editor = TableEditor(center_col)

        right_col = card(workspace)
        right_col.mapping_canvas_h = self.mapping_canvas_h
        right_col.grid(row=0, column=2, sticky="nsew")
        section_title(right_col, "XÁC NHẬN ÁNH XẠ CỘT", "Kéo thả để ánh xạ dữ liệu giữa 2 nguồn")
        self.mapping_editor = MappingEditor(right_col)

        workflow = card(main, padx=12 if self.compact_ui else 18, pady=8 if self.compact_ui else 14)
        workflow.pack(fill="x", pady=(8 if self.compact_ui else 12, 0))
        tk.Label(workflow, text="QUY TRÌNH XỬ LÝ", bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 10, "bold")).pack(anchor="w")
        steps = tk.Frame(workflow, bg=UI_SURFACE)
        steps.pack(fill="x", pady=(12, 0))
        for i, (title, sub, color) in enumerate([
            ("Chọn file / ảnh OCR", "Tải lên ảnh hoặc chọn file Excel", UI_PRIMARY),
            ("Đọc & Trích xuất", "Hệ thống đọc và trích xuất dữ liệu", "#38bdf8"),
            ("Kiểm tra & Sửa dữ liệu", "Xem trước và chỉnh sửa dữ liệu", "#14b8a6"),
            ("Ánh xạ & Xác nhận", "Map cột và xác nhận dữ liệu", UI_SUCCESS),
            ("Xuất ra Excel", "Xuất dữ liệu đã xử lý ra Excel", UI_PRIMARY),
        ]):
            item = tk.Frame(steps, bg=UI_SURFACE)
            item.pack(side="left", fill="x", expand=True)
            tk.Label(item, text=str(i + 1), bg="#eef4ff", fg=color, width=3, height=2, font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 10))
            text_box = tk.Frame(item, bg=UI_SURFACE)
            text_box.pack(side="left", fill="x")
            tk.Label(text_box, text=title, bg=UI_SURFACE, fg=UI_TEXT, font=("Segoe UI", 8, "bold")).pack(anchor="w")
            tk.Label(text_box, text=sub, bg=UI_SURFACE, fg=UI_MUTED, font=("Segoe UI", 7)).pack(anchor="w", pady=(2, 0))

    def save_key(self):
        save_env(self.api_key_var.get(), self.model_var.get())
        self.status.config(text="Đã lưu model vào file .env")

    def run_gemini_phieu_coc(self):
        """
        Đọc phiếu cọc bằng Gemini.
        Workflow:
          1. BẮT BUỘC phải chọn Excel trước để biết cột cần điền.
          2. Gemini trả về 1 bảng duy nhất với cột Y HỆT Excel.
          3. Auto-map 1:1 vì tên cột đã khớp.
          4. Sẵn sàng điền vào Excel ngay.
        """
        if not self.image_path:
            messagebox.showwarning("Thiếu ảnh", "Bạn chưa chọn ảnh.")
            return
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("Thiếu API key", "Bạn chưa nhập Gemini API key.")
            return
        if not self.excel_headers:
            messagebox.showwarning(
                "Chưa chọn Excel",
                "Bạn cần chọn file Excel trước để tool biết cần đọc cột nào.\n\n"
                "Bước 1: Chọn Excel → bấm Đọc lại Excel\n"
                "Bước 2: Chọn ảnh phiếu cọc → bấm Đọc phiếu cọc"
            )
            return

        self.save_key()
        excel_col_names = [name for _, name in self.excel_headers]
        self.status.config(
            text=f"Đang gửi phiếu cọc lên Gemini... ({len(excel_col_names)} cột Excel: "
                 + ", ".join(excel_col_names[:5])
                 + ("..." if len(excel_col_names) > 5 else "") + ")"
        )
        self.root.update()

        try:
            tables, raw = call_gemini_phieu_coc(
                self.image_path, api_key, self.model_var.get().strip(),
                excel_columns=excel_col_names
            )

            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "phieu_coc_raw_response.txt").write_text(raw, encoding="utf-8")
            (out / "phieu_coc_tables.json").write_text(
                json.dumps(tables, ensure_ascii=False, indent=2), encoding="utf-8"
            )

            self.tables = tables
            self.table_editor.set_tables(tables)

            # Tự chuyển sang bảng có nhiều cột nhất (bảng dữ liệu, không phải key-value)
            best_idx = 0
            best_ncols = 0
            for i, t in enumerate(tables):
                nc = len(t.get("columns", []))
                if nc > best_ncols:
                    best_ncols = nc
                    best_idx = i
            if best_idx != 0:
                try:
                    self.table_editor.combo.current(best_idx)
                    self.table_editor.switch_table()
                except Exception:
                    pass

            # Auto-map: vì cột Gemini trả về đã đặt tên y hệt Excel
            # → auto_map_columns sẽ khớp 1:1, không cần user chỉnh
            self.build_mapping()

            total_rows = sum(len(t["rows"]) for t in tables)
            data_table = tables[best_idx] if tables else None
            data_cols = data_table.get("columns", []) if data_table else []
            data_rows = len(data_table.get("rows", [])) if data_table else 0

            self.status.config(
                text=f"Đã đọc phiếu cọc: {data_rows} dòng × {len(data_cols)} cột. "
                     "Kiểm tra preview rồi bấm 'Điền tiếp vào Excel'."
            )

            # Hiển thị tóm tắt trong excel_info
            info_lines = ["KẾT QUẢ ĐỌC PHIẾU CỌC\n", "=" * 50 + "\n"]
            info_lines.append(f"Cột Excel ({len(excel_col_names)}): " + " | ".join(excel_col_names) + "\n\n")
            for t in tables:
                title = t.get("title") or "Bảng"
                cols = t.get("columns", [])
                rows = t.get("rows", [])
                info_lines.append(f"{title} ({len(rows)} dòng):\n")
                if cols == ["Trường", "Giá trị"]:
                    for r in rows:
                        if len(r) >= 2 and (r[0] or r[1]):
                            info_lines.append(f"  {r[0]}: {r[1]}\n")
                else:
                    info_lines.append("  Cột: " + " | ".join(cols) + "\n")
                    info_lines.append(f"  Số dòng: {len(rows)}\n")
                info_lines.append("\n")
            info_lines.append("Log: last_run_v12\\phieu_coc_tables.json\n")
            self.excel_info.delete("1.0", "end")
            self.excel_info.insert("1.0", "".join(info_lines))

        except Exception:
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "last_error_phieu_coc.txt").write_text(traceback.format_exc(), encoding="utf-8")
            messagebox.showerror("Lỗi đọc phiếu cọc", "Có lỗi. Xem last_run_v12/last_error_phieu_coc.txt")
            self.status.config(text="Lỗi đọc phiếu cọc.")


    def read_each_sheet_content(self):
        if not self.excel_path:
            messagebox.showwarning("Thiếu Excel", "Bạn chưa chọn file Excel.")
            return
        try:
            analysis = analyze_workbook_sheets(self.excel_path)
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "excel_each_sheet_content.json").write_text(json.dumps(analysis, ensure_ascii=False, indent=2), encoding="utf-8")

            self.excel_info.delete("1.0", "end")
            lines = []
            lines.append("ĐỌC TỪNG SHEET TRONG FILE EXCEL\n")
            lines.append("=" * 60 + "\n")
            lines.append(f"File: {self.excel_path}\n\n")

            for sh in analysis["sheets"]:
                lines.append(f"Sheet: {sh.get('sheet')}\n")
                if sh.get("empty"):
                    lines.append("- Sheet trống\n\n")
                    continue

                ur = sh.get("used_range", {})
                lines.append(f"- Loại: {sh.get('sheet_type')}\n")
                lines.append(f"- Vùng: {ur.get('from')} → {ur.get('to')} | Header: {sh.get('header_rows')} | Dữ liệu: {sh.get('data_row_count')} dòng\n")
                lines.append(f"- TỔNG: {sh.get('total_rows')} | Công thức: {sh.get('formula_count')} ô | Merge: {sh.get('merged_ranges_count')}\n")

                cols = [f"{h['col']}:{short_header_name(h['name'], 24)}" for h in sh.get("headers", [])[:30]]
                if cols:
                    lines.append("- Cột: " + " | ".join(cols) + "\n")

                if sh.get("formula_samples"):
                    fs = sh.get("formula_samples", [])[:8]
                    lines.append("- Công thức mẫu: " + " | ".join([f"{f['cell']}={f['formula']}" for f in fs]) + "\n")

                lines.append("\n")

            lines.append("Log đầy đủ: last_run_v12\\excel_each_sheet_content.json\n")
            self.excel_info.insert("1.0", "".join(lines))
            self.status.config(text="Đã đọc từng sheet trong file Excel.")
        except Exception:
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "last_error_each_sheet.txt").write_text(traceback.format_exc(), encoding="utf-8")
            messagebox.showerror("Lỗi đọc từng sheet", "Có lỗi. Xem last_run_v12/last_error_each_sheet.txt")
            self.status.config(text="Lỗi đọc từng sheet.")

    def read_current_excel_formulas(self):
        if not self.excel_path:
            messagebox.showwarning("Thiếu Excel", "Bạn chưa chọn file Excel.")
            return
        try:
            logic = read_formula_logic_for_workbook(self.excel_path)
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "excel_formula_logic.json").write_text(json.dumps(logic, ensure_ascii=False, indent=2), encoding="utf-8")

            self.excel_info.delete("1.0", "end")
            lines = []
            lines.append("ĐÃ ĐỌC CÔNG THỨC & CÁCH LÀM EXCEL\n")
            lines.append("=" * 70 + "\n")
            lines.append(f"File: {self.excel_path}\n\n")
            for sh in logic.get("sheets", []):
                lines.append(f"Sheet: {sh.get('sheet')}\n")
                if sh.get("error"):
                    lines.append(f"Lỗi: {sh.get('error')}\n\n")
                    continue
                lines.append(f"- Header row: {sh.get('header_row')}\n")
                lines.append(f"- Total row: {sh.get('total_row')}\n")
                lines.append(f"- STT col: {sh.get('stt_col')}\n")
                lines.append(f"- Số ô công thức: {sh.get('formula_count')}\n")
                if sh.get("formula_columns"):
                    lines.append("- Các cột có công thức:\n")
                    for fc in sh.get("formula_columns", [])[:30]:
                        lines.append(f"  + {fc['col']} ({fc.get('header','')}): {fc['count']} ô, mẫu {fc['sample_cell']} = {fc['sample_formula']}\n")
                if sh.get("total_formulas"):
                    lines.append("- Công thức dòng TỔNG:\n")
                    for tf in sh.get("total_formulas", [])[:30]:
                        lines.append(f"  + {tf['cell']} ({tf.get('header','')}): {tf['formula']}\n")
                if sh.get("rules_text"):
                    lines.append("- Cách làm tóm tắt:\n")
                    for rule in sh.get("rules_text", [])[:30]:
                        lines.append(f"  + {rule}\n")
                lines.append("\n")
            lines.append("Log đầy đủ: last_run_v12\\excel_formula_logic.json\n")
            self.excel_info.insert("1.0", "".join(lines))
            self.status.config(text="Đã đọc công thức và cách làm Excel.")
        except Exception:
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "last_error_formula_logic.txt").write_text(traceback.format_exc(), encoding="utf-8")
            messagebox.showerror("Lỗi đọc công thức", "Có lỗi. Xem last_run_v12/last_error_formula_logic.txt")
            self.status.config(text="Lỗi đọc công thức Excel.")

    def _profile_workbook(self, excel_path):
        """
        Đọc toàn bộ workbook: tất cả sheet, header, STT, TỔNG, chuỗi STT, công thức tổng.
        Không ghi gì vào file.
        """
        wb = load_workbook(excel_path, data_only=False)
        profiles = []
        for ws in wb.worksheets:
            try:
                header_row = find_header_row_smart(ws)
                headers = get_headers_smart(ws, header_row)
                total_row = find_total_row(ws, header_row)
                no_col = find_no_column_smart(ws, headers, header_row, total_row) if total_row else None
                chains = []
                best = None
                if total_row and no_col:
                    chains = find_all_stt_chains(ws, no_col, header_row, total_row)
                    best = select_longest_stt_chain(ws, no_col, header_row, total_row)

                formula_cols = []
                if total_row:
                    try:
                        formula_cols = capture_total_sum_columns(
                            ws,
                            total_row,
                            best[0][0] if best else None,
                            best[-1][0] if best else None
                        )
                    except Exception:
                        formula_cols = capture_formula_columns(ws, total_row)

                profiles.append({
                    "file": str(excel_path),
                    "sheet": ws.title,
                    "max_row": ws.max_row,
                    "max_col": ws.max_column,
                    "header_row": header_row,
                    "headers": [{"col": get_column_letter(c), "index": c, "name": name} for c, name in headers],
                    "total_row": total_row,
                    "stt_col": get_column_letter(no_col) if no_col else None,
                    "stt_col_index": no_col,
                    "stt_chains": [
                        {
                            "from_row": ch[0][0],
                            "to_row": ch[-1][0],
                            "from_stt": ch[0][1],
                            "to_stt": ch[-1][1],
                            "length": len(ch),
                        }
                        for ch in chains
                    ],
                    "selected_chain": {
                        "from_row": best[0][0],
                        "to_row": best[-1][0],
                        "from_stt": best[0][1],
                        "to_stt": best[-1][1],
                        "length": len(best),
                    } if best else None,
                    "sum_columns": [get_column_letter(c) for c in formula_cols],
                })
            except Exception as e:
                profiles.append({
                    "file": str(excel_path),
                    "sheet": ws.title,
                    "error": repr(e),
                })
        return profiles

    def _display_profiles(self, profiles, title="Kết quả đọc Excel"):
        self.excel_info.delete("1.0", "end")
        lines = [title + "\n", "=" * 50 + "\n"]
        for p in profiles:
            lines.append(f"\nFile: {p.get('file')}\n")
            lines.append(f"Sheet: {p.get('sheet')}\n")
            if p.get("error"):
                lines.append(f"LỖI: {p.get('error')}\n")
                continue
            lines.append(
                f"Header: {p.get('header_row')} | TỔNG: {p.get('total_row')} | STT: {p.get('stt_col')}\n"
            )
            sc = p.get("selected_chain")
            if sc:
                lines.append(
                    f"STT chọn: {sc['from_stt']} → {sc['to_stt']} "
                    f"(dòng {sc['from_row']} → {sc['to_row']}) | Tiếp: {sc['to_stt'] + 1}\n"
                )
            if p.get("sum_columns"):
                lines.append("SUM: " + ", ".join(p.get("sum_columns")) + "\n")
            headers = p.get("headers", [])
            if headers:
                lines.append("Cột: " + " | ".join([f"{h['col']}:{short_header_name(h['name'], 28)}" for h in headers[:25]]) + "\n")
        self.excel_info.insert("1.0", "".join(lines))

    def scan_current_workbook(self):
        if not self.excel_path:
            messagebox.showwarning("Thiếu Excel", "Bạn chưa chọn file Excel.")
            return
        try:
            profiles = self._profile_workbook(self.excel_path)
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "current_workbook_profiles.json").write_text(json.dumps(profiles, ensure_ascii=False, indent=2), encoding="utf-8")
            each_sheet = analyze_workbook_sheets(self.excel_path)
            (out / "excel_each_sheet_content.json").write_text(json.dumps(each_sheet, ensure_ascii=False, indent=2), encoding="utf-8")
            formula_logic = read_formula_logic_for_workbook(self.excel_path)
            (out / "excel_formula_logic.json").write_text(json.dumps(formula_logic, ensure_ascii=False, indent=2), encoding="utf-8")
            self._display_profiles(profiles, "Đã đọc toàn bộ workbook hiện tại")
            self.status.config(text="Đã đọc toàn bộ workbook.")
        except Exception:
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "last_error_scan_workbook.txt").write_text(traceback.format_exc(), encoding="utf-8")
            messagebox.showerror("Lỗi đọc workbook", "Có lỗi. Xem last_run_v12/last_error_scan_workbook.txt")
            self.status.config(text="Lỗi đọc workbook.")

    def scan_excel_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục chứa các file Excel")
        if not folder:
            return
        self.excel_folder = folder
        paths = []
        for ext in ("*.xlsx", "*.xlsm"):
            paths.extend(Path(folder).glob(ext))
        paths = [p for p in paths if not p.name.startswith("~$")]

        if not paths:
            messagebox.showwarning("Không có Excel", "Thư mục này không có file .xlsx/.xlsm.")
            return

        all_profiles = []
        errors = []
        for p in paths:
            try:
                all_profiles.extend(self._profile_workbook(p))
            except Exception as e:
                errors.append({"file": str(p), "error": repr(e)})

        out = app_dir() / "last_run_v12"
        out.mkdir(exist_ok=True)
        (out / "all_excel_profiles.json").write_text(json.dumps(all_profiles, ensure_ascii=False, indent=2), encoding="utf-8")
        if errors:
            (out / "all_excel_scan_errors.json").write_text(json.dumps(errors, ensure_ascii=False, indent=2), encoding="utf-8")

        self._display_profiles(all_profiles, f"Đã quét {len(paths)} file Excel trong thư mục")
        self.status.config(text=f"Đã quét {len(paths)} file Excel. Log: last_run_v12/all_excel_profiles.json")

    def choose_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xlsm"), ("All", "*.*")])
        if not p:
            return
        try:
            self.excel_path = p
            self.workbook = load_workbook(p, data_only=False)
            sheets = self.workbook.sheetnames
            self.sheet_combo["values"] = sheets

            # Đọc toàn workbook ngay khi chọn file
            profiles = self._profile_workbook(p)
            best = choose_best_sheet_profile(profiles)

            if best and best.get("sheet") in sheets:
                self.sheet_var.set(best["sheet"])
                self.sheet_combo.current(sheets.index(best["sheet"]))
            elif sheets:
                self.sheet_combo.current(0)
                self.sheet_var.set(sheets[0])

            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "current_workbook_profiles.json").write_text(json.dumps(profiles, ensure_ascii=False, indent=2), encoding="utf-8")

            self._display_profiles(profiles, "Đã đọc toàn bộ file Excel vừa chọn")
            self.status.config(text="Đã đọc toàn bộ file Excel và tự chọn sheet phù hợp.")
        except Exception:
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "last_error_open_excel.txt").write_text(traceback.format_exc(), encoding="utf-8")
            messagebox.showerror("Lỗi mở Excel", "Có lỗi khi đọc Excel. Xem last_run_v12/last_error_open_excel.txt")
            self.status.config(text="Lỗi mở Excel.")

    def refresh_excel_header_info(self):
        self.excel_info.delete("1.0", "end")
        if not self.workbook or not self.sheet_var.get():
            self.excel_info.insert("1.0", "Chưa chọn Excel hoặc sheet.\n")
            return

        try:
            ws = self.workbook[self.sheet_var.get()]

            self.header_row = find_header_row_smart(ws)
            self.excel_headers = get_headers_smart(ws, self.header_row)

            total_row = find_total_row(ws, self.header_row)
            no_col = find_no_column_smart(ws, self.excel_headers, self.header_row, total_row) if total_row else None

            txt = []
            txt.append(f"File: {self.excel_path}\n")
            txt.append(f"Sheet: {ws.title}\n")
            txt.append(f"Header: dòng {self.header_row}")
            if total_row:
                txt.append(f" | TỔNG: dòng {total_row}")
            if no_col:
                txt.append(f" | STT: cột {get_column_letter(no_col)}")
            txt.append("\n\n")

            txt.append("Cột phát hiện:\n")
            for col_idx, name in self.excel_headers:
                txt.append(f"- {get_column_letter(col_idx)}: {short_header_name(name)}\n")

            if total_row and no_col:
                try:
                    best = select_longest_stt_chain(ws, no_col, self.header_row, total_row)
                    if best:
                        txt.append(
                            f"\nChuỗi STT chọn: {best[0][1]} → {best[-1][1]} "
                            f"(dòng {best[0][0]} → {best[-1][0]}) | STT tiếp: {best[-1][1] + 1}\n"
                        )
                    else:
                        txt.append("\nChưa tìm thấy chuỗi STT chuẩn.\n")
                except Exception:
                    txt.append("\nKhông đọc được chuỗi STT.\n")

            self.excel_info.insert("1.0", "".join(txt))

        except Exception:
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "last_error_excel_info.txt").write_text(traceback.format_exc(), encoding="utf-8")
            self.excel_info.insert("1.0", "Lỗi đọc thông tin Excel. Xem last_run_v12/last_error_excel_info.txt\n")
            self.status.config(text="Lỗi đọc thông tin Excel.")

    def choose_image(self):
        p = filedialog.askopenfilename(filetypes=[("Image", "*.png;*.jpg;*.jpeg;*.bmp;*.webp"), ("All", "*.*")])
        if not p:
            return
        self.set_image_path(p, "Đã chọn ảnh: ")

    def set_image_path(self, path, status_prefix="Đã chọn ảnh: "):
        p = str(path)
        self.image_path = p
        im = Image.open(p)
        im.thumbnail((260, 360))
        self.tk_img = ImageTk.PhotoImage(im)
        self.img_label.config(image=self.tk_img, text="")
        self.status.config(text=status_prefix + p)

    def _clipboard_text_widget_has_focus(self):
        try:
            widget = self.root.focus_get()
        except Exception:
            return False
        if widget is None:
            return False
        try:
            cls = widget.winfo_class()
        except Exception:
            cls = ""
        return isinstance(widget, (tk.Entry, tk.Text, ttk.Entry)) or cls in {"Entry", "TEntry", "Text"}

    def paste_image_from_clipboard(self, event=None):
        if self._clipboard_text_widget_has_focus():
            return None
        try:
            from PIL import ImageGrab
            clip = ImageGrab.grabclipboard()
        except Exception:
            messagebox.showwarning("Clipboard", "Không đọc được clipboard ảnh trên máy này.")
            return "break"

        image = None
        source_path = None
        if isinstance(clip, Image.Image):
            image = clip.convert("RGB")
        elif isinstance(clip, list):
            for item in clip:
                try:
                    p = Path(item)
                    if p.exists() and p.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp", ".webp"}:
                        source_path = p
                        break
                except Exception:
                    pass

        if source_path:
            self.set_image_path(source_path, "Đã dán ảnh từ clipboard: ")
            return "break"

        if image is None:
            try:
                text = self.root.clipboard_get()
            except Exception:
                text = ""
            if str(text or "").strip():
                self.status.config(text="Clipboard đang là text. Click vào ô nhập rồi bấm Ctrl+V để dán text.")
            else:
                messagebox.showinfo(
                    "Clipboard",
                    "Clipboard chưa có ảnh. Hãy copy ảnh bằng Ctrl+C rồi bấm Ctrl+V trong ứng dụng."
                )
            return "break"

        try:
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            paste_path = out / "clipboard_paste.png"
            image.save(paste_path)
            self.set_image_path(paste_path, "Đã dán ảnh từ clipboard: ")
        except Exception:
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "last_error_clipboard.txt").write_text(traceback.format_exc(), encoding="utf-8")
            messagebox.showerror("Clipboard", "Không lưu được ảnh clipboard. Xem last_run_v12/last_error_clipboard.txt")
        return "break"

    def run_gemini(self):
        if not self.image_path:
            messagebox.showwarning("Thiếu ảnh", "Bạn chưa chọn ảnh.")
            return
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("Thiếu API key", "Bạn chưa nhập Gemini API key.")
            return

        self.save_key()
        self.status.config(text="Đang gửi ảnh lên Gemini...")
        self.root.update()

        try:
            tables, raw = call_gemini(self.image_path, api_key, self.model_var.get().strip())
            tables = postprocess_to_hop_coc_d1_d2(tables)
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "gemini_raw_response.txt").write_text(raw, encoding="utf-8")
            (out / "gemini_tables.json").write_text(json.dumps(tables, ensure_ascii=False, indent=2), encoding="utf-8")

            # V20: giữ nguyên cấu trúc bảng mà ảnh trả về, không ép mẫu cố định.
            self.tables = tables
            self.table_editor.set_tables(tables)
            self.build_mapping()

            total_rows = sum(len(t["rows"]) for t in tables)
            self.status.config(text=f"Gemini đọc xong: {len(tables)} bảng, {total_rows} dòng. Đã giữ cấu trúc đúng theo ảnh.")
        except Exception:
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "last_error.txt").write_text(traceback.format_exc(), encoding="utf-8")
            messagebox.showerror("Lỗi Gemini API", "Có lỗi. Xem last_run_v12/last_error.txt")
            self.status.config(text="Lỗi Gemini API")

    def build_mapping(self):
        table = self.table_editor.get_current_table()
        if not table:
            return
        if not self.excel_headers:
            messagebox.showwarning("Thiếu Excel", "Bạn chưa chọn Excel hoặc sheet.")
            return

        auto = auto_map_columns(table["columns"], self.excel_headers)
        auto = ensure_no_column_in_mapping(table["columns"], auto, self.excel_headers)
        self.mapping_editor.set_mapping(table["columns"], self.excel_headers, auto)
        self.status.config(text="Đã auto map từ bảng trong ảnh sang cột của file Excel.")



    def _merged_master_cell(self, ws, row, col):
        """
        Nếu ô đang ghi là MergedCell thì trả về ô góc trên-trái của vùng merge.
        Nếu không merge thì trả về chính ô đó.
        """
        cell = ws.cell(row, col)
        if cell.__class__.__name__ != "MergedCell":
            return cell
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                return ws.cell(rng.min_row, rng.min_col)
        return cell

    def _safe_set_cell_value(self, ws, row, col, value):
        """
        Ghi an toàn vào Excel.
        Tránh lỗi: MergedCell object attribute 'value' is read-only
        """
        cell = ws.cell(row, col)
        if cell.__class__.__name__ == "MergedCell":
            # Nếu ô đang nằm trong vùng merge, chỉ ghi khi nó là ô master.
            # Nếu không phải master thì bỏ qua để không làm vỡ form.
            master = self._merged_master_cell(ws, row, col)
            if master.__class__.__name__ == "MergedCell":
                return False
            master.value = value
            return True
        cell.value = value
        return True

    def _apply_rows_to_workbook(self, wb):
        """
        V20.5:
        - Excel bất kỳ: đọc header từ sheet đang chọn.
        - Nếu chưa có mapping thì tự map sau khi đã có excel_headers.
        - Ảnh quyết định cột nguồn, Excel quyết định cột đích/STT/TỔNG/SUM.
        """
        if not self.sheet_var.get():
            raise ValueError("Bạn chưa chọn sheet.")

        table = self.table_editor.get_current_table()
        if not table:
            raise ValueError("Chưa có dữ liệu Gemini.")

        ws = wb[self.sheet_var.get()]
        header_row = find_header_row_smart(ws)
        excel_headers = get_headers_smart(ws, header_row)
        self.header_row = header_row
        self.excel_headers = excel_headers

        mapping = self.mapping_editor.get_mapping()
        if not mapping:
            mapping = auto_mapping_to_excel_columns(table["columns"], excel_headers)
            try:
                auto_idx = auto_map_columns(table["columns"], excel_headers)
                auto_idx = ensure_no_column_in_mapping(table["columns"], auto_idx, excel_headers)
                self.mapping_editor.set_mapping(table["columns"], excel_headers, auto_idx)
            except Exception:
                pass

        if not mapping:
            raise ValueError("Chưa có mapping cột. Tool không tự map được vì Excel chưa có header rõ.")

        rows = table["rows"]

        total_row = find_total_row(ws, header_row)
        if not total_row:
            raise ValueError("Không tìm thấy dòng TỔNG/TOTAL trong Excel. Tool cần dòng TỔNG để biết chèn ở đâu.")

        no_col = find_no_column_smart(ws, excel_headers, header_row, total_row)
        if not no_col:
            raise ValueError("Không tìm thấy cột STT/No trong Excel.")

        chains = find_all_stt_chains(ws, no_col, header_row, total_row)
        best_chain = select_longest_stt_chain(ws, no_col, header_row, total_row)

        # Nếu không có chuỗi STT liên tục thì dùng fallback:
        # lấy dòng dữ liệu cuối trước TỔNG + số STT lớn nhất đang có.
        used_stt_fallback = False
        if best_chain:
            first_seq_row = best_chain[0][0]
            last_seq_row = best_chain[-1][0]
            last_no = best_chain[-1][1]
        else:
            used_stt_fallback = True
            last_seq_row = find_last_data_row_before_total(ws, header_row, total_row, mapping, no_col)
            if not last_seq_row:
                raise ValueError("Không tìm thấy dòng dữ liệu cuối trước dòng TỔNG. Kiểm tra lại sheet/header/dòng tổng.")
            first_seq_row = header_row + 1
            last_no = find_last_stt_number_loose(ws, no_col, header_row, total_row)

        # Cột cần SUM dựa trên dòng TỔNG mẫu trước khi sửa
        total_sum_cols_before = capture_total_sum_columns(ws, total_row, first_seq_row, last_seq_row)
        total_sum_cols_before = [c for c in total_sum_cols_before if c != no_col]

        # V22 FIX CHUẨN:
        # Không ghi vào vùng trống/merged có sẵn vì dễ phá form.
        # Luôn insert dòng mới ngay trước dòng TỔNG rồi mới nhập dữ liệu.
        garbage_count = 0

        insert_at = total_row
        ws.insert_rows(insert_at, amount=len(rows))

        total_row_after = total_row + len(rows)
        style_row = last_seq_row

        target_rows = list(range(insert_at, insert_at + len(rows)))

        for r_offset, row in enumerate(rows):
            dst_row = target_rows[r_offset]

            # Copy đúng format từ dòng dữ liệu mẫu
            copy_style_row(ws, style_row, dst_row, ws.max_column)
            try:
                copy_row_dimension(ws, style_row, dst_row)
            except Exception:
                pass

            # Copy công thức dòng mẫu nếu có
            try:
                apply_row_formulas_from_template(ws, style_row, dst_row)
            except Exception:
                pass

            # STT tự nối tiếp, không lấy từ ảnh
            self._safe_set_cell_value(ws, dst_row, no_col, last_no + r_offset + 1)

            # Ghi dữ liệu OCR theo mapping
            for src_idx, excel_col in enumerate(mapping):
                if excel_col is None:
                    continue
                if excel_col == no_col:
                    continue

                # Nếu ô đang có công thức thì giữ công thức
                if is_formula_value(ws.cell(dst_row, excel_col).value):
                    continue

                val = row[src_idx] if src_idx < len(row) else ""
                self._safe_set_cell_value(ws, dst_row, excel_col, convert_excel_value(val))

        sum_first_row = first_seq_row
        sum_last_row = target_rows[-1]
        set_total_formulas_by_template(
            ws,
            total_row_after,
            total_sum_cols_before,
            sum_first_row,
            sum_last_row
        )
        force_workbook_recalculate(wb)

        out = app_dir() / "last_run_v12"
        out.mkdir(exist_ok=True)
        logic = {
            "rule": "V20.5 auto map sau khi đọc excel_headers",
            "header_row": header_row,
            "no_col": no_col,
            "total_row_before": total_row,
            "total_row_after": total_row_after,
            "selected_chain": {
                "from_row": first_seq_row,
                "to_row": last_seq_row,
                "from_stt": (best_chain[0][1] if best_chain else None),
                "to_stt": last_no,
                "length": (len(best_chain) if best_chain else 0),
                "fallback": used_stt_fallback,
            },
            "insert_at_row": insert_at,
            "new_stt_start": last_no + 1,
            "new_stt_end": last_no + len(rows),
            "sum_range_rows": [sum_first_row, sum_last_row],
            "sum_columns": total_sum_cols_before,
            "mapping": [
                {
                    "source": table["columns"][i] if i < len(table["columns"]) else "",
                    "excel_col": excel_col,
                    "excel_letter": get_column_letter(excel_col) if excel_col else None
                }
                for i, excel_col in enumerate(mapping)
            ]
        }
        (out / "v20_5_apply_logic.json").write_text(json.dumps(logic, ensure_ascii=False, indent=2), encoding="utf-8")

        return {
            "sheet": ws.title,
            "header_row": header_row,
            "first_stt_row": first_seq_row,
            "last_stt_row": last_seq_row,
            "last_stt_before": last_no,
            "start_fill_row": insert_at,
            "next_stt_start": last_no + 1,
            "rows_added": len(rows),
            "garbage_deleted_count": garbage_count,
            "total_row_after": total_row_after,
            "sum_first_row": sum_first_row,
            "sum_last_row": sum_last_row,
            "sum_columns_count": len(total_sum_cols_before),
            "used_stt_fallback": used_stt_fallback,
        }

    def preview_excel(self):
        if not self.excel_path:
            messagebox.showwarning("Thiếu Excel", "Bạn chưa chọn file Excel.")
            return
        try:
            wb = load_workbook(self.excel_path)
            info = self._apply_rows_to_workbook(wb)

            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            preview_path = out / "preview_sau_khi_ghep.xlsx"
            force_workbook_recalculate(wb)
            wb.save(preview_path)

            try:
                os.startfile(str(preview_path))
            except Exception:
                pass

            self.status.config(text=f"Đã tạo file xem trước: {preview_path}")
            messagebox.showinfo(
                "Đã tạo xem trước",
                f"Đã tạo file preview để kiểm tra trước khi lưu thật.\n\n"
                f"Sheet: {info['sheet']}\n"
                f"Bắt đầu chèn từ dòng: {info['start_fill_row']}\n"
                f"Số dòng thêm: {info['rows_added']}\n\n"
                f"File preview:\n{preview_path}\n\n"
                f"Lưu ý: nếu STT bắt đầu không đúng, bạn đang chọn file đã bị ghi thử trước đó. Hãy dùng lại file Excel gốc sạch.\n\n"
                f"Nếu đúng thì quay lại tool bấm: Điền tiếp vào Excel."
            )
        except Exception:
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "last_error_preview.txt").write_text(traceback.format_exc(), encoding="utf-8")
            messagebox.showerror("Lỗi xem trước", "Có lỗi. Xem last_run_v12/last_error_preview.txt")
            self.status.config(text="Lỗi xem trước Excel.")

    def fill_excel(self):
        if not self.excel_path:
            messagebox.showwarning("Thiếu Excel", "Bạn chưa chọn file Excel.")
            return

        try:
            # Luôn nạp lại file gốc để preview không làm nhân đôi dữ liệu trong bộ nhớ
            wb = load_workbook(self.excel_path)
            info = self._apply_rows_to_workbook(wb)

            out_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile=Path(self.excel_path).stem + "_da_dien_tiep.xlsx"
            )
            if not out_path:
                return

            force_workbook_recalculate(wb)
            wb.save(out_path)

            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "last_fill_info.json").write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")

            self.status.config(text=f"Đã điền {info['rows_added']} dòng vào sheet {info['sheet']}, bắt đầu từ dòng {info['start_fill_row']}.")
            messagebox.showinfo(
                "Xong",
                f"Đã điền {info['rows_added']} dòng vào Excel.\n"
                f"Bắt đầu từ dòng {info['start_fill_row']}\n"
                f"File: {out_path}"
            )
        except Exception:
            out = app_dir() / "last_run_v12"
            out.mkdir(exist_ok=True)
            (out / "last_error_fill.txt").write_text(traceback.format_exc(), encoding="utf-8")
            messagebox.showerror("Lỗi điền Excel", "Có lỗi. Xem last_run_v12/last_error_fill.txt")
            self.status.config(text="Lỗi điền Excel.")



def _apply_rows_insert_before_total_chot(self, wb):
    """
    CHỐT:
    - Luôn insert dòng mới ngay trước dòng TỔNG.
    - Điền dữ liệu vào dòng vừa insert.
    - Dòng TỔNG và toàn bộ phần sau TỔNG bị đẩy xuống.
    - Không ghi vào vùng trắng/merged có sẵn.
    - SUM lại các cột cần tổng.
    """
    if not self.sheet_var.get():
        raise ValueError("Bạn chưa chọn sheet.")

    table = self.table_editor.get_current_table()
    if not table:
        raise ValueError("Chưa có dữ liệu Gemini.")

    ws = wb[self.sheet_var.get()]
    header_row = find_header_row_smart(ws)
    excel_headers = get_headers_smart(ws, header_row)

    self.header_row = header_row
    self.excel_headers = excel_headers

    total_row = find_total_row(ws, header_row)
    if not total_row:
        raise ValueError("Không tìm thấy dòng TỔNG/TOTAL trong Excel.")

    no_col = find_no_column_smart(ws, excel_headers, header_row, total_row)
    if not no_col:
        raise ValueError("Không tìm thấy cột STT/No trong Excel.")

    rows = table.get("rows", [])
    if not rows:
        raise ValueError("Không có dòng dữ liệu để nhập.")

    mapping = self.mapping_editor.get_mapping()
    if not mapping:
        mapping = auto_mapping_to_excel_columns(table["columns"], excel_headers)
        try:
            auto_idx = auto_map_columns(table["columns"], excel_headers)
            auto_idx = ensure_no_column_in_mapping(table["columns"], auto_idx, excel_headers)
            self.mapping_editor.set_mapping(table["columns"], excel_headers, auto_idx)
        except Exception:
            pass

    if not mapping:
        raise ValueError("Chưa có mapping cột.")

    # Lấy STT lớn nhất trước dòng TỔNG
    stt_nums = []
    memo = {}

    for r in range(header_row + 1, total_row):
        try:
            if row_has_grey_background(ws, r):
                continue
        except Exception:
            pass

        row_text = " ".join(str(ws.cell(r, c).value or "") for c in range(1, ws.max_column + 1))
        try:
            if is_total_marker_text(row_text):
                continue
        except Exception:
            pass

        n = None
        try:
            n = get_stt_value(ws, r, no_col, memo)
        except Exception:
            n = None

        if not isinstance(n, int):
            s = str(ws.cell(r, no_col).value or "").strip()
            if s.isdigit():
                n = int(s)

        if isinstance(n, int):
            stt_nums.append((r, n))

    if stt_nums:
        first_data_row = sorted(stt_nums, key=lambda x: x[0])[0][0]
        style_row, last_no = sorted(stt_nums, key=lambda x: (x[1], x[0]))[-1]
    else:
        best = select_longest_stt_chain(ws, no_col, header_row, total_row)
        if best:
            first_data_row = best[0][0]
            style_row = best[-1][0]
            last_no = best[-1][1]
        else:
            first_data_row = header_row + 1
            style_row = total_row - 1
            last_no = 0

    # Nhận diện cột cần SUM dựa trên dòng TỔNG mẫu trước khi insert
    sum_columns = []
    for c in range(1, ws.max_column + 1):
        if c == no_col:
            continue

        total_val = ws.cell(total_row, c).value

        if is_formula_value(total_val):
            sum_columns.append(c)
            continue

        if isinstance(total_val, (int, float)):
            for rr in range(first_data_row, total_row):
                vv = ws.cell(rr, c).value
                if isinstance(vv, (int, float)):
                    sum_columns.append(c)
                    break

    # CHỖ QUAN TRỌNG NHẤT:
    # Insert ngay trước dòng TỔNG, không ghi vào dòng trắng sẵn có
    insert_at = total_row
    row_count = len(rows)

    ws.insert_rows(insert_at, amount=row_count)

    # Sau insert, dòng TỔNG mới bị đẩy xuống
    total_row_after = total_row + row_count
    target_rows = list(range(insert_at, insert_at + row_count))

    for i, data_row in enumerate(rows):
        dst_row = target_rows[i]

        # Copy style từ dòng dữ liệu mẫu
        copy_style_row(ws, style_row, dst_row, ws.max_column)

        try:
            copy_row_dimension(ws, style_row, dst_row)
        except Exception:
            pass

        try:
            apply_row_formulas_from_template(ws, style_row, dst_row)
        except Exception:
            pass

        # STT nối tiếp
        self._safe_set_cell_value(ws, dst_row, no_col, last_no + i + 1)

        # Điền dữ liệu theo mapping
        for src_idx, excel_col in enumerate(mapping):
            if excel_col is None:
                continue
            if excel_col == no_col:
                continue

            # Ô công thức thì giữ công thức
            if is_formula_value(ws.cell(dst_row, excel_col).value):
                continue

            val = data_row[src_idx] if src_idx < len(data_row) else ""
            self._safe_set_cell_value(ws, dst_row, excel_col, convert_excel_value(val))

    # SUM lại các cột cần tổng
    sum_first_row = first_data_row
    sum_last_row = target_rows[-1]

    for c in sorted(set(sum_columns)):
        letter = get_column_letter(c)
        ws.cell(total_row_after, c).value = f"=SUM({letter}{sum_first_row}:{letter}{sum_last_row})"

    force_workbook_recalculate(wb)

    out = app_dir() / "last_run_v12"
    out.mkdir(exist_ok=True)

    logic = {
        "rule": "CHOT_INSERT_TRUOC_TONG",
        "sheet": ws.title,
        "header_row": header_row,
        "total_row_before": total_row,
        "insert_at": insert_at,
        "rows_added": row_count,
        "data_rows": target_rows,
        "total_row_after": total_row_after,
        "last_stt_before": last_no,
        "new_stt_start": last_no + 1,
        "new_stt_end": last_no + row_count,
        "sum_first_row": sum_first_row,
        "sum_last_row": sum_last_row,
        "sum_columns": sorted(set(sum_columns)),
    }

    (out / "chot_insert_truoc_tong_logic.json").write_text(
        json.dumps(logic, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    return {
        "sheet": ws.title,
        "header_row": header_row,
        "last_stt_before": last_no,
        "start_fill_row": insert_at,
        "next_stt_start": last_no + 1,
        "rows_added": row_count,
        "total_row_after": total_row_after,
        "sum_first_row": sum_first_row,
        "sum_last_row": sum_last_row,
        "sum_columns_count": len(set(sum_columns)),
    }




# =========================
# V22.9 FINAL OVERRIDE - GIỮ TOÀN BỘ CHỨC NĂNG CŨ, CHỈ CHỐT LOGIC XUẤT EXCEL
# =========================

def _v229_is_formula(v):
    try:
        return is_formula_value(v)
    except Exception:
        return isinstance(v, str) and v.startswith("=")


def _v229_merged_ranges_shift_for_insert(ws, insert_at, amount):
    """
    openpyxl insert_rows không tự dịch merged ranges.
    Hàm này lưu merge, unmerge, insert, rồi merge lại đúng vị trí.
    Nhờ vậy dòng TỔNG + chữ ký + form dưới TỔNG được đẩy xuống đúng, không vỡ layout.
    """
    old_ranges = []
    try:
        for rng in list(ws.merged_cells.ranges):
            old_ranges.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
        for rng in list(ws.merged_cells.ranges):
            try:
                ws.unmerge_cells(str(rng))
            except Exception:
                pass
    except Exception:
        old_ranges = []

    ws.insert_rows(insert_at, amount=amount)

    for min_row, min_col, max_row, max_col in old_ranges:
        if min_row >= insert_at:
            min_row += amount
            max_row += amount
        elif min_row < insert_at <= max_row:
            max_row += amount
        try:
            ws.merge_cells(
                start_row=min_row, start_column=min_col,
                end_row=max_row, end_column=max_col
            )
        except Exception:
            pass


def _v229_find_stt_before_total(ws, no_col, header_row, total_row):
    """
    Lấy STT lớn nhất thật sự trước dòng TỔNG.
    Không lấy vùng sau TỔNG/chữ ký. Đọc được STT số và STT công thức =A16+1.
    """
    nums = []
    memo = {}
    for r in range(header_row + 1, total_row):
        try:
            if row_has_grey_background(ws, r):
                continue
        except Exception:
            pass
        try:
            row_text = " ".join(str(ws.cell(r, c).value or "") for c in range(1, ws.max_column + 1))
            if is_total_marker_text(row_text):
                continue
        except Exception:
            pass

        n = None
        try:
            n = get_stt_value(ws, r, no_col, memo)
        except Exception:
            n = None
        if not isinstance(n, int):
            s = str(ws.cell(r, no_col).value or "").strip()
            if s.isdigit():
                n = int(s)
            else:
                try:
                    f = float(s.replace(",", "."))
                    if f.is_integer():
                        n = int(f)
                except Exception:
                    pass
        if isinstance(n, int):
            nums.append((r, n))

    if nums:
        first_row = sorted(nums, key=lambda x: x[0])[0][0]
        style_row, last_no = sorted(nums, key=lambda x: (x[1], x[0]))[-1]
        return first_row, style_row, last_no

    # fallback giữ chức năng cũ
    best = None
    try:
        best = select_longest_stt_chain(ws, no_col, header_row, total_row)
    except Exception:
        best = None
    if best:
        return best[0][0], best[-1][0], best[-1][1]

    last_row = None
    try:
        last_row = find_last_data_row_before_total(ws, header_row, total_row, None, no_col)
    except Exception:
        last_row = None
    return header_row + 1, (last_row or total_row - 1), 0


def _v229_capture_sum_columns(ws, total_row, first_data_row, last_data_row, no_col):
    """
    Chỉ SUM các cột mà file Excel mẫu cần tổng:
    - ô dòng TỔNG có công thức
    - hoặc ô dòng TỔNG là số và phía trên có dữ liệu số
    Không SUM cột STT.
    """
    cols = set()
    for c in range(1, ws.max_column + 1):
        if c == no_col:
            continue
        v = ws.cell(total_row, c).value
        if _v229_is_formula(v):
            cols.add(c)
            continue
        is_num_total = False
        if isinstance(v, (int, float)):
            is_num_total = True
        else:
            sv = str(v or "").strip().replace(".", "").replace(",", ".")
            try:
                if sv != "":
                    float(sv)
                    is_num_total = True
            except Exception:
                is_num_total = False
        if is_num_total:
            for rr in range(first_data_row, max(first_data_row, last_data_row) + 1):
                vv = ws.cell(rr, c).value
                if isinstance(vv, (int, float)):
                    cols.add(c)
                    break
                svv = str(vv or "").strip().replace(".", "").replace(",", ".")
                try:
                    if svv != "":
                        float(svv)
                        cols.add(c)
                        break
                except Exception:
                    pass
    return sorted(cols)


def _v229_normalize_mapping_to_excel_columns(source_cols, mapping, excel_headers):
    """
    Chấp nhận cả mapping dạng index combobox và dạng cột Excel thật.
    Trả về list source_idx -> excel column number.
    """
    out = []
    header_cols = [c for c, _ in excel_headers]
    max_col = max(header_cols) if header_cols else 0
    for m in (mapping or []):
        if m is None:
            out.append(None)
            continue
        try:
            mi = int(m)
        except Exception:
            out.append(None)
            continue
        # Nếu mi là index trong excel_headers
        if 0 <= mi < len(excel_headers) and mi not in header_cols:
            out.append(excel_headers[mi][0])
        # Nếu mi là cột Excel thật
        elif 1 <= mi <= max_col:
            out.append(mi)
        elif 0 <= mi < len(excel_headers):
            out.append(excel_headers[mi][0])
        else:
            out.append(None)

    # kéo dài mapping nếu thiếu
    while len(out) < len(source_cols):
        out.append(None)

    # STT/No trong ảnh luôn bỏ qua, tool tự nối STT theo Excel
    try:
        for i, src in enumerate(source_cols):
            if is_no_header(src):
                out[i] = None
    except Exception:
        pass
    return out[:len(source_cols)]


def _v229_safe_copy_row(ws, src_row, dst_row):
    try:
        copy_style_row(ws, src_row, dst_row, ws.max_column)
    except Exception:
        pass
    try:
        copy_row_dimension(ws, src_row, dst_row)
    except Exception:
        pass
    try:
        apply_row_formulas_from_template(ws, src_row, dst_row)
    except Exception:
        pass


def _v229_filter_data_rows(rows):
    """Bỏ dòng rỗng hoàn toàn trong preview, không đụng dòng có dữ liệu."""
    out = []
    for r in rows or []:
        vals = list(r) if isinstance(r, (list, tuple)) else [r]
        if any(str(x or "").strip() for x in vals):
            out.append(vals)
    return out


def postprocess_to_hop_coc_d1_d2(tables):
    """
    Bổ sung/ghi đè nhẹ rule Tổ hợp cọc:
    - Nếu Gemini đã trả D1/D2 thì giữ nguyên.
    - Nếu còn cột 'Tổ hợp cọc' + cột số kế bên thì đổi header thành D1/D2, không làm lệch dữ liệu.
    - Nếu ô tổ hợp chứa '6 10' hoặc '6|10' thì tách thành D1/D2.
    """
    if not tables:
        return tables

    def _n(x):
        try:
            return norm(x)
        except Exception:
            return str(x or "").lower().strip()

    def _looks_number(x):
        s = str(x or "").strip().replace(",", ".")
        if s.startswith("+") or s.startswith("-"):
            s = s[1:]
        return bool(re.fullmatch(r"\d+(?:\.\d+)?", s))

    for t in tables:
        cols = list(t.get("columns", []))
        rows = [list(r) for r in t.get("rows", [])]
        if not cols:
            continue
        ncols = [_n(c) for c in cols]
        if any(x in {"d1", "đ1", "1st"} for x in ncols) and any(x in {"d2", "đ2", "2nd"} for x in ncols):
            t["rows"] = rows
            continue
        idx = None
        for i, c in enumerate(ncols):
            if "to hop coc" in c or c == "to hop" or "pile combination" in c:
                idx = i
                break
        if idx is None:
            t["rows"] = rows
            continue

        # Case A: cột tổ hợp chứa 2 số trong 1 ô
        split_needed = False
        for r in rows[:15]:
            if idx < len(r):
                s = str(r[idx] or "").strip()
                nums = re.findall(r"[-+]?\d+(?:[,.]\d+)?", s)
                if len(nums) >= 2:
                    split_needed = True
                    break
        if split_needed:
            new_cols = cols[:]
            new_cols[idx] = "D1"
            new_cols.insert(idx + 1, "D2")
            new_rows = []
            for r in rows:
                rr = list(r)
                s = str(rr[idx] if idx < len(rr) else "").strip()
                nums = re.findall(r"[-+]?\d+(?:[,.]\d+)?", s)
                if len(nums) >= 2:
                    rr[idx] = nums[0]
                    rr.insert(idx + 1, nums[1])
                else:
                    rr.insert(idx + 1, "")
                if len(rr) < len(new_cols):
                    rr += [""] * (len(new_cols) - len(rr))
                new_rows.append(rr[:len(new_cols)])
            t["columns"] = new_cols
            t["rows"] = new_rows
            continue

        # Case B: header thiếu D2 nhưng dữ liệu đã có cột số kế bên
        if idx + 1 < len(cols):
            sample = 0
            ok = 0
            for r in rows[:15]:
                if idx + 1 < len(r):
                    a = r[idx]
                    b = r[idx + 1]
                    if str(a or "").strip() or str(b or "").strip():
                        sample += 1
                        if _looks_number(a) and _looks_number(b):
                            ok += 1
            if sample and ok >= max(1, sample // 2):
                cols[idx] = "D1"
                # Không insert value, chỉ đổi tên header cột kế bên thành D2 để giữ đúng cột dữ liệu hiện có.
                cols[idx + 1] = "D2"
                t["columns"] = cols
                t["rows"] = rows
                continue
        t["rows"] = rows
    return tables


def _v229_apply_rows_to_workbook(self, wb):
    """
    Logic chốt:
    1. Tìm dòng TỔNG chính.
    2. Insert đúng số dòng mới ngay trước dòng TỔNG.
    3. Dòng TỔNG + toàn bộ phần sau TỔNG + merged ranges được đẩy xuống.
    4. Điền dữ liệu vào đúng dòng vừa insert.
    5. Copy style/công thức dòng mẫu; SUM lại đúng cột cần SUM theo file mẫu.
    """
    if not self.sheet_var.get():
        raise ValueError("Bạn chưa chọn sheet.")
    table = self.table_editor.get_current_table()
    if not table:
        raise ValueError("Chưa có dữ liệu Gemini.")

    # Đảm bảo bảng nguồn đã được xử lý Tổ hợp cọc
    fixed_tables = postprocess_to_hop_coc_d1_d2([table])
    table = fixed_tables[0] if fixed_tables else table

    ws = wb[self.sheet_var.get()]
    header_row = find_header_row_smart(ws)
    excel_headers = get_headers_smart(ws, header_row)
    self.header_row = header_row
    self.excel_headers = excel_headers

    total_row = find_total_row(ws, header_row)
    if not total_row:
        raise ValueError("Không tìm thấy dòng TỔNG/TOTAL trong Excel.")

    no_col = find_no_column_smart(ws, excel_headers, header_row, total_row)
    if not no_col:
        raise ValueError("Không tìm thấy cột STT/No trong Excel.")

    source_cols = list(table.get("columns", []))
    rows = _v229_filter_data_rows(table.get("rows", []))
    if not rows:
        raise ValueError("Không có dòng dữ liệu để nhập.")

    raw_mapping = None
    try:
        raw_mapping = self.mapping_editor.get_mapping()
    except Exception:
        raw_mapping = None
    if not raw_mapping or all(x is None for x in raw_mapping):
        raw_mapping = auto_mapping_to_excel_columns(source_cols, excel_headers)
        try:
            auto_idx = auto_map_columns(source_cols, excel_headers)
            auto_idx = ensure_no_column_in_mapping(source_cols, auto_idx, excel_headers)
            self.mapping_editor.set_mapping(source_cols, excel_headers, auto_idx)
        except Exception:
            pass

    mapping = _v229_normalize_mapping_to_excel_columns(source_cols, raw_mapping, excel_headers)
    if not mapping or all(x is None for x in mapping):
        raise ValueError("Chưa có mapping cột hoặc mapping đang bỏ qua toàn bộ cột.")

    first_data_row, style_row, last_no = _v229_find_stt_before_total(ws, no_col, header_row, total_row)
    sum_columns = _v229_capture_sum_columns(ws, total_row, first_data_row, total_row - 1, no_col)

    insert_at = total_row
    row_count = len(rows)

    # Chèn dòng và dịch merged ranges để giữ nguyên form phía dưới
    _v229_merged_ranges_shift_for_insert(ws, insert_at, row_count)
    total_row_after = total_row + row_count
    target_rows = list(range(insert_at, insert_at + row_count))

    for i, data_row in enumerate(rows):
        dst_row = target_rows[i]
        _v229_safe_copy_row(ws, style_row, dst_row)

        # STT tự nối tiếp theo Excel, không lấy STT từ ảnh
        self._safe_set_cell_value(ws, dst_row, no_col, last_no + i + 1)

        for src_idx, excel_col in enumerate(mapping):
            if excel_col is None:
                continue
            if excel_col == no_col:
                continue
            try:
                if src_idx < len(source_cols) and is_no_header(source_cols[src_idx]):
                    continue
            except Exception:
                pass

            # Nếu dòng mẫu đã copy công thức vào ô này thì giữ công thức
            try:
                if _v229_is_formula(ws.cell(dst_row, excel_col).value):
                    continue
            except Exception:
                pass

            val = data_row[src_idx] if src_idx < len(data_row) else ""
            self._safe_set_cell_value(ws, dst_row, excel_col, convert_excel_value(val))

    # SUM lại đúng các cột cần tổng theo file mẫu
    sum_first_row = first_data_row
    sum_last_row = target_rows[-1]
    for c in sorted(set(sum_columns)):
        if c == no_col:
            continue
        try:
            letter = get_column_letter(c)
            ws.cell(total_row_after, c).value = f"=SUM({letter}{sum_first_row}:{letter}{sum_last_row})"
        except Exception:
            pass

    force_workbook_recalculate(wb)

    out = app_dir() / "last_run_v12"
    out.mkdir(exist_ok=True)
    logic = {
        "rule": "V22.9_FINAL_INSERT_BEFORE_TOTAL_KEEP_ALL_FEATURES",
        "sheet": ws.title,
        "header_row": header_row,
        "total_row_before": total_row,
        "insert_at": insert_at,
        "rows_added": row_count,
        "data_rows": target_rows,
        "total_row_after": total_row_after,
        "last_stt_before": last_no,
        "new_stt_start": last_no + 1,
        "new_stt_end": last_no + row_count,
        "sum_first_row": sum_first_row,
        "sum_last_row": sum_last_row,
        "sum_columns": sorted(set(sum_columns)),
        "mapping": [
            {
                "source": source_cols[i] if i < len(source_cols) else "",
                "excel_col": c,
                "excel_letter": get_column_letter(c) if c else None,
            }
            for i, c in enumerate(mapping)
        ],
    }
    (out / "v22_9_final_apply_logic.json").write_text(
        json.dumps(logic, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    return {
        "sheet": ws.title,
        "header_row": header_row,
        "last_stt_before": last_no,
        "start_fill_row": insert_at,
        "next_stt_start": last_no + 1,
        "rows_added": row_count,
        "total_row_after": total_row_after,
        "sum_first_row": sum_first_row,
        "sum_last_row": sum_last_row,
        "sum_columns_count": len(set(sum_columns)),
        "logic_file": str(out / "v22_9_final_apply_logic.json"),
    }


def _v229_preview_excel(self):
    """Preview không ghi đè file cũ, tránh PermissionError khi Excel đang mở."""
    if not self.excel_path:
        messagebox.showwarning("Thiếu Excel", "Bạn chưa chọn file Excel.")
        return
    try:
        from datetime import datetime
        wb = load_workbook(self.excel_path)
        info = self._apply_rows_to_workbook(wb)

        out = app_dir() / "last_run_v12"
        out.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        preview_path = out / f"preview_sau_khi_ghep_{stamp}.xlsx"
        force_workbook_recalculate(wb)
        wb.save(preview_path)

        try:
            os.startfile(str(preview_path))
        except Exception:
            pass

        self.status.config(text=f"Đã tạo file xem trước: {preview_path}")
        messagebox.showinfo(
            "Đã tạo xem trước",
            f"Đã tạo file preview để kiểm tra trước khi lưu thật.\n\n"
            f"Sheet: {info['sheet']}\n"
            f"Bắt đầu chèn từ dòng: {info['start_fill_row']}\n"
            f"Số dòng thêm: {info['rows_added']}\n"
            f"Dòng TỔNG sau khi đẩy xuống: {info['total_row_after']}\n\n"
            f"File preview:\n{preview_path}\n\n"
            f"Nếu đúng thì quay lại tool bấm: Điền tiếp vào Excel."
        )
    except Exception:
        out = app_dir() / "last_run_v12"
        out.mkdir(exist_ok=True)
        (out / "last_error_preview.txt").write_text(traceback.format_exc(), encoding="utf-8")
        messagebox.showerror("Lỗi xem trước", "Có lỗi. Xem last_run_v12/last_error_preview.txt")
        self.status.config(text="Lỗi xem trước Excel.")


# Ghi đè nhẹ, không xoá chức năng cũ.
App._apply_rows_to_workbook = _v229_apply_rows_to_workbook
App.preview_excel = _v229_preview_excel



# =========================
# V23 STRICT - DỮ LIỆU THEO ẢNH + TỔNG THEO FILE MẪU
# =========================

def _v23_cell_to_text(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        s = ("%s" % v).rstrip("0").rstrip(".")
        return s.replace(".", ",")
    return str(v).strip()


def _v23_norm_compare(v):
    s = _v23_cell_to_text(v).strip()
    s = s.replace(" ", "")
    s = s.replace(".", ",")
    # +1,5 trong ảnh và 1,5 trong Excel coi là cùng dữ liệu số dương
    if s.startswith("+"):
        s = s[1:]
    return s.lower()


def _v23_source_should_be_text(src_name):
    n = norm(src_name)
    return any(x in n for x in [
        "ngay", "date", "gio", "time", "bat dau", "ket thuc",
        "ten", "coc", "pile", "loai", "type", "vi tri", "location",
        "ghi chu", "note", "remark"
    ])


def _v23_convert_by_source(src_name, value):
    """
    Giữ nguyên dữ liệu dễ sai format; chỉ chuyển số cho cột cần tính toán.
    """
    s = str(value or "").strip()
    if s == "":
        return ""
    n = norm(src_name)
    # các cột này phải giữ đúng như ảnh
    if any(x in n for x in ["ngay", "date", "gio", "time", "bat dau", "ket thuc", "ten", "pile", "loai", "type", "vi tri", "ghi chu", "note", "remark"]):
        return s
    return convert_excel_value(s)


def _v23_validate_written_cells(ws, target_rows, rows, source_cols, mapping, no_col):
    """
    Log kiểm tra: ô nào ghi trực tiếp thì so lại với dữ liệu preview.
    Không chặn save vì Excel có thể định dạng số khác ảnh, nhưng log rõ để kiểm tra.
    """
    report = []
    for r_i, dst_row in enumerate(target_rows):
        data_row = rows[r_i]
        for src_idx, excel_col in enumerate(mapping):
            if excel_col is None or excel_col == no_col:
                continue
            src = source_cols[src_idx] if src_idx < len(source_cols) else ""
            img_val = data_row[src_idx] if src_idx < len(data_row) else ""
            cell = ws.cell(dst_row, excel_col)
            if _v229_is_formula(cell.value):
                status = "FORMULA_GIU_THEO_EXCEL_MAU"
            else:
                ex_val = cell.value
                status = "OK" if _v23_norm_compare(img_val) == _v23_norm_compare(ex_val) else "CHECK_FORMAT_OR_VALUE"
            report.append({
                "row": dst_row,
                "cell": f"{get_column_letter(excel_col)}{dst_row}",
                "source": src,
                "image_value": img_val,
                "excel_value": cell.value,
                "status": status,
            })
    return report


def _v23_apply_rows_to_workbook(self, wb):
    """
    V23 chốt:
    - Dữ liệu nguồn giữ theo ảnh ở preview.
    - Khi xuất Excel: chỉ ghi vào dòng mới insert trước TỔNG.
    - Dòng TỔNG và phần sau TỔNG đẩy xuống.
    - Cột có công thức ở dòng dữ liệu giữ công thức mẫu.
    - Dòng TỔNG chỉ SUM các cột mẫu đang SUM/cần tổng; không SUM STT.
    - Có log kiểm tra từng ô ghi trực tiếp.
    """
    if not self.sheet_var.get():
        raise ValueError("Bạn chưa chọn sheet.")
    table = self.table_editor.get_current_table()
    if not table:
        raise ValueError("Chưa có dữ liệu Gemini.")

    fixed_tables = postprocess_to_hop_coc_d1_d2([table])
    table = fixed_tables[0] if fixed_tables else table

    ws = wb[self.sheet_var.get()]
    header_row = find_header_row_smart(ws)
    excel_headers = get_headers_smart(ws, header_row)
    self.header_row = header_row
    self.excel_headers = excel_headers

    total_row = find_total_row(ws, header_row)
    if not total_row:
        raise ValueError("Không tìm thấy dòng TỔNG/TOTAL trong Excel.")

    no_col = find_no_column_smart(ws, excel_headers, header_row, total_row)
    if not no_col:
        raise ValueError("Không tìm thấy cột STT/No trong Excel.")

    source_cols = list(table.get("columns", []))
    rows = _v229_filter_data_rows(table.get("rows", []))
    if not rows:
        raise ValueError("Không có dòng dữ liệu để nhập.")

    raw_mapping = None
    try:
        raw_mapping = self.mapping_editor.get_mapping()
    except Exception:
        raw_mapping = None
    if not raw_mapping or all(x is None for x in raw_mapping):
        raw_mapping = auto_mapping_to_excel_columns(source_cols, excel_headers)
        try:
            auto_idx = auto_map_columns(source_cols, excel_headers)
            auto_idx = ensure_no_column_in_mapping(source_cols, auto_idx, excel_headers)
            self.mapping_editor.set_mapping(source_cols, excel_headers, auto_idx)
        except Exception:
            pass

    mapping = _v229_normalize_mapping_to_excel_columns(source_cols, raw_mapping, excel_headers)
    if not mapping or all(x is None for x in mapping):
        raise ValueError("Chưa có mapping cột hoặc mapping đang bỏ qua toàn bộ cột.")

    # Check cột quan trọng không bị bỏ qua hết
    mapped_names = {norm(source_cols[i]) for i, c in enumerate(mapping) if c and i < len(source_cols)}
    critical_any = ["d1", "d2", "loai coc", "ten tim coc", "ten coc"]
    # Không chặn cứng mọi form, nhưng log cảnh báo nếu thiếu
    warnings = []
    for key in critical_any:
        if not any(key in m or m in key for m in mapped_names):
            warnings.append(f"Có thể chưa map cột: {key}")

    first_data_row, style_row, last_no = _v229_find_stt_before_total(ws, no_col, header_row, total_row)
    sum_columns = _v229_capture_sum_columns(ws, total_row, first_data_row, total_row - 1, no_col)

    insert_at = total_row
    row_count = len(rows)

    _v229_merged_ranges_shift_for_insert(ws, insert_at, row_count)
    total_row_after = total_row + row_count
    target_rows = list(range(insert_at, insert_at + row_count))

    for i, data_row in enumerate(rows):
        dst_row = target_rows[i]
        _v229_safe_copy_row(ws, style_row, dst_row)

        self._safe_set_cell_value(ws, dst_row, no_col, last_no + i + 1)

        for src_idx, excel_col in enumerate(mapping):
            if excel_col is None or excel_col == no_col:
                continue
            try:
                if src_idx < len(source_cols) and is_no_header(source_cols[src_idx]):
                    continue
            except Exception:
                pass
            try:
                if _v229_is_formula(ws.cell(dst_row, excel_col).value):
                    continue
            except Exception:
                pass
            src_name = source_cols[src_idx] if src_idx < len(source_cols) else ""
            val = data_row[src_idx] if src_idx < len(data_row) else ""
            self._safe_set_cell_value(ws, dst_row, excel_col, _v23_convert_by_source(src_name, val))

    sum_first_row = first_data_row
    sum_last_row = target_rows[-1]
    for c in sorted(set(sum_columns)):
        if c == no_col:
            continue
        try:
            letter = get_column_letter(c)
            ws.cell(total_row_after, c).value = f"=SUM({letter}{sum_first_row}:{letter}{sum_last_row})"
        except Exception:
            pass

    force_workbook_recalculate(wb)

    validation = _v23_validate_written_cells(ws, target_rows, rows, source_cols, mapping, no_col)

    out = app_dir() / "last_run_v12"
    out.mkdir(exist_ok=True)
    logic = {
        "rule": "V23_CHUAN_DU_LIEU_TONG_CHINH_XAC",
        "sheet": ws.title,
        "header_row": header_row,
        "total_row_before": total_row,
        "insert_at": insert_at,
        "rows_added": row_count,
        "data_rows": target_rows,
        "total_row_after": total_row_after,
        "last_stt_before": last_no,
        "new_stt_start": last_no + 1,
        "new_stt_end": last_no + row_count,
        "sum_first_row": sum_first_row,
        "sum_last_row": sum_last_row,
        "sum_columns": sorted(set(sum_columns)),
        "warnings": warnings,
        "mapping": [
            {
                "source": source_cols[i] if i < len(source_cols) else "",
                "excel_col": c,
                "excel_letter": get_column_letter(c) if c else None,
            }
            for i, c in enumerate(mapping)
        ],
        "written_cell_validation": validation,
    }
    (out / "v23_chuan_du_lieu_tong_logic.json").write_text(
        json.dumps(logic, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    return {
        "sheet": ws.title,
        "header_row": header_row,
        "last_stt_before": last_no,
        "start_fill_row": insert_at,
        "next_stt_start": last_no + 1,
        "rows_added": row_count,
        "total_row_after": total_row_after,
        "sum_first_row": sum_first_row,
        "sum_last_row": sum_last_row,
        "sum_columns_count": len(set(sum_columns)),
        "logic_file": str(out / "v23_chuan_du_lieu_tong_logic.json"),
    }


def _v23_run_gemini(self):
    if not self.image_path:
        messagebox.showwarning("Thiếu ảnh", "Bạn chưa chọn ảnh.")
        return
    api_key = self.api_key_var.get().strip()
    if not api_key:
        messagebox.showwarning("Thiếu API key", "Bạn chưa nhập Gemini API key.")
        return
    try:
        tables, raw = call_gemini(self.image_path, api_key, self.model_var.get().strip())
        tables = postprocess_to_hop_coc_d1_d2(tables)
        out = app_dir() / "last_run_v12"
        out.mkdir(exist_ok=True)
        (out / "gemini_raw_response.txt").write_text(raw, encoding="utf-8")
        (out / "gemini_tables.json").write_text(json.dumps(tables, ensure_ascii=False, indent=2), encoding="utf-8")
        self.tables = tables
        self.table_editor.set_tables(tables)
        if self.excel_headers and tables:
            self.build_mapping()
        self.status.config(text=f"Gemini đọc xong: {len(tables)} bảng. Kiểm tra từng ô trong preview trước khi xuất.")
    except Exception:
        out = app_dir() / "last_run_v12"
        out.mkdir(exist_ok=True)
        (out / "last_error.txt").write_text(traceback.format_exc(), encoding="utf-8")
        messagebox.showerror("Lỗi Gemini API", "Có lỗi. Xem last_run_v12/last_error.txt")
        self.status.config(text="Lỗi Gemini API.")


# Override cuối cùng cho bản V23
App._apply_rows_to_workbook = _v23_apply_rows_to_workbook
App.run_gemini = _v23_run_gemini



# =========================
# V23.1 FINAL SUM FIX - SUM HẾT DỮ LIỆU TRƯỚC DÒNG TỔNG
# =========================

def _v231_parse_sum_first_row(formula, col_letter):
    """Lấy dòng bắt đầu từ công thức SUM mẫu ở dòng TỔNG, ví dụ =SUM(K9:K28) -> 9."""
    s = str(formula or "").replace(" ", "")
    if not s.startswith("="):
        return None
    import re
    col = str(col_letter).upper()
    m = re.search(r"SUM\(\$?" + re.escape(col) + r"\$?(\d+):\$?" + re.escape(col) + r"\$?\d+\)", s, flags=re.I)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return None
    m = re.search(r"\$?" + re.escape(col) + r"\$?(\d+):\$?" + re.escape(col) + r"\$?\d+", s, flags=re.I)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return None
    return None


def _v231_detect_first_data_row_strict(ws, header_row, total_row, no_col):
    """Tìm dòng dữ liệu đầu tiên thật sự trước TỔNG, ưu tiên STT số/công thức STT."""
    memo = {}
    for r in range(header_row + 1, total_row):
        try:
            row_text = " ".join(str(ws.cell(r, c).value or "") for c in range(1, ws.max_column + 1))
            if is_total_marker_text(row_text):
                continue
        except Exception:
            pass
        n = None
        try:
            n = get_stt_value(ws, r, no_col, memo)
        except Exception:
            n = None
        if isinstance(n, int):
            return r
        s = str(ws.cell(r, no_col).value or "").strip()
        if s.isdigit():
            return r
    return header_row + 1


def _v231_capture_sum_columns_and_starts(ws, total_row, default_first_row, no_col):
    """
    Chốt cột cần SUM và dòng bắt đầu SUM.
    - Nếu dòng TỔNG có công thức SUM: lấy dòng bắt đầu theo công thức mẫu.
    - Nếu dòng TỔNG là số: SUM từ dòng dữ liệu đầu tiên.
    - Không SUM cột STT.
    """
    result = {}
    for c in range(1, ws.max_column + 1):
        if c == no_col:
            continue
        letter = get_column_letter(c)
        v = ws.cell(total_row, c).value
        if _v229_is_formula(v):
            first = _v231_parse_sum_first_row(v, letter) or default_first_row
            result[c] = first
            continue
        is_num_total = False
        if isinstance(v, (int, float)):
            is_num_total = True
        else:
            sv = str(v or "").strip().replace(".", "").replace(",", ".")
            try:
                if sv != "":
                    float(sv)
                    is_num_total = True
            except Exception:
                is_num_total = False
        if not is_num_total:
            continue
        # Có số phía trên thì đây là cột tổng cần SUM.
        for rr in range(default_first_row, total_row):
            vv = ws.cell(rr, c).value
            if isinstance(vv, (int, float)):
                result[c] = default_first_row
                break
            svv = str(vv or "").strip().replace(".", "").replace(",", ".")
            try:
                if svv != "":
                    float(svv)
                    result[c] = default_first_row
                    break
            except Exception:
                pass
    return result


def _v231_apply_rows_to_workbook(self, wb):
    """
    V23.1 chốt SUM:
    - Insert dữ liệu mới ngay trước dòng TỔNG.
    - Dòng TỔNG và toàn bộ dòng sau TỔNG tự đẩy xuống.
    - Dòng TỔNG SUM từ dòng đầu dữ liệu theo công thức mẫu đến ngay trước dòng TỔNG mới.
    - Không dừng SUM ở dòng mới cuối; vì trước TỔNG có thể còn dòng trắng/form, vẫn phải SUM hết.
    """
    if not self.sheet_var.get():
        raise ValueError("Bạn chưa chọn sheet.")
    table = self.table_editor.get_current_table()
    if not table:
        raise ValueError("Chưa có dữ liệu Gemini.")

    fixed_tables = postprocess_to_hop_coc_d1_d2([table])
    table = fixed_tables[0] if fixed_tables else table

    ws = wb[self.sheet_var.get()]
    header_row = find_header_row_smart(ws)
    excel_headers = get_headers_smart(ws, header_row)
    self.header_row = header_row
    self.excel_headers = excel_headers

    total_row = find_total_row(ws, header_row)
    if not total_row:
        # Fallback: tìm dòng trống đầu tiên sau dữ liệu
        for r in range(header_row + 1, ws.max_row + 2):
            if not any(str(ws.cell(r, c).value or "").strip() for c in range(1, 4)):
                total_row = r
                break
        if not total_row:
            total_row = ws.max_row + 1

    no_col = find_no_column_smart(ws, excel_headers, header_row, total_row)
    if not no_col:
        no_col = 1 # Fallback cột A

    source_cols = list(table.get("columns", []))
    rows = _v229_filter_data_rows(table.get("rows", []))
    if not rows:
        raise ValueError("Không có dòng dữ liệu để nhập.")

    raw_mapping = None
    try:
        raw_mapping = self.mapping_editor.get_mapping()
    except Exception:
        raw_mapping = None
    if not raw_mapping or all(x is None for x in raw_mapping):
        raw_mapping = auto_mapping_to_excel_columns(source_cols, excel_headers)
        try:
            auto_idx = auto_map_columns(source_cols, excel_headers)
            auto_idx = ensure_no_column_in_mapping(source_cols, auto_idx, excel_headers)
            self.mapping_editor.set_mapping(source_cols, excel_headers, auto_idx)
        except Exception:
            pass

    mapping = _v229_normalize_mapping_to_excel_columns(source_cols, raw_mapping, excel_headers)
    if not mapping or all(x is None for x in mapping):
        raise ValueError("Chưa có mapping cột hoặc mapping đang bỏ qua toàn bộ cột.")

    first_data_row_old, style_row, last_no = _v229_find_stt_before_total(ws, no_col, header_row, total_row)
    strict_first_row = _v231_detect_first_data_row_strict(ws, header_row, total_row, no_col)
    first_data_row = min(first_data_row_old or strict_first_row, strict_first_row)

    # Quan trọng: lấy mẫu công thức TỔNG trước khi insert để giữ đúng dòng bắt đầu SUM.
    sum_col_first_rows = _v231_capture_sum_columns_and_starts(ws, total_row, first_data_row, no_col)

    insert_at = total_row
    row_count = len(rows)
    _v229_merged_ranges_shift_for_insert(ws, insert_at, row_count)
    total_row_after = total_row + row_count
    target_rows = list(range(insert_at, insert_at + row_count))

    for i, data_row in enumerate(rows):
        dst_row = target_rows[i]
        _v229_safe_copy_row(ws, style_row, dst_row)
        self._safe_set_cell_value(ws, dst_row, no_col, last_no + i + 1)

        for src_idx, excel_col in enumerate(mapping):
            if excel_col is None or excel_col == no_col:
                continue
            try:
                if src_idx < len(source_cols) and is_no_header(source_cols[src_idx]):
                    continue
            except Exception:
                pass
            try:
                # Ô có công thức mẫu thì giữ công thức, không ghi đè OCR.
                if _v229_is_formula(ws.cell(dst_row, excel_col).value):
                    continue
            except Exception:
                pass
            src_name = source_cols[src_idx] if src_idx < len(source_cols) else ""
            val = data_row[src_idx] if src_idx < len(data_row) else ""
            self._safe_set_cell_value(ws, dst_row, excel_col, _v23_convert_by_source(src_name, val))

    # CHỐT: SUM hết tới dòng ngay trước dòng TỔNG mới.
    sum_last_row = total_row_after - 1
    sum_columns = sorted(sum_col_first_rows.keys())
    for c in sum_columns:
        if c == no_col:
            continue
        try:
            letter = get_column_letter(c)
            start_row = sum_col_first_rows.get(c) or first_data_row
            ws.cell(total_row_after, c).value = f"=SUM({letter}{start_row}:{letter}{sum_last_row})"
        except Exception:
            pass

    force_workbook_recalculate(wb)

    validation = _v23_validate_written_cells(ws, target_rows, rows, source_cols, mapping, no_col)
    out = app_dir() / "last_run_v12"
    out.mkdir(exist_ok=True)
    logic = {
        "rule": "V23_1_SUM_HET_TRUOC_DONG_TONG",
        "sheet": ws.title,
        "header_row": header_row,
        "total_row_before": total_row,
        "insert_at": insert_at,
        "rows_added": row_count,
        "data_rows": target_rows,
        "total_row_after": total_row_after,
        "last_stt_before": last_no,
        "new_stt_start": last_no + 1,
        "new_stt_end": last_no + row_count,
        "first_data_row_detected_old": first_data_row_old,
        "first_data_row_strict": strict_first_row,
        "sum_last_row_before_total": sum_last_row,
        "sum_columns": [
            {"col": c, "letter": get_column_letter(c), "start_row": sum_col_first_rows.get(c), "formula": ws.cell(total_row_after, c).value}
            for c in sum_columns
        ],
        "mapping": [
            {
                "source": source_cols[i] if i < len(source_cols) else "",
                "excel_col": c,
                "excel_letter": get_column_letter(c) if c else None,
            }
            for i, c in enumerate(mapping)
        ],
        "written_cell_validation": validation,
    }
    (out / "v23_1_sum_het_truoc_dong_tong_logic.json").write_text(
        json.dumps(logic, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    return {
        "sheet": ws.title,
        "header_row": header_row,
        "last_stt_before": last_no,
        "start_fill_row": insert_at,
        "next_stt_start": last_no + 1,
        "rows_added": row_count,
        "total_row_after": total_row_after,
        "sum_first_row": first_data_row,
        "sum_last_row": sum_last_row,
        "sum_columns_count": len(sum_columns),
        "logic_file": str(out / "v23_1_sum_het_truoc_dong_tong_logic.json"),
    }

# Override cuối cùng: ép tool dùng logic SUM hết trước dòng TỔNG.
App._apply_rows_to_workbook = _v231_apply_rows_to_workbook

def main():
    try:
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("GiaKhanh.App")
        except Exception:
            pass
        root = tk.Tk()
        App(root)
        root.mainloop()
    except Exception as exc:
        try:
            import traceback
            (app_dir() / "gk_pilepro_error.log").write_text(
                traceback.format_exc(),
                encoding="utf-8",
            )
        except Exception:
            pass
        try:
            messagebox.showerror("GK PilePro", f"Không mở được ứng dụng:\n{exc}")
        except Exception:
            pass

if __name__ == "__main__":
    main()
