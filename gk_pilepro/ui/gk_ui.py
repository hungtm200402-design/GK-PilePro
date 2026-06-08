# -*- coding: utf-8 -*-

import math
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk


UI_BG = "#f5f8f7"

UI_SURFACE = "#ffffff"

UI_SURFACE_2 = "#f7fbf9"

UI_BORDER = "#e3eae7"

UI_TEXT = "#0f2433"

UI_MUTED = "#5c6f78"

UI_PRIMARY = "#007a45"

UI_PRIMARY_ACTIVE = "#006b3f"

UI_SUCCESS = "#007a45"

UI_SUCCESS_ACTIVE = "#006b3f"

UI_WARN = "#d88914"

UI_ERROR = "#dc2626"

UI_SCALE = 1.0
UI_FONT_BONUS = 0
UI_FONT_MIN_SIZE = 12
UI_READABILITY_MODE = False
UI_FONT_FAMILY = "Segoe UI"
UI_FONT_FAMILY_BOLD = "Segoe UI Semibold"


def configure_ui_metrics(scale=None, font_bonus=None, font_min_size=None, readability_mode=None):
    global UI_SCALE, UI_FONT_BONUS, UI_FONT_MIN_SIZE, UI_READABILITY_MODE
    if scale is not None:
        UI_SCALE = scale
    if font_bonus is not None:
        UI_FONT_BONUS = font_bonus
    if font_min_size is not None:
        UI_FONT_MIN_SIZE = font_min_size
    if readability_mode is not None:
        UI_READABILITY_MODE = readability_mode




def scale_px(value, minimum=1):

    try:

        return max(minimum, int(math.ceil(float(value) * float(UI_SCALE))))

    except Exception:

        return max(minimum, int(value))


def ui_font(size=None, bold=False):

    try:

        requested = size if size is not None else 12
        if requested <= 11:
            requested += UI_FONT_BONUS
        base = max(UI_FONT_MIN_SIZE, scale_px(requested))

    except Exception:

        base = max(UI_FONT_MIN_SIZE, 12 if size is None else int(size))

    family = UI_FONT_FAMILY_BOLD if bold else UI_FONT_FAMILY

    return (family, base, "normal")



class RoundedButton(tk.Canvas):

    def __init__(self, parent, text, command, width=15, variant="default"):

        self.command = command

        self.variant = variant

        self.text = text
        self.icon_image = None

        self.colors = {

            "default": (UI_SURFACE, UI_TEXT, "#f4faf7", UI_BORDER),

            "primary": (UI_PRIMARY, "#ffffff", UI_PRIMARY_ACTIVE, UI_PRIMARY),

            "success": (UI_SUCCESS, "#ffffff", UI_SUCCESS_ACTIVE, UI_SUCCESS),

            "soft": ("#ffffff", UI_PRIMARY, "#eefaf4", "#d8e7e1"),

            "warn": ("#fff8e8", "#c66b00", "#ffedc4", "#f0c15f"),

        }

        self.bg_color, self.fg_color, self.hover_color, self.border_color = self.colors.get(variant, self.colors["default"])
        self.press_color = self._press_color(self.hover_color)
        self._hovered = False
        self._pressed = False

        if width <= 0:

            self.pixel_width = max(scale_px(76), min(scale_px(124), int(len(str(text)) * 6.8 * UI_SCALE) + scale_px(30)))

        else:

            self.pixel_width = max(scale_px(74), scale_px(width * 9.0))

        self.pixel_height = scale_px(40)

        super().__init__(

            parent,

            width=self.pixel_width,

            height=self.pixel_height,

            bg=parent.cget("bg") if hasattr(parent, "cget") else UI_SURFACE,

            bd=0,

            highlightthickness=0,

            cursor="hand2",

        )

        self.bind("<ButtonPress-1>", self._press)

        self.bind("<ButtonRelease-1>", self._release)

        self.bind("<Enter>", self._enter)

        self.bind("<Leave>", self._leave)

        self._draw(self.bg_color)


    def _press_color(self, color):
        try:
            r, g, b = self.winfo_rgb(color)
            return "#{:02x}{:02x}{:02x}".format(
                max(0, int(r / 256 * 0.82)),
                max(0, int(g / 256 * 0.82)),
                max(0, int(b / 256 * 0.82)),
            )
        except Exception:
            return color

    def config(self, cnf=None, **kwargs):
        if cnf:
            kwargs.update(cnf)
        if "text" in kwargs:
            self.text = kwargs.pop("text")
            self._draw(self.hover_color if self._hovered else self.bg_color)
        if "image" in kwargs:
            self.icon_image = kwargs.pop("image")
            kwargs.pop("compound", None)
            self._draw(self.hover_color if self._hovered else self.bg_color)
        if kwargs:
            return super().config(**kwargs)
        return None

    configure = config


    def _draw(self, fill):

        self.delete("all")

        try:

            self.create_round_rect(1, 1, self.pixel_width - 1, self.pixel_height - 1, radius=max(11, scale_px(12)), fill=fill, outline=self.border_color)

        except Exception:

            self.create_rectangle(1, 1, self.pixel_width - 1, self.pixel_height - 1, fill=fill, outline=self.border_color)

        font_size = 9 if self.pixel_width <= scale_px(92) else 10
        font = ui_font(font_size, bold=self.variant in {"primary", "success"})

        if self.icon_image is not None:
            try:
                text_width = max(0, int(tkfont.Font(font=font).measure(str(self.text))))
            except Exception:
                text_width = int(len(str(self.text)) * 7 * UI_SCALE)
            gap = scale_px(8)
            icon_w = scale_px(16)
            total_width = icon_w + gap + text_width
            icon_x = max(scale_px(13), (self.pixel_width - total_width) // 2 + icon_w // 2)
            self.create_image(icon_x, self.pixel_height // 2, image=self.icon_image)
            self.create_text(icon_x + icon_w // 2 + gap, self.pixel_height // 2, text=self.text, fill=self.fg_color, font=font, anchor="w")
        else:
            self.create_text(self.pixel_width // 2, self.pixel_height // 2, text=self.text, fill=self.fg_color, font=font)



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

    def _enter(self, _event=None):
        self._hovered = True
        self._draw(self.press_color if self._pressed else self.hover_color)

    def _leave(self, _event=None):
        self._hovered = False
        self._pressed = False
        self._draw(self.bg_color)

    def _press(self, _event=None):
        self._pressed = True
        self._draw(self.press_color)

    def _release(self, event=None):
        was_pressed = self._pressed
        self._pressed = False
        inside = True
        try:
            inside = 0 <= event.x <= self.pixel_width and 0 <= event.y <= self.pixel_height
        except Exception:
            pass
        self._draw(self.hover_color if inside else self.bg_color)
        if was_pressed and inside:
            self._click(event)



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

        self.width_px = scale_px(width)
        self.height_px = scale_px(height)

        super().__init__(

            parent,

            width=self.width_px,

            height=self.height_px,

            bg=parent.cget("bg") if hasattr(parent, "cget") else UI_SURFACE,

            bd=0,

            highlightthickness=0,

        )

        self.text = text

        self.bg_color = bg_color

        self.border_color = border_color

        self.text_color = text_color

        self.radius = radius

        self.bind("<Configure>", lambda _e: self._draw())

        self._draw()

    def _fit_text(self, text, max_width):
        text = str(text or "")
        if max_width <= 20:
            return ""
        font = ui_font(10)
        item = self.create_text(0, -100, text=text, font=font, anchor="w")
        try:
            bbox = self.bbox(item)
            if bbox and bbox[2] <= max_width:
                self.delete(item)
                return text
        except Exception:
            pass
        self.delete(item)
        suffix = "..."
        out = text
        while out:
            trial = out[:-1].rstrip() + suffix
            item = self.create_text(0, -100, text=trial, font=font, anchor="w")
            try:
                bbox = self.bbox(item)
                if bbox and bbox[2] <= max_width:
                    self.delete(item)
                    return trial
            except Exception:
                pass
            self.delete(item)
            out = out[:-1]
        return suffix



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

        font_size = 9 if w < 110 else 10
        self._text_font = ui_font(font_size)

        self.create_text(

            9,

            self.height_px // 2,

            text=self._fit_text(self.text, max(10, w - 18)),

            fill=self.text_color,

            font=self._text_font,

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

        self.width_px = scale_px(width)
        self.height_px = scale_px(height)

        super().__init__(

            parent,

            width=self.width_px,

            height=self.height_px,

            bg=parent.cget("bg") if hasattr(parent, "cget") else UI_SURFACE,

            bd=0,

            highlightthickness=0,

            cursor="hand2",

        )

        self.values = list(values or [])

        self.variable = variable or tk.StringVar()

        self.bg_color = bg_color

        self.border_color = border_color

        self.radius = radius

        self._combo_callbacks = []
        self.popup = None
        self.popup_canvas = None
        self.popup_scrollbar = None
        self.popup_inner = None
        self.popup_content = None
        self._popup_global_click_binding = None

        self.menu = tk.Menu(self, tearoff=0)
        try:
            self.menu.configure(
                bg="#ffffff",
                fg=UI_TEXT,
                activebackground="#e8f1ff",
                activeforeground=UI_PRIMARY,
                bd=1,
                relief="solid",
                font=ui_font(10),
                cursor="hand2",
                activeborderwidth=0,
                borderwidth=1,
            )
        except Exception:
            pass

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
        if getattr(self, "popup", None) is not None and self.popup.winfo_exists():
            self._build_popup()



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
        self._close_popup()

        for callback in list(self._combo_callbacks):

            try:

                callback(None)

            except Exception:

                pass



    def _open_menu(self, event=None):
        self._close_popup()
        if not self.values:
            return
        self.popup = tk.Toplevel(self)
        self.popup.overrideredirect(True)
        try:
            self.popup.attributes("-topmost", True)
        except Exception:
            pass
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + scale_px(2)
        width = max(self.width_px, scale_px(150))
        row_h = scale_px(28)
        max_visible_rows = 8
        height = max(scale_px(24), min(row_h * len(self.values) + scale_px(6), row_h * max_visible_rows + scale_px(6)))
        self.popup.geometry(f"{width}x{height}+{x}+{y}")
        self.popup.configure(bg="#d8e5f4")
        self.popup.bind("<Escape>", lambda _e: self._close_popup())
        self._build_popup()
        try:
            self._popup_global_click_binding = self.winfo_toplevel().bind("<Button-1>", self._close_popup_if_clicked_outside, add="+")
        except Exception:
            self._popup_global_click_binding = None
        try:
            self.popup.focus_force()
        except Exception:
            pass

    def _widget_is_popup_child(self, widget):
        popup = getattr(self, "popup", None)
        while widget is not None:
            if widget is self or widget is popup:
                return True
            widget = getattr(widget, "master", None)
        return False

    def _close_popup_if_clicked_outside(self, event=None):
        try:
            widget = getattr(event, "widget", None)
            if self._widget_is_popup_child(widget):
                return
        except Exception:
            pass
        self._close_popup()

    def _close_popup(self):
        popup = getattr(self, "popup", None)
        self.popup = None
        self.popup_canvas = None
        self.popup_scrollbar = None
        self.popup_inner = None
        self.popup_content = None
        try:
            if self._popup_global_click_binding:
                self.winfo_toplevel().unbind("<Button-1>", self._popup_global_click_binding)
        except Exception:
            pass
        self._popup_global_click_binding = None
        if popup is not None:
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except Exception:
                pass

    def _on_popup_mousewheel(self, event):
        canvas = getattr(self, "popup_canvas", None)
        if canvas is None:
            return
        try:
            delta = getattr(event, "delta", 0) or 0
            if delta:
                canvas.yview_scroll(int(-1 * (delta / 120)), "units")
                return
            num = getattr(event, "num", None)
            if num == 4:
                canvas.yview_scroll(-1, "units")
            elif num == 5:
                canvas.yview_scroll(1, "units")
        except Exception:
            pass

    def _bind_popup_mousewheel_recursive(self, widget=None):
        widget = widget or getattr(self, "popup_content", None)
        if widget is None:
            return
        try:
            widget.bind("<MouseWheel>", self._on_popup_mousewheel, add="+")
            widget.bind("<Button-4>", self._on_popup_mousewheel, add="+")
            widget.bind("<Button-5>", self._on_popup_mousewheel, add="+")
        except Exception:
            pass
        try:
            for child in widget.winfo_children():
                self._bind_popup_mousewheel_recursive(child)
        except Exception:
            pass

    def _build_popup(self):
        popup = getattr(self, "popup", None)
        if popup is None or not popup.winfo_exists():
            return
        for child in popup.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass
        outer = tk.Frame(popup, bg="#d8e5f4", bd=0, highlightthickness=0)
        outer.pack(fill="both", expand=True)
        inner = tk.Frame(outer, bg="#ffffff", bd=1, relief="solid", highlightthickness=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        self.popup_inner = inner
        self.popup_canvas = tk.Canvas(inner, bg="#ffffff", bd=0, highlightthickness=0)
        self.popup_scrollbar = ttk.Scrollbar(inner, orient="vertical", command=self.popup_canvas.yview)
        self.popup_canvas.configure(yscrollcommand=self.popup_scrollbar.set)
        self.popup_canvas.pack(side="left", fill="both", expand=True)
        self.popup_scrollbar.pack(side="right", fill="y")
        self.popup_content = tk.Frame(self.popup_canvas, bg="#ffffff")
        content_window = self.popup_canvas.create_window((0, 0), window=self.popup_content, anchor="nw")

        def _sync_scrollregion(_event=None):
            try:
                self.popup_canvas.configure(scrollregion=self.popup_canvas.bbox("all"))
            except Exception:
                pass

        def _sync_width(event):
            try:
                self.popup_canvas.itemconfigure(content_window, width=event.width)
            except Exception:
                pass

        self.popup_content.bind("<Configure>", _sync_scrollregion)
        self.popup_canvas.bind("<Configure>", _sync_width)
        self.popup_canvas.bind("<Enter>", lambda _e: self.popup_canvas.focus_set())
        for value in self.values:
            item = tk.Label(
                self.popup_content,
                text=value,
                bg="#ffffff",
                fg=UI_TEXT,
                font=ui_font(10),
                anchor="w",
                padx=12,
                pady=4,
            )
            item.pack(fill="x")
            item.bind("<Enter>", lambda _e, w=item: w.configure(bg="#e8f1ff", fg=UI_PRIMARY))
            item.bind("<Leave>", lambda _e, w=item: w.configure(bg="#ffffff", fg=UI_TEXT))
            item.bind("<Button-1>", lambda _e, v=value: self._select(v))
        self._bind_popup_mousewheel_recursive(self.popup_content)
        try:
            self.popup_canvas.yview_moveto(0)
        except Exception:
            pass



    def _fit_text(self, text, max_width):

        text = str(text or "")

        if max_width <= 20:

            return ""

        font = ui_font(10)

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

        self.create_line(arrow_x - 8, 6, arrow_x - 8, self.height_px - 6, fill="#d5e3f5")

        self.create_polygon(

            arrow_x - 4,

            self.height_px // 2 - 2,

            arrow_x + 4,

            self.height_px // 2 - 2,

            arrow_x,

            self.height_px // 2 + 3,

            fill=UI_PRIMARY,

            outline=UI_PRIMARY,

        )

        text = self._fit_text(self.variable.get(), max(10, w - 38))

        self.create_text(12, self.height_px // 2, text=text, fill=UI_TEXT, font=ui_font(10), anchor="w")



class RoundedMappingEntry(tk.Canvas):

    def __init__(

        self,

        parent,

        textvariable,

        bg_color,

        border_color,

        width=220,

        height=30,

        radius=7,

        font=None,

    ):

        self.width_px = scale_px(width)
        self.height_px = scale_px(height)

        super().__init__(

            parent,

            width=self.width_px,

            height=self.height_px,

            bg=parent.cget("bg") if hasattr(parent, "cget") else UI_SURFACE,

            bd=0,

            highlightthickness=0,

        )

        self.variable = textvariable or tk.StringVar()
        self.bg_color = bg_color
        self.border_color = border_color
        self.radius = radius
        self.border_color_current = border_color
        self._entry_font = font or ui_font(11)

        self.inner = tk.Frame(self, bg=self.bg_color, bd=0, highlightthickness=0)
        self.entry = tk.Entry(
            self.inner,
            textvariable=self.variable,
            relief="flat",
            bd=0,
            bg=self.bg_color,
            fg=UI_TEXT,
            insertbackground=UI_TEXT,
            highlightthickness=0,
            font=self._entry_font,
        )
        self._window_id = self.create_window((0, 0), window=self.inner, anchor="nw")

        self.bind("<Configure>", self._draw)
        self.entry.bind("<FocusIn>", self._focus)
        self.entry.bind("<FocusOut>", self._blur)
        self.inner.bind("<Button-1>", lambda _e: self.entry.focus_set())
        self.bind("<Button-1>", lambda _e: self.entry.focus_set())
        self.entry.bind("<Button-1>", lambda _e: self.entry.focus_set())
        self.variable.trace_add("write", lambda *_: self._draw())
        self._draw()

    def _focus(self, _event=None):
        self.border_color_current = UI_PRIMARY
        self._draw()

    def _blur(self, _event=None):
        self.border_color_current = self.border_color
        self._draw()

    def _draw(self, _event=None):
        self.delete("bg")
        w = max(40, self.winfo_width())
        h = max(20, self.winfo_height())
        border = getattr(self, "border_color_current", self.border_color)
        self._round_rect(1, 1, w - 1, h - 1, self.radius, fill=self.bg_color, outline=border, tags=("bg",))
        pad_x = max(10, scale_px(10))
        pad_y = max(5, scale_px(5))
        inner_w = max(1, w - pad_x * 2)
        inner_h = max(1, h - pad_y * 2)
        self.coords(self._window_id, pad_x, pad_y)
        self.itemconfigure(self._window_id, width=inner_w, height=inner_h)
        self.tag_raise(self._window_id)
        self.inner.configure(bg=self.bg_color, width=inner_w, height=inner_h)
        self.entry.configure(width=max(1, int(inner_w / 8)))
        try:
            self.entry.pack_forget()
        except Exception:
            pass
        self.entry.pack(fill="both", expand=True)

    def _round_rect(self, x1, y1, x2, y2, radius=7, **kwargs):
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        return self.create_polygon(points, smooth=True, splinesteps=18, **kwargs)



# Các marker dùng để nhận diện dòng TỔNG/TOTAL.
