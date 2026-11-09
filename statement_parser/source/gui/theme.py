import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk
from typing import Any, Callable


BACKGROUND = "#111820"
SURFACE = "#1a232c"
SURFACE_RAISED = "#202c36"
SURFACE_HOVER = "#293841"
BORDER = "#33414b"
DROP_SURFACE = "#2b343b"
DROP_BORDER = "#66747d"
TEXT_PRIMARY = "#edf2f2"
TEXT_SECONDARY = "#a8b5bc"
ACCENT = "#65c6aa"
ACCENT_HOVER = "#7bd6bb"
ACCENT_PRESSED = "#4eac92"
STATUS_ERROR = "#e4777f"
LOG_BACKGROUND = "#141b22"
STATUS_OVER = "#e4777f"
STATUS_UNDER = "#7bd6bb"
METRIC_SPENT = "#e5a94d"
METRIC_BUDGET = "#78a9e3"
SELECTION_BACKGROUND = "#244e43"


def configure_styles(style: ttk.Style) -> None:
    if "clam" in style.theme_names():
        style.theme_use("clam")

    style.configure(
        ".",
        background=BACKGROUND,
        foreground=TEXT_PRIMARY,
        font=("Arial", 15),
    )
    style.configure("App.TFrame", background=BACKGROUND)
    style.configure(
        "Card.TFrame",
        background=SURFACE,
        borderwidth=0,
    )
    style.configure(
        "Title.TLabel",
        background=BACKGROUND,
        foreground=TEXT_PRIMARY,
        font=("Arial", 34, "bold"),
    )
    style.configure(
        "Subtitle.TLabel",
        background=BACKGROUND,
        foreground=TEXT_SECONDARY,
        font=("Arial", 14),
    )
    style.configure(
        "CardTitle.TLabel",
        background=SURFACE,
        foreground=TEXT_PRIMARY,
        font=("Arial", 18, "bold"),
    )
    style.configure(
        "App.CardTitle.TLabel",
        background=BACKGROUND,
        foreground=TEXT_PRIMARY,
        font=("Arial", 18, "bold"),
    )
    style.configure(
        "Muted.TLabel",
        background=SURFACE,
        foreground=TEXT_SECONDARY,
        font=("Arial", 14),
    )
    style.configure(
        "Primary.TButton",
        background=ACCENT,
        foreground=BACKGROUND,
        font=("Arial", 14, "bold"),
        padding=(16, 11),
        borderwidth=0,
    )
    style.map(
        "Primary.TButton",
        background=[
            ("pressed", ACCENT_PRESSED),
            ("active", ACCENT_HOVER),
            ("disabled", SURFACE_RAISED),
        ],
        foreground=[("disabled", TEXT_SECONDARY)],
    )
    style.configure(
        "TButton",
        background=SURFACE_RAISED,
        foreground=TEXT_PRIMARY,
        font=("Arial", 13, "bold"),
        padding=(15, 11),
        bordercolor=BORDER,
        focuscolor=ACCENT,
    )
    style.map(
        "TButton",
        background=[
            ("pressed", SURFACE_HOVER),
            ("active", SURFACE_HOVER),
            ("disabled", SURFACE_RAISED),
        ],
        foreground=[("disabled", TEXT_SECONDARY)],
    )
    style.configure(
        "TEntry",
        fieldbackground=LOG_BACKGROUND,
        foreground=TEXT_PRIMARY,
        insertcolor=ACCENT,
        font=("Arial", 15),
        padding=10,
        borderwidth=1,
        relief="flat",
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        focuscolor=ACCENT,
    )
    style.map(
        "TEntry",
        fieldbackground=[("disabled", SURFACE_RAISED)],
        foreground=[("disabled", TEXT_SECONDARY)],
    )
    style.configure(
        "TCombobox",
        fieldbackground=LOG_BACKGROUND,
        background=SURFACE_RAISED,
        foreground=TEXT_PRIMARY,
        arrowcolor=TEXT_SECONDARY,
        font=("Arial", 15),
        padding=9,
        borderwidth=1,
        relief="flat",
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        focuscolor=ACCENT,
    )
    style.map(
        "TCombobox",
        fieldbackground=[
            ("readonly", LOG_BACKGROUND),
            ("disabled", SURFACE_RAISED),
        ],
        foreground=[("readonly", TEXT_PRIMARY), ("disabled", TEXT_SECONDARY)],
        selectbackground=[("readonly", SURFACE_HOVER)],
        selectforeground=[("readonly", TEXT_PRIMARY)],
    )
    style.configure(
        "Treeview",
        rowheight=44,
        background=SURFACE,
        fieldbackground=SURFACE,
        foreground=TEXT_PRIMARY,
        font=("Arial", 14),
        borderwidth=0,
        relief="flat",
        bordercolor=SURFACE,
        lightcolor=SURFACE,
        darkcolor=SURFACE,
    )
    style.configure(
        "Treeview.Heading",
        background=SURFACE_RAISED,
        foreground=TEXT_SECONDARY,
        font=("Arial", 13, "bold"),
        padding=(10, 13),
        borderwidth=0,
        relief="flat",
        bordercolor=SURFACE_RAISED,
        lightcolor=SURFACE_RAISED,
        darkcolor=SURFACE_RAISED,
    )
    style.configure(
        "CategoryTrend.Treeview",
        rowheight=48,
        background=SURFACE,
        fieldbackground=SURFACE,
        foreground=TEXT_PRIMARY,
        font=("Arial", 16),
        borderwidth=0,
        relief="flat",
        bordercolor=SURFACE,
        lightcolor=SURFACE,
        darkcolor=SURFACE,
    )
    style.configure(
        "CategoryTrend.Treeview.Heading",
        background=SURFACE_RAISED,
        foreground=TEXT_SECONDARY,
        font=("Arial", 14, "bold"),
        padding=(10, 13),
        borderwidth=0,
        relief="flat",
        bordercolor=SURFACE_RAISED,
        lightcolor=SURFACE_RAISED,
        darkcolor=SURFACE_RAISED,
    )
    style.map(
        "CategoryTrend.Treeview",
        background=[("selected", SELECTION_BACKGROUND)],
        foreground=[("selected", TEXT_PRIMARY)],
    )
    style.map(
        "CategoryTrend.Treeview.Heading",
        background=[("active", SURFACE_HOVER)],
        foreground=[("active", TEXT_PRIMARY)],
    )
    style.configure(
        "DataTable.Heading.TLabel",
        background=SURFACE,
        foreground=TEXT_SECONDARY,
        font=("Arial", 15, "bold"),
        padding=(12, 14),
    )
    style.configure(
        "DataTable.Horizontal.TSeparator",
        background=BORDER,
    )
    for stripe, background in (("Even", SURFACE), ("Odd", SURFACE)):
        style.configure(
            f"DataTable.{stripe}.TFrame",
            background=background,
        )
        style.configure(
            f"DataTable.{stripe}.TLabel",
            background=background,
            foreground=TEXT_PRIMARY,
            font=("Arial", 16),
            padding=(12, 14),
        )
        for status, foreground in (
            ("Under", STATUS_UNDER),
            ("Over", STATUS_OVER),
            ("NoTarget", TEXT_SECONDARY),
        ):
            style.configure(
                f"DataTable.{stripe}.{status}.TLabel",
                background=background,
                foreground=foreground,
                font=("Arial", 16, "bold"),
                padding=(12, 14),
            )
    style.map(
        "Treeview",
        background=[("selected", SELECTION_BACKGROUND)],
        foreground=[("selected", TEXT_PRIMARY)],
    )
    style.map(
        "Treeview.Heading",
        background=[("active", SURFACE_HOVER)],
        foreground=[("active", TEXT_PRIMARY)],
    )
    style.configure(
        "Vertical.TScrollbar",
        background=SURFACE_RAISED,
        troughcolor=LOG_BACKGROUND,
        arrowcolor=TEXT_SECONDARY,
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        arrowsize=14,
        width=12,
        borderwidth=1,
        relief="flat",
    )
    style.map(
        "Vertical.TScrollbar",
        background=[("pressed", ACCENT_PRESSED), ("active", ACCENT)],
    )
    style.configure(
        "TPanedwindow",
        background=BACKGROUND,
        borderwidth=0,
        relief="flat",
        sashwidth=8,
        sashpad=0,
        sashrelief="flat",
    )
    style.configure(
        "TNotebook",
        background=BACKGROUND,
        borderwidth=0,
        relief="flat",
        tabmargins=(0, 0, 0, 0),
        bordercolor=BACKGROUND,
        lightcolor=BACKGROUND,
        darkcolor=BACKGROUND,
        focuscolor=BACKGROUND,
    )
    style.configure(
        "TNotebook.Tab",
        background=SURFACE_RAISED,
        foreground=TEXT_SECONDARY,
        font=("Arial", 14, "bold"),
        padding=(16, 9),
        width=24,
        anchor="center",
        borderwidth=0,
        relief="flat",
        bordercolor=BACKGROUND,
        lightcolor=BACKGROUND,
        darkcolor=BACKGROUND,
    )
    style.map(
        "TNotebook.Tab",
        background=[
            ("selected", SURFACE_HOVER),
            ("active", SURFACE_HOVER),
        ],
        foreground=[("selected", ACCENT), ("active", TEXT_PRIMARY)],
        font=[("selected", ("Arial", 14, "bold"))],
        padding=[("selected", (16, 9)), ("active", (16, 9))],
    )


def _rounded_fill(
    canvas: tk.Canvas,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    radius: float,
    color: str,
    tag: str,
) -> None:
    if x2 <= x1 or y2 <= y1:
        return
    radius = min(radius, (x2 - x1) / 2, (y2 - y1) / 2)
    canvas.create_rectangle(
        x1 + radius,
        y1,
        x2 - radius,
        y2,
        fill=color,
        outline=color,
        tags=(tag,),
    )
    canvas.create_rectangle(
        x1,
        y1 + radius,
        x2,
        y2 - radius,
        fill=color,
        outline=color,
        tags=(tag,),
    )
    for left, top, right, bottom, start in (
        (x1, y1, x1 + 2 * radius, y1 + 2 * radius, 90),
        (x2 - 2 * radius, y1, x2, y1 + 2 * radius, 0),
        (x2 - 2 * radius, y2 - 2 * radius, x2, y2, 270),
        (x1, y2 - 2 * radius, x1 + 2 * radius, y2, 180),
    ):
        canvas.create_arc(
            left,
            top,
            right,
            bottom,
            start=start,
            extent=90,
            style=tk.PIESLICE,
            fill=color,
            outline=color,
            tags=(tag,),
        )


def _canvas_background(master: Any) -> str:
    current = master
    while current is not None:
        try:
            style_name = current.cget("style")
        except (AttributeError, tk.TclError):
            style_name = ""
        if style_name.startswith("Card"):
            return SURFACE
        if style_name.startswith("DataTable"):
            return LOG_BACKGROUND
        current = getattr(current, "master", None)
    return BACKGROUND


class RoundedButton(tk.Canvas):
    def __init__(
        self,
        master: Any,
        text: str = "",
        command: Callable[[], Any] | None = None,
        *,
        style: str = "",
        textvariable: tk.StringVar | None = None,
        state: str = tk.NORMAL,
        width: int | None = None,
        height: int = 44,
    ) -> None:
        self.__primary = style == "Primary.TButton"
        self.__tab = style == "Tab.TButton"
        self.__selected = False
        self.__text = text
        self.__textvariable = textvariable
        self.__command = command
        self.__state = state
        self.__radius = 16 if self.__tab else 12
        self.__font_spec = (
            ("Arial", 14, "bold")
            if self.__primary or self.__tab
            else ("Arial", 12, "bold")
        )
        self.__font = None
        self.__fixed_width = width
        self.__canvas_background = _canvas_background(master)
        super().__init__(
            master,
            width=width or 150,
            height=height,
            bg=self.__canvas_background,
            borderwidth=0,
            highlightthickness=0,
            takefocus=True,
            cursor="hand2",
        )
        self.__font = tkfont.Font(self, font=self.__font_spec)
        self.__hovered = False
        self.__pressed = False
        self.bind("<Configure>", self.__draw)
        self.bind("<Enter>", self.__on_enter)
        self.bind("<Leave>", self.__on_leave)
        self.bind("<ButtonPress-1>", self.__on_press)
        self.bind("<ButtonRelease-1>", self.__on_release)
        self.bind("<Return>", self.__on_key)
        self.bind("<space>", self.__on_key)
        self.bind("<FocusIn>", lambda _event: self.__draw())
        self.bind("<FocusOut>", lambda _event: self.__draw())
        self.__trace_id = None
        if textvariable is not None:
            self.__trace_id = textvariable.trace_add("write", self.__on_text_change)
            self.__text = textvariable.get()
        self.__update_requested_width()

    def configure(self, cnf=None, **kwargs):
        if isinstance(cnf, dict):
            kwargs = {**cnf, **kwargs}
            cnf = None
        elif cnf is not None:
            return tk.Canvas.configure(self, cnf, **kwargs)

        text = kwargs.pop("text", None)
        if text is not None:
            self.__text = text
            if self.__textvariable is not None:
                self.__textvariable.trace_remove("write", self.__trace_id)
                self.__textvariable = None
                self.__trace_id = None
        textvariable = kwargs.pop("textvariable", None)
        if textvariable is not None:
            if self.__textvariable is not None and self.__trace_id is not None:
                self.__textvariable.trace_remove("write", self.__trace_id)
            self.__textvariable = textvariable
            self.__trace_id = textvariable.trace_add("write", self.__on_text_change)
            self.__text = textvariable.get()
        if "command" in kwargs:
            self.__command = kwargs.pop("command")
        if "state" in kwargs:
            self.__state = kwargs.pop("state")
            tk.Canvas.configure(
                self,
                cursor="arrow" if self.__state == tk.DISABLED else "hand2",
            )
        if "style" in kwargs:
            style = kwargs.pop("style")
            self.__primary = style == "Primary.TButton"
            self.__tab = style == "Tab.TButton"
            self.__radius = 16 if self.__tab else 12
            self.__font_spec = (
                ("Arial", 14, "bold")
                if self.__primary or self.__tab
                else ("Arial", 12, "bold")
            )
            self.__font = tkfont.Font(self, font=self.__font_spec)
            self.__update_requested_width()
        if "width" in kwargs:
            self.__fixed_width = kwargs["width"]
        if text is not None or textvariable is not None or "width" in kwargs:
            self.__update_requested_width()
        if kwargs:
            tk.Canvas.configure(self, **kwargs)
        self.__draw()

    config = configure

    def cget(self, key: str):
        if key == "text":
            return self.__text
        if key == "textvariable":
            return self.__textvariable
        if key == "command":
            return self.__command
        if key == "state":
            return self.__state
        return tk.Canvas.cget(self, key)

    def invoke(self):
        if self.__state != tk.DISABLED and self.__command is not None:
            return self.__command()
        return None

    def set_text(self, text: str) -> None:
        self.configure(text=text)

    def set_selected(self, selected: bool) -> None:
        self.__selected = selected
        self.__draw()

    def __update_requested_width(self) -> None:
        if self.__fixed_width is not None:
            width = self.__fixed_width
        else:
            padding = 34 if self.__primary or self.__tab else 28
            width = max(92, self.__font.measure(self.__text) + padding)
        tk.Canvas.configure(self, width=width)

    def __on_text_change(self, *_args) -> None:
        if self.__textvariable is not None:
            self.__text = self.__textvariable.get()
        self.__update_requested_width()
        self.__draw()

    def __draw(self, _event=None) -> None:
        width = self.winfo_width()
        height = self.winfo_height()
        if width <= 1 or height <= 1:
            return
        self.delete("rounded-button")
        disabled = self.__state == tk.DISABLED
        if disabled:
            background, foreground = SURFACE_RAISED, TEXT_SECONDARY
        elif self.__tab:
            background = SURFACE_HOVER if self.__selected else SURFACE
            foreground = ACCENT if self.__selected else TEXT_SECONDARY
            if self.__pressed:
                background = SURFACE_HOVER
            elif self.__hovered and not self.__selected:
                background = SURFACE_RAISED
        elif self.__primary:
            background, foreground = ACCENT, BACKGROUND
        else:
            background, foreground = SURFACE_RAISED, TEXT_PRIMARY
        if not self.__tab:
            if self.__pressed and not disabled:
                background = ACCENT_PRESSED if self.__primary else BORDER
            elif self.__hovered and not disabled:
                background = ACCENT_HOVER if self.__primary else SURFACE_HOVER

        if self.focus_get() is self:
            _rounded_fill(
                self,
                0,
                0,
                width - 1,
                height - 1,
                self.__radius,
                ACCENT,
                "rounded-button",
            )
            _rounded_fill(
                self,
                1,
                1,
                width - 2,
                height - 2,
                self.__radius - 1,
                background,
                "rounded-button",
            )
        else:
            _rounded_fill(
                self,
                0,
                0,
                width - 1,
                height - 1,
                self.__radius,
                background,
                "rounded-button",
            )
        self.create_text(
            width / 2,
            height / 2,
            text=self.__text,
            fill=foreground,
            font=self.__font_spec,
            width=max(1, width - 24),
            tags=("rounded-button",),
        )

    def __on_enter(self, _event) -> None:
        self.__hovered = True
        self.__draw()

    def __on_leave(self, _event) -> None:
        self.__hovered = False
        self.__pressed = False
        self.__draw()

    def __on_press(self, _event) -> str:
        if self.__state == tk.DISABLED:
            return "break"
        self.__pressed = True
        self.focus_set()
        self.__draw()
        return "break"

    def __on_release(self, event) -> str:
        was_pressed = self.__pressed
        self.__pressed = False
        self.__draw()
        if (
            was_pressed
            and self.__state != tk.DISABLED
            and 0 <= event.x <= self.winfo_width()
            and 0 <= event.y <= self.winfo_height()
        ):
            self.invoke()
        return "break"

    def __on_key(self, _event) -> str:
        self.invoke()
        return "break"


class RoundedCard(tk.Canvas):
    def __init__(
        self,
        master: Any,
        *,
        background: str = SURFACE,
        canvas_background: str | None = None,
        outline: str = BORDER,
        width: int = 500,
        height: int = 78,
        radius: int = 16,
        inset: int = 4,
    ) -> None:
        if canvas_background is None:
            canvas_background = _canvas_background(master)
        super().__init__(
            master,
            width=width,
            height=height,
            bg=canvas_background,
            borderwidth=0,
            highlightthickness=0,
        )
        self.content = tk.Frame(
            self,
            bg=background,
            borderwidth=0,
            highlightthickness=0,
        )
        self.__background = background
        self.__outline = outline
        self.__radius = radius
        self.__inset = inset
        self.__content_window = self.create_window(
            inset,
            inset,
            window=self.content,
            anchor="nw",
        )
        self.bind("<Configure>", self.__resize)

    def __resize(self, event) -> None:
        self.coords(self.__content_window, self.__inset, self.__inset)
        self.itemconfigure(
            self.__content_window,
            width=max(1, event.width - 2 * self.__inset),
            height=max(1, event.height - 2 * self.__inset),
        )
        self.__draw(event.width, event.height)

    def __draw(self, width: int, height: int) -> None:
        self.delete("rounded-card")
        _rounded_fill(
            self,
            0,
            0,
            width - 1,
            height - 1,
            self.__radius,
            self.__outline,
            "rounded-card",
        )
        _rounded_fill(
            self,
            1,
            1,
            width - 2,
            height - 2,
            max(0, self.__radius - 1),
            self.__background,
            "rounded-card",
        )
        self.tag_lower("rounded-card")
