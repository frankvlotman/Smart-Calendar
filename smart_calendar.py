
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date, datetime, timedelta
import calendar
import os

# --- Blank icon support (uses Pillow if available) ---
ICON_PATH = r'C:\Users\Frank\Desktop\blank.ico'

def _create_blank_icon_if_needed(path: str):
    try:
        from PIL import Image  # Pillow
    except Exception:
        return  # Pillow not available; skip silently

    parent = os.path.dirname(path)
    if parent and not os.path.isdir(parent):
        try:
            os.makedirs(parent, exist_ok=True)
        except Exception:
            return

    if not os.path.isfile(path):
        try:
            size = (16, 16)
            image = Image.new("RGBA", size, (255, 255, 255, 0))  # fully transparent
            image.save(path, format="ICO")
        except Exception:
            pass

APP_TITLE = "Smart Calendar"
DATE_FMT = "%d/%m/%Y"  # UK format
WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

def add_working_days(start: date, days: int) -> date:
    step = 1 if days >= 0 else -1
    remaining = abs(days)
    current = start
    while remaining > 0:
        current += timedelta(days=step)
        if current.weekday() < 5:  # Mon-Fri
            remaining -= 1
    return current

def working_days_between(a: date, b: date) -> int:
    """Inclusive count of Mon-Fri between a and b."""
    if a > b:
        a, b = b, a
    days = (b - a).days + 1  # inclusive
    full_weeks, extra_days = divmod(days, 7)
    # Each full week contributes 5 working days
    workdays = full_weeks * 5
    # Handle the remainder days
    for i in range(extra_days):
        d = a + timedelta(days=full_weeks * 7 + i)
        if d.weekday() < 5:
            workdays += 1
    return workdays

class WorkingDaysDialog(tk.Toplevel):
    def __init__(self, master, on_result):
        super().__init__(master)
        self.title("Add Working Days")
        self.resizable(False, False)
        self.on_result = on_result
        self.configure(bg="#ffffff")

        self.transient(master)
        self.grab_set()

        frm = ttk.Frame(self, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frm, text="Start date (DD/MM/YYYY):").grid(row=0, column=0, sticky="w")
        self.ent_start = ttk.Entry(frm, width=18)
        self.ent_start.grid(row=0, column=1, sticky="w", padx=(8,0))
        self.ent_start.insert(0, date.today().strftime(DATE_FMT))

        ttk.Label(frm, text="Number of working days:").grid(row=1, column=0, sticky="w", pady=(8,0))
        self.ent_days = ttk.Entry(frm, width=18)
        self.ent_days.grid(row=1, column=1, sticky="w", padx=(8,0), pady=(8,0))
        self.ent_days.insert(0, "0")

        btns = ttk.Frame(frm)
        btns.grid(row=2, column=0, columnspan=2, sticky="e", pady=(12,0))
        ttk.Button(btns, text="Calculate", command=self._ok).grid(row=0, column=0, padx=(0,8))
        ttk.Button(btns, text="Cancel", command=self.destroy).grid(row=0, column=1)

        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self.ent_start.focus_set()

    def _ok(self):
        try:
            start_s = self.ent_start.get().strip()
            days_s = self.ent_days.get().strip()
            start = datetime.strptime(start_s, DATE_FMT).date()
            days = int(days_s)
        except Exception:
            messagebox.showerror("Invalid input", "Please provide a valid date (DD/MM/YYYY) and an integer for days.")
            return

        result = add_working_days(start, days)
        self.on_result(start, days, result)
        self.destroy()


class ThreeMonthWindow(tk.Toplevel):
    def __init__(self, master, style_helper, show_week_numbers=False):
        super().__init__(master)
        self.title("3-Month View")
        self.resizable(True, True)
        self.configure(bg="#ffffff")
        self.style_helper = style_helper
        self.base_year = date.today().year
        self.base_month = date.today().month
        self.show_week_numbers = show_week_numbers

        self._build_ui()
        self.render_all()

    def _build_ui(self):
        hdr = ttk.Frame(self, padding=(8,8,8,0))
        hdr.pack(fill="x")

        ttk.Button(hdr, text="◀", command=self.prev_three, width=3).pack(side="left")
        self.lbl_range = ttk.Label(hdr, text="", style="Header.TLabel", anchor="center")
        self.lbl_range.pack(side="left", expand=True)
        ttk.Button(hdr, text="▶", command=self.next_three, width=3).pack(side="right")

        self.body = ttk.Frame(self, padding=8)
        self.body.pack(fill="both", expand=True)

        self.month_cards = []
        for i in range(5):
            self.body.grid_columnconfigure(i, weight=(1 if i in (0,2,4) else 0))
        self.body.grid_rowconfigure(0, weight=1)

        def make_card(col):
            card = tk.Frame(self.body, bg="#FFFFFF",
                            highlightthickness=1, highlightbackground="#E6E6E6",
                            highlightcolor="#E6E6E6", bd=0)
            card.grid(row=0, column=col, sticky="nsew", padx=2, pady=2)
            inner = ttk.Frame(card, padding=8)
            inner.pack(fill="both", expand=True)
            self.month_cards.append(inner)

        make_card(0)
        ttk.Separator(self.body, orient="vertical").grid(row=0, column=1, sticky="ns", padx=(8,8))
        make_card(2)
        ttk.Separator(self.body, orient="vertical").grid(row=0, column=3, sticky="ns", padx=(8,8))
        make_card(4)

    def prev_three(self):
        m = self.base_month - 3
        y = self.base_year
        while m <= 0:
            m += 12
            y -= 1
        self.base_month, self.base_year = m, y
        self.render_all()

    def next_three(self):
        m = self.base_month + 3
        y = self.base_year
        while m > 12:
            m -= 12
            y += 1
        self.base_month, self.base_year = m, y
        self.render_all()

    def render_all(self):
        left = date(self.base_year, self.base_month, 1)
        m2 = self.base_month + 2
        y2 = self.base_year + (m2 - 1) // 12
        m2 = ((m2 - 1) % 12) + 1
        right = date(y2, m2, 1)
        self.lbl_range.config(text=f"{left.strftime('%B %Y')} \u2192 {right.strftime('%B %Y')}")

        y, m = self.base_year, self.base_month
        for inner in self.month_cards:
            for w in inner.winfo_children():
                w.destroy()
            self._render_one_month(inner, y, m, self.show_week_numbers)
            m += 1
            if m > 12:
                m = 1
                y += 1

    def _render_one_month(self, parent, year, month, show_week_numbers):
        ttk.Label(parent, text=date(year, month, 1).strftime("%B %Y"),
                  style="Header.TLabel", anchor="center").pack(fill="x", pady=(0,6))

        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True)

        dow = ttk.Frame(container)
        dow.grid(row=0, column=0, sticky="ew")
        if show_week_numbers:
            ttk.Label(dow, text="Wk", style="WeekNumHeader.TLabel", anchor="center").grid(row=0, column=0, sticky="nsew", padx=(0,4))
            start_col = 1
        else:
            start_col = 0
        for i, w in enumerate(WEEKDAYS):
            style_name = "DowWeekend.TLabel" if i >= 5 else "Dow.TLabel"
            ttk.Label(dow, text=w, style=style_name, anchor="center").grid(row=0, column=i+start_col, sticky="nsew", padx=2, pady=(0,4))
            dow.grid_columnconfigure(i+start_col, weight=1)

        grid = ttk.Frame(container)
        grid.grid(row=1, column=0, sticky="nsew")
        container.grid_rowconfigure(1, weight=1)
        container.grid_columnconfigure(0, weight=1)

        cal = calendar.Calendar(firstweekday=0)  # Monday start
        month_weeks = cal.monthdatescalendar(year, month)

        for r in range(6):
            if show_week_numbers:
                wk_label = ttk.Label(grid, text="", style="WeekNum.TLabel", anchor="center")
                wk_label.configure(background="#FFF9C4")
                wk_label.grid(row=r, column=0, sticky="nsew", padx=(0,4), pady=2, ipadx=3)
            for c in range(7):
                col_index = c + (1 if show_week_numbers else 0)
                lbl = ttk.Label(grid, text="", style="Day.TLabel", anchor="center")
                lbl.grid(row=r, column=col_index, sticky="nsew", padx=2, pady=2, ipadx=4, ipady=4)
                grid.grid_columnconfigure(col_index, weight=1)
            grid.grid_rowconfigure(r, weight=1)

        for r, week in enumerate(month_weeks[:6]):
            if show_week_numbers:
                iso_week = week[0].isocalendar()[1]
                grid.grid_slaves(row=r, column=0)[0].config(text=str(iso_week))
            for c, d in enumerate(week):
                col_index = c + (1 if show_week_numbers else 0)
                lbl = grid.grid_slaves(row=r, column=col_index)[0]
                lbl.config(text=str(d.day))
                if d.month != month:
                    lbl.config(foreground="#C8C8C8")
                else:
                    lbl.config(foreground="#333333" if c < 5 else "#9E9E9E")
                    if c >= 5:
                        lbl.config(style="Weekend.TLabel")
                def on_enter(e, dd=d, cc=c):
                    e.widget.config(style="Hover.TLabel")
                def on_leave(e, dd=d, cc=c):
                    if dd.month != month:
                        e.widget.config(style="Day.TLabel", foreground="#C8C8C8")
                    else:
                        e.widget.config(style=("Weekend.TLabel" if cc >= 5 else "Day.TLabel"),
                                        foreground=("#9E9E9E" if cc >= 5 else "#333333"))
                lbl.bind("<Enter>", on_enter)
                lbl.bind("<Leave>", on_leave)


class UKCalendarApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.pack(fill="both", expand=True)

        self.today = date.today()
        self.current_year = self.today.year
        self.current_month = self.today.month
        self.selected_date = None
        self.highlighted_result = None  # from working-days calc

        # --- New: range selection state ---
        self.range_start = None  # type: date | None
        self.range_end = None    # type: date | None

        self.show_week_numbers = False

        self._build_style()
        self._build_menu()
        self._build_header()
        self._build_calendar_area()

        self.master.bind("<Left>", lambda e: self.prev_month())
        self.master.bind("<Right>", lambda e: self.next_month())
        self.master.bind("t", lambda e: self.go_today())
        self.master.bind("<Escape>", lambda e: self.reset_to_default())

        self.render_month()

    def _build_style(self):
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")

        base_font = ("Segoe UI", 10)
        italic_font = ("Segoe UI", 10, "italic")

        style.configure("Day.TLabel", padding=4, anchor="center", font=base_font)
        style.configure("Dow.TLabel", foreground="#555555", padding=4, font=base_font)
        style.configure("DowWeekend.TLabel", foreground="#9A9A9A", padding=4, font=italic_font)
        style.configure("WeekNum.TLabel", foreground="#003366", background="#FFF9C4", font=base_font, padding=(2,1))
        style.configure("WeekNumHeader.TLabel", foreground="#003366", background="#FFF9C4", font=base_font)

        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Today.TLabel", background="#9eccff", font=base_font)
        style.configure("Selected.TLabel", background="#DDEEFF", font=base_font)
        style.configure("Result.TLabel", background="#D9F7BE", font=base_font)

        style.configure("Weekend.TLabel", foreground="#B3B3B3", font=italic_font)
        style.configure("Hover.TLabel", background="#FFF7DA", font=base_font, anchor="center")

        # --- New: styles for range highlighting ---
        style.configure("Range.TLabel", background="#FFECC7", font=base_font)      # general in-range
        style.configure("RangeEdge.TLabel", background="#FFD89A", font=base_font)  # start/end

        style.configure("Nav.TButton", padding=(8,4))
        style.configure("Status.TLabel", foreground="#666666")

    def _build_menu(self):
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        m_file = tk.Menu(menubar, tearoff=False)
        m_file.add_command(label="Go to Today\tT", command=self.go_today)
        m_file.add_separator()
        m_file.add_command(label="Exit", command=self.master.destroy)
        menubar.add_cascade(label="File", menu=m_file)

        m_tools = tk.Menu(menubar, tearoff=False)
        m_tools.add_command(label="Add Working Days…", command=self.open_working_days_dialog)
        m_tools.add_command(label="Show 3-Month View…", command=self.open_three_month_view)
        menubar.add_cascade(label="Tools", menu=m_tools)

        m_view = tk.Menu(menubar, tearoff=False)
        self._weeknum_var = tk.BooleanVar(value=self.show_week_numbers)
        m_view.add_checkbutton(label="Show Week Numbers", onvalue=True, offvalue=False,
                               variable=self._weeknum_var, command=self._toggle_week_numbers)
        # Optional helper to clear range quickly
        m_view.add_separator()
        m_view.add_command(label="Clear Range Selection", command=self._clear_range_selection)
        menubar.add_cascade(label="View", menu=m_view)

    def _toggle_week_numbers(self):
        self.show_week_numbers = bool(self._weeknum_var.get())
        self._rebuild_calendar_area()

    def _build_header(self):
        hdr = ttk.Frame(self, padding=(8, 8, 8, 0))
        hdr.pack(fill="x")

        ttk.Button(hdr, text="◀", style="Nav.TButton", command=self.prev_month, width=3).pack(side="left")
        self.lbl_month = ttk.Label(hdr, text="", style="Header.TLabel", anchor="center")
        self.lbl_month.pack(side="left", expand=True)
        ttk.Button(hdr, text="▶", style="Nav.TButton", command=self.next_month, width=3).pack(side="right")

    def _build_calendar_area(self):
        self.calendar_area = ttk.Frame(self, padding=(8,8,8,8))
        self.calendar_area.pack(fill="both", expand=True)

        self.dow = ttk.Frame(self.calendar_area)
        self.dow.pack(fill="x")
        if self.show_week_numbers:
            ttk.Label(self.dow, text="Wk", style="WeekNumHeader.TLabel", anchor="center").grid(row=0, column=0, sticky="nsew", padx=(0,4))
            start_col = 1
        else:
            start_col = 0
        for i, w in enumerate(WEEKDAYS):
            style_name = "DowWeekend.TLabel" if i >= 5 else "Dow.TLabel"
            ttk.Label(self.dow, text=w, style=style_name, anchor="center").grid(row=0, column=i+start_col, sticky="nsew", padx=2, pady=(0,6))
            self.dow.grid_columnconfigure(i+start_col, weight=1)

        self.grid_frame = ttk.Frame(self.calendar_area)
        self.grid_frame.pack(fill="both", expand=True)

        self.day_labels = []
        self.weeknum_labels = []
        for r in range(6):
            if self.show_week_numbers:
                wk_label = ttk.Label(self.grid_frame, text="", style="WeekNum.TLabel", anchor="center")
                wk_label.configure(background="#FFF9C4")
                wk_label.grid(row=r, column=0, sticky="nsew", padx=(0,4), pady=2, ipadx=3)
                self.weeknum_labels.append(wk_label)

            row_labels = []
            for c in range(7):
                col_index = c + (1 if self.show_week_numbers else 0)
                lbl = ttk.Label(self.grid_frame, text="", style="Day.TLabel", anchor="center")
                lbl.grid(row=r, column=col_index, sticky="nsew", padx=2, pady=2, ipadx=6, ipady=6)
                lbl.bind("<Button-1>", self._on_day_click)
                lbl.bind("<Enter>", self._on_day_enter)
                lbl.bind("<Leave>", self._on_day_leave)
                row_labels.append(lbl)
                self.grid_frame.grid_columnconfigure(col_index, weight=1)
            self.day_labels.append(row_labels)
            self.grid_frame.grid_rowconfigure(r, weight=1)

        self.status = ttk.Label(self, text="", style="Status.TLabel", anchor="w")
        self.status.pack(fill="x", padx=8, pady=(0,6))

    def reset_to_default(self):
        # Clear selections and highlights
        self.range_start = None
        self.range_end = None
        self.selected_date = None
        self.highlighted_result = None
        # Reset to today's month
        self.today = date.today()
        self.current_year = self.today.year
        self.current_month = self.today.month
        # Reset view options to defaults
        if self.show_week_numbers:
            self.show_week_numbers = False
            if hasattr(self, '_weeknum_var'):
                try:
                    self._weeknum_var.set(False)
                except Exception:
                    pass
            # rebuild calendar grid to hide week numbers
            self._rebuild_calendar_area()
        else:
            self.render_month()

    def _rebuild_calendar_area(self):
        self.calendar_area.destroy()
        self.status.destroy()
        self._build_calendar_area()
        self.render_month()

    def _style_for_day(self, d: date, col_idx: int) -> str:
        # Range highlighting takes precedence
        if self.range_start and self.range_end:
            a = self.range_start if self.range_start <= self.range_end else self.range_end
            b = self.range_end if self.range_end >= self.range_start else self.range_start
            if a <= d <= b:
                if d == a or d == b:
                    base = "RangeEdge.TLabel"
                else:
                    base = "Range.TLabel"
            else:
                base = None
        else:
            base = None

        if base is None:
            if d == self.today:
                base = "Today.TLabel"
            elif self.highlighted_result == d:
                base = "Result.TLabel"
            elif self.selected_date == d:
                base = "Selected.TLabel"
            else:
                base = "Day.TLabel"

        if d.month != self.current_month:
            return base
        if col_idx >= 5 and base == "Day.TLabel":
            return "Weekend.TLabel"
        return base

    def render_month(self):
        cal = calendar.Calendar(firstweekday=0)
        month_days = cal.monthdatescalendar(self.current_year, self.current_month)

        month_name = date(self.current_year, self.current_month, 1).strftime("%B %Y")
        self.lbl_month.config(text=month_name)

        for r in range(6):
            if self.show_week_numbers:
                self.weeknum_labels[r].config(text="")
            for c in range(7):
                lbl = self.day_labels[r][c]
                lbl.config(text="", style="Day.TLabel", foreground="#000000")
                lbl.date_value = None

        for r, week in enumerate(month_days[:6]):
            if self.show_week_numbers:
                iso_week = week[0].isocalendar()[1]
                self.weeknum_labels[r].config(text=str(iso_week))
            for c, day in enumerate(week):
                lbl = self.day_labels[r][c]
                lbl.config(text=str(day.day))
                lbl.date_value = day

                if day.month != self.current_month:
                    lbl.config(foreground="#C8C8C8")
                else:
                    lbl.config(foreground="#333333" if c < 5 else "#9E9E9E")

                lbl.config(style=self._style_for_day(day, c))

        # Status text
        if self.range_start and self.range_end:
            a = self.range_start if self.range_start <= self.range_end else self.range_end
            b = self.range_end if self.range_end >= self.range_start else self.range_start
            total_days = (b - a).days + 1
            workdays = working_days_between(a, b)
            self.status.config(text=f"Range: {a.strftime(DATE_FMT)} → {b.strftime(DATE_FMT)} | Days: {total_days} | Working days: {workdays}")
        elif self.range_start and not self.range_end:
            self.status.config(text=f"Range start: {self.range_start.strftime(DATE_FMT)} (click another date to complete range)")
        elif self.selected_date:
            self.status.config(text=f"Selected: {self.selected_date.strftime(DATE_FMT)}")
        else:
            self.status.config(text=f"Today: {self.today.strftime(DATE_FMT)}")

    def prev_month(self):
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self.render_month()

    def next_month(self):
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self.render_month()

    def go_today(self):
        self.current_year = self.today.year
        self.current_month = self.today.month
        self.selected_date = self.today
        self.render_month()

    def _clear_range_selection(self):
        self.range_start = None
        self.range_end = None
        self.render_month()

    def _on_day_click(self, event):
        lbl = event.widget
        d = getattr(lbl, "date_value", None)
        if d is None:
            return

        # If a full range already exists, start a new one
        if self.range_start is not None and self.range_end is not None:
            self.range_start = d
            self.range_end = None
        # If starting a range
        elif self.range_start is None:
            self.range_start = d
            self.range_end = None
        # If completing the range
        else:  # range_start set, range_end None
            self.range_end = d

        # Keep selected_date for single-date highlight while picking
        self.selected_date = d

        # If clicked date is in another month, navigate there (so user sees the range end)
        if d.month != self.current_month or d.year != self.current_year:
            self.current_month = d.month
            self.current_year = d.year

        self.render_month()

    def _on_day_enter(self, event):
        lbl = event.widget
        d = getattr(lbl, "date_value", None)
        if d is None:
            return
        lbl.config(style="Hover.TLabel")

    def _on_day_leave(self, event):
        lbl = event.widget
        d = getattr(lbl, "date_value", None)
        if d is None:
            return
        info = lbl.grid_info()
        col_idx = int(info.get("column", 0))
        if self.show_week_numbers:
            col_idx -= 1
        if col_idx < 0:
            col_idx = 0
        lbl.config(style=self._style_for_day(d, col_idx))

    def open_working_days_dialog(self):
        def on_done(start, days, result):
            self.highlighted_result = result
            self.current_year = result.year
            self.current_month = result.month
            self.render_month()
            messagebox.showinfo(
                "Result",
                f"Start: {start.strftime(DATE_FMT)}\n"
                f"Working days: {days}\n"
                f"Result: {result.strftime(DATE_FMT)}"
            )
        WorkingDaysDialog(self.master, on_done)

    def open_three_month_view(self):
        ThreeMonthWindow(self.master, style_helper=self, show_week_numbers=self.show_week_numbers)

def main():
    _create_blank_icon_if_needed(ICON_PATH)

    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("440x460")
    root.minsize(360, 360)

    try:
        if os.path.isfile(ICON_PATH):
            root.iconbitmap(default=ICON_PATH)
    except Exception:
        pass

    app = UKCalendarApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
