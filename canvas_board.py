# canvas_board.py
# ตัวช่วยวาดหน้ากระดาน/รายงานแบบ Canvas (วาดข้อความ+กรอบโดยตรง ไม่สร้างวิดเจ็ตต่อชิ้น)
# เหตุผล: วิดเจ็ต Tk/CustomTkinter ต้องจัดวางทีละชิ้น ทำให้หน้าที่มีวิดเจ็ตหลายร้อยชิ้นใช้เวลาโหลดเป็นสิบวินาที
#         (วัดแล้ว 41 การ์ด ≈ 35 วินาที) ส่วน Canvas วาด 1,000 ข้อความใช้ไม่ถึง 0.1 วินาที
import re
import unicodedata
import tkinter as tk
import tkinter.font as tkfont

# ตัวอักษรเสียที่ค้างในฐานข้อมูล (private-use / ช่วง U+F0000 ขึ้นไป) — เกิดจากบั๊กคลิปบอร์ดเดิม แสดงเป็นกล่องสี่เหลี่ยม จึงตัดออกตอนแสดงผล
_PUA = re.compile("[\ue000-\uf8ff\U000f0000-\U0010ffff]")


def clean_text(s):
    """ตัดตัวอักษรเสียออกจากข้อความที่จะวาด (ไม่แก้ข้อมูลในฐานข้อมูล)"""
    return _PUA.sub("", s) if isinstance(s, str) else s


class BoardCanvas(tk.Frame):
    """Canvas + แถบเลื่อนแนวตั้ง + ฟังก์ชันวาดการ์ด/ป้าย/แถบความคืบหน้า (หน่วยพิกเซลคูณอัตราส่วนจอ S())"""

    # Segoe UI (มากับ Windows ทุกเครื่อง, ไทยใช้ Leelawadee UI เป็นฟอนต์สำรองอัตโนมัติ) — ทดสอบแล้ว Tahoma/Leelawadee UI ทำให้สระ ื ู ในบางบรรทัดแสดงเป็นวงกลมประ
    FONT_FAMILY = "Segoe UI"

    def __init__(self, master, bg="#F3F4F6", on_resize=None):
        super().__init__(master, bg=bg)
        self.bg = bg
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.vsb = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self._fonts = {}
        self._k = None
        self._on_resize = on_resize
        self._resize_job = None
        self._last_w = 0
        self.canvas.bind("<Configure>", self._configure)
        self.canvas.bind("<Enter>", lambda e: self._bind_wheel(True))
        self.canvas.bind("<Leave>", lambda e: self._bind_wheel(False))

    # ---------------------------------------------------------------- scale / fonts
    @property
    def k(self):
        if self._k is None:
            try:
                self._k = max(self.winfo_fpixels("1i") / 96.0, 1.0)
            except Exception:
                self._k = 1.0
        return self._k

    def S(self, v):
        return int(round(v * self.k))

    def font(self, size=11, bold=False):
        key = (size, bold)
        f = self._fonts.get(key)
        if f is None:
            f = tkfont.Font(family=self.FONT_FAMILY, size=size, weight="bold" if bold else "normal")
            self._fonts[key] = f
        return f

    @property
    def width(self):
        return max(self.canvas.winfo_width(), 300)

    # ---------------------------------------------------------------- events
    def _configure(self, e):
        if abs(e.width - self._last_w) < 4:
            return
        self._last_w = e.width
        if self._on_resize:
            if self._resize_job:
                self.after_cancel(self._resize_job)
            self._resize_job = self.after(120, self._on_resize)

    def _bind_wheel(self, on):
        top = self.winfo_toplevel()
        if on:
            top.bind_all("<MouseWheel>", self._wheel)
        else:
            top.unbind_all("<MouseWheel>")

    def _wheel(self, e):
        try:
            self.canvas.yview_scroll(int(-1 * (e.delta / 120) * 3), "units")
        except Exception:
            pass

    # ---------------------------------------------------------------- drawing
    def clear(self):
        self.canvas.delete("all")

    def set_height(self, h):
        w = self.width
        self.canvas.configure(scrollregion=(0, 0, w, max(h, 1)))

    def wrap_lines(self, txt, font, max_px):
        """ตัดบรรทัดเอง: ตัดที่ช่องว่างก่อน, คำยาวเกินจึงตัดที่ขอบพยัญชนะ (ไม่ตัดระหว่างพยัญชนะกับสระ/วรรณยุกต์)
        เหตุผล: ตัวเลือก width ของ Canvas แบ่งชิ้นข้อความกลางคำไทย ทำให้สระ/วรรณยุกต์ไปขึ้นต้นชิ้นแล้วแสดงเป็นวงกลมประ"""
        lines = []
        for para in str(txt).split("\n"):
            cur = ""
            for word in para.split(" "):
                trial = (cur + " " + word) if cur else word
                if font.measure(trial) <= max_px:
                    cur = trial
                    continue
                if cur:
                    lines.append(cur)
                    cur = ""
                if font.measure(word) <= max_px:
                    cur = word
                    continue
                buf = ""
                for ch in word:
                    if buf and unicodedata.category(ch) != "Mn" and font.measure(buf + ch) > max_px:
                        lines.append(buf)
                        buf = ch
                    else:
                        buf += ch
                cur = buf
            lines.append(cur)
        return lines

    def text(self, x, y, txt, size=11, bold=False, fill="#111827", anchor="nw", width=None, tags=(), justify="left"):
        font = self.font(size, bold)
        if width:
            txt = "\n".join(self.wrap_lines(txt, font, int(width)))
        item = self.canvas.create_text(x, y, text=txt, font=font, fill=fill, anchor=anchor, tags=tags, justify=justify)
        return item, self.canvas.bbox(item)

    @staticmethod
    def _rr_points(x1, y1, x2, y2, r):
        r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
        return [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
                x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]

    def rrect(self, x1, y1, x2, y2, r=10, fill="white", outline="", tags=(), width=1):
        return self.canvas.create_polygon(self._rr_points(x1, y1, x2, y2, r), smooth=True, fill=fill,
                                          outline=outline, width=width, tags=tags)

    def set_rrect(self, item, x1, y1, x2, y2, r=10):
        """ปรับขนาดกรอบมุมมนที่วาดไว้แล้ว (ใช้กับการ์ดที่ความสูงรู้หลังวาดเนื้อหา)"""
        self.canvas.coords(item, *self._rr_points(x1, y1, x2, y2, r))

    def rect(self, x1, y1, x2, y2, fill="white", outline="", tags=(), width=1):
        return self.canvas.create_rectangle(x1, y1, x2, y2, fill=fill, outline=outline, tags=tags, width=width)

    def line(self, x1, y1, x2, y2, fill="#E5E7EB", width=1):
        return self.canvas.create_line(x1, y1, x2, y2, fill=fill, width=width)

    def badge(self, x, y, txt, bg, fg, size=10, bold=True, anchor="nw", tags=(), padx=8, pady=3):
        """ป้ายมุมมน; anchor 'nw' = (x,y) เป็นมุมซ้ายบน, 'ne' = มุมขวาบน. คืน (กว้าง, สูง)"""
        f = self.font(size, bold)
        tw = f.measure(str(txt))
        th = f.metrics("linespace")
        w, h = tw + self.S(padx) * 2, th + self.S(pady) * 2
        x1 = x if anchor.endswith("w") else x - w
        self.rrect(x1, y, x1 + w, y + h, r=h / 2, fill=bg, tags=tags)
        self.canvas.create_text(x1 + w / 2, y + h / 2, text=str(txt), font=f, fill=fg, tags=tags)
        return w, h

    def bar(self, x, y, w, h, pct, color="#1D4ED8", track="#E5E7EB"):
        self.rrect(x, y, x + w, y + h, r=h / 2, fill=track)
        pct = max(0.0, min(float(pct or 0), 1.0))
        if pct > 0:
            self.rrect(x, y, x + max(w * pct, h), y + h, r=h / 2, fill=color)

    def tag_click(self, tag, callback):
        self.canvas.tag_bind(tag, "<Button-1>", lambda e: callback())
        self.canvas.tag_bind(tag, "<Enter>", lambda e: self.canvas.configure(cursor="hand2"))
        self.canvas.tag_bind(tag, "<Leave>", lambda e: self.canvas.configure(cursor=""))
