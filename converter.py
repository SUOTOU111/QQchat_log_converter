# -*- coding: utf-8 -*-
"""
群聊记录 -> 跑团记录 格式转换（带 GUI）
运行：python convert.py
作者锁头（3229622745）
"""
import colorsys
import random
import re
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ==================== 1. 配置区（只改这里） ====================

# 正文命中这些就整条丢掉（图片 / 表情 / 语音 / 空消息）
DROP_PATTERNS = [
    # ⚠️ 顺序敏感：宽泛的模式必须排在具体标记之后
    r"<\s*img\b[^>]*/?>",                              # HTML 图片标签（含自闭合、含换行）
    r"<\s*/?\s*(?:img|image|p|br|div|span|a)\b[^>]*>",  # 其它 HTML 残留标签
    r"file:///?" ,                                     # 本地文件链接
    r"\[聊天记录\]", r"\[转发[^]]*\]", r"\[合并转发[^]]*\]",   # 转发/聊天记录
    r"\.(?:png|jpe?g|gif|bmp|webp|heic|tiff?)(?:\s|$)",  # 裸图片文件名
    r"\s*",                                            # 空消息
    r"\[图片\]", r"\[图片消息\]",
    r"\[表情\]", r"\[动画表情\]",
    r"\[语音\]", r"\[视频\]", r"\[文件\]",
    r"\[链接\]", r"\[QQ红包\]", r"\[转账\]",
]
# 消息头：昵称 + 冒号 + 日期 + 时间
HEAD_RE = re.compile(
    r"([^\s:：\n]{1,20})[:：]\s*(\d{1,2}-\d{1,2})\s*(\d{1,2}:\d{2}:\d{2})"
)

CN_FONT = "宋体"      # 中文字体（eastAsia）
EN_FONT = "Calibri"   # 西文/数字字体（ascii、hAnsi）
SIZE_PT = 10.5        # 五号

# 随机配色候选；名字按出现顺序依次取，用完循环
PALETTE = [
    "C00000", "1F4E79", "375623", "7030A0", "BF8F00",
    "2E75B6", "C55A11", "7F7F7F", "2F5597", "A64D79",
]
# 续行缩进；段后间距沿用 set_para_format 里的 6 磅
CONT_INDENT = Cm(1.41)

# 断行点：句末标点/引号 后跟 空格或换行；冒号、逗号不在其中，不会误断
BREAK_RE = re.compile(r'(?<=[。？！”?!])(?:[ \u3000\t]+|\n+)')

def gen_hex_colors(n, seed=None):
    """生成 n 个区分度高的随机色，返回 6 位 HEX（不带 #）
    色相按黄金角 0.618 递进，相邻角色色差拉大；明度压在 0.80 以下，白底才看得清"""
    rnd = random.Random(seed)
    start = rnd.random()
    out = []
    for i in range(n):
        h = (start + i * 0.618033988749895) % 1.0
        s = rnd.uniform(0.62, 0.88)
        v = rnd.uniform(0.50, 0.80)
        r, g, b = colorsys.hsv_to_rgb(h, s, v)
        out.append("%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255)))
    return out


def normalize_hex(s):
    """规范成 6 位大写 HEX（不带 #）；非法返回 None
    支持 DB2777 / #DB2777 / #db2777 / #F0C"""
    if not s:
        return None
    s = s.strip().lstrip("#").strip()
    if len(s) == 3 and all(c in "0123456789abcdefABCDEF" for c in s):
        s = "".join(c * 2 for c in s)
    if len(s) == 6 and all(c in "0123456789abcdefABCDEF" for c in s):
        return s.upper()
    return None
# ==================== 2. 解析 ====================

def read_text(path):
    """读 docx 或 txt，返回一整串文本"""
    p = Path(path)
    if p.suffix.lower() == ".docx":
        doc = Document(str(p))
        parts = [para.text for para in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    parts.append(cell.text)
        return "\n".join(parts)
    for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
        try:
            return p.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    raise RuntimeError("无法识别文件编码：" + str(path))


def split_messages(text):
    """切成 [(昵称, 日期, 时间, 内容), ...]；内容保留内部换行"""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    heads = list(HEAD_RE.finditer(text))
    out = []
    for i, m in enumerate(heads):
        start = m.end()
        end = heads[i + 1].start() if i + 1 < len(heads) else len(text)
        body = text[start:end].strip("\n")   # 只去首尾空行，内部换行保留
        out.append((m.group(1), m.group(2), m.group(3), body))
    return out


def body_dropped(body):
    """正文是图片/表情/空消息则丢弃"""
    b = body.strip()
    for pat in DROP_PATTERNS:
        if re.fullmatch(pat, b):
            return True
    return False

def split_paragraphs(body):
    """按 句末标点+空白 拆成 [首段, 续行, ...]；无分行则只有 1 段"""
    parts = [p.strip() for p in BREAK_RE.split(body)]
    return [p for p in parts if p]

# ==================== 3. 排版 ====================

def set_run_font(run, color=None):
    """五号；中文宋体、西文数字 Calibri；不加粗；可带颜色"""
    run.font.size = Pt(SIZE_PT)
    run.font.bold = False
    run.font.name = EN_FONT

    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), EN_FONT)
    rFonts.set(qn("w:hAnsi"), EN_FONT)
    rFonts.set(qn("w:eastAsia"), CN_FONT)
    rFonts.set(qn("w:cs"), EN_FONT)

    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def set_para_format(p):
    """1 倍行距，段前 0，段后 6 磅"""
    pf = p.paragraph_format
    pf.line_spacing = 1         
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)


def _init_normal_style(doc):
    normal = doc.styles["Normal"]
    normal.font.name = EN_FONT
    normal.font.size = Pt(SIZE_PT)
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), CN_FONT)
    rFonts.set(qn("w:ascii"), EN_FONT)
    rFonts.set(qn("w:hAnsi"), EN_FONT)


            
def write_docx(lines, colors, out_path):
    doc = Document()

    # 设置页边距和纸张大小
    for section in doc.sections:
        # 【新增】强制设置为 A4 纸尺寸 (宽 21cm，高 29.7cm)
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        
        # 原有的页边距设置（上下左右各 2.54 cm）
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    # 初始化 Normal 样式
    _init_normal_style(doc)

    for name, body in lines:
        color = colors.get(name, "000000")
        paras = split_paragraphs(body)

        # 首段：紧跟 <名字>: 后面，不缩进
        p = doc.add_paragraph()
        set_para_format(p)
        prefix = "<" + name + ">:"
        set_run_font(p.add_run(prefix), color)
        if paras:
            set_run_font(p.add_run(paras[0]), color)

        # 续行：各自独立成段，缩进 1.41 cm，段后 6 磅
        for seg in paras[1:]:
            cp = doc.add_paragraph()
            set_para_format(cp)
            cp.paragraph_format.left_indent = CONT_INDENT
            set_run_font(cp.add_run(seg), color)

    doc.save(out_path)
    
def write_txt(lines, out_path):
    with open(out_path, "w", encoding="utf-8") as f:
        for name, body in lines:
            f.write("<" + name + ">:" + body + "\n")


# ==================== 4. 转换主逻辑 ====================

def convert(src, dst, name_map, color_map=None, drop_ooc=False):
    """name_map: 原昵称 -> 新昵称；值为 None 表示整条丢弃。
    drop_ooc=True 时，正文以（或(开头的整条消息连名字一起丢弃。
    返回 (总条数, 保留条数, 剔除场外条数)"""
    raw_msgs = split_messages(read_text(src))
    if not raw_msgs:
        raise ValueError("没解析到任何消息，检查输入格式是否为『昵称:日期 时间 内容』")

    kept = []
    used_color = {}          # 新昵称 -> 该角色的原昵称（用来查颜色）
    dropped_ooc = 0
    for name, _d, _t, body in raw_msgs:
        new_name = name_map.get(name, name)
        if new_name is None:          # 被清空 = 删除该角色
            continue
        if drop_ooc and body.lstrip()[:1] in ("（", "("):
            dropped_ooc += 1          # 冒号后紧跟左括号：整条连名字一起丢掉
            continue
        if body_dropped(body):        # 图片/表情/空消息
            continue
        kept.append((new_name, body))
        used_color.setdefault(new_name, name)

    # 按出场顺序自动分配颜色，同名同色
    # 先用 PALETTE 兜底，再用 GUI 传来的颜色覆盖
    order = list(dict.fromkeys(n for n, _ in kept))
    colors = {n: PALETTE[i % len(PALETTE)] for i, n in enumerate(order)}
    color_map = color_map or {}
    for n in order:
        c = normalize_hex(color_map.get(used_color.get(n, n)))
        if c:
            colors[n] = c

    if dst.lower().endswith(".docx"):
        write_docx(kept, colors, dst)
    else:
        write_txt(kept, dst)
    return len(raw_msgs), len(kept), dropped_ooc


# ==================== 5. GUI ====================

class ScrollFrame(ttk.Frame):
    """可滚动容器，昵称多了也能滚"""

    def __init__(self, master):
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=1, borderwidth=1,
                                background="white")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>",
                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._wheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _wheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")


class App:
    def __init__(self, root):
        self.root = root
        self.src = None
        self.msgs = []
        self.entries = {}        # 原昵称 -> StringVar（新名字）
        self.hex_vars = {}       # 原昵称 -> StringVar("#RRGGBB")
        self.swatches = {}       # 原昵称 -> 色块 Label
        self.default_colors = {} # 原昵称 -> 该行初始随机色（不带 #）
        
        self.del_vars = {}       # 原昵称 -> BooleanVar（勾选=删除该角色）
        self.name_entries = {}   # 原昵称 -> 名字输入框控件
        self.hex_entries = {}    # 原昵称 -> 颜色输入框控件
        self.ooc_var = tk.BooleanVar(value=True)   # True=默认勾上「去除场外」

        root.title("群聊记录转换器")
        root.geometry("980x640")

        # --- 顶部：选择文件 ---
        top = ttk.Frame(root, padding=(10, 10, 10, 4))
        top.pack(fill="x")
        ttk.Button(top, text="① 选择未处理记录…", command=self.choose_src).pack(side="left")
        self.src_var = tk.StringVar(value="（还没选文件）")
        ttk.Label(top, textvariable=self.src_var, foreground="#444").pack(
            side="left", padx=10, fill="x", expand=True)

        # --- 中部：昵称列表 ---
        mid = ttk.Frame(root, padding=(10, 4))
        mid.pack(fill="both", expand=True)
        ttk.Label(mid, text="② 确认/修改角色名：左边是原文档里的名字，右边填写想替换成的名字。"
                            "不改就直接留着；清空则删除该角色的全部消息。颜色按出场顺序自动分配。",
                  foreground="#555").pack(anchor="w", pady=(0, 6))
        self.scroll = ScrollFrame(mid)
        self.scroll.pack(fill="both", expand=True)
        self.count_var = tk.StringVar(value="")
        ttk.Label(mid, textvariable=self.count_var, foreground="#777").pack(anchor="w", pady=(6, 0))

        # --- 底部：转换 ---
        bot = ttk.Frame(root, padding=(10, 6, 10, 10))
        bot.pack(fill="x")
        ttk.Button(bot, text="③ 开始转换并保存…", command=self.do_convert).pack(side="right")
        ttk.Button(bot, text="全部还原为原名", command=self.reset_all).pack(side="right", padx=8)
        ttk.Button(bot, text="重新随机配色", command=self.reroll_colors).pack(side="right", padx=8)
        ttk.Checkbutton(bot, text="去除场外（删除冒号后紧跟左括号的整条消息）",
                        variable=self.ooc_var).pack(side="left")

    # ---------- 选文件 + 抽取昵称 ----------
    def choose_src(self):
        path = filedialog.askopenfilename(
            title="选择未处理记录",
            filetypes=[("Word / 文本", "*.docx *.txt"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            msgs = split_messages(read_text(path))
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        if not msgs:
            messagebox.showerror("读取失败",
                                 "没解析到消息。\n请确认文档内容是『昵称:日期 时间 内容』这种格式。")
            return
        self.src = path
        self.msgs = msgs
        self.src_var.set(path)
        self.build_name_rows()

    def build_name_rows(self):
        for w in self.scroll.inner.winfo_children():
            w.destroy()
        self.entries.clear()
        self.del_vars.clear()
        self.name_entries.clear()
        self.hex_entries.clear()
        self.hex_vars.clear()
        self.swatches.clear()
        self.default_colors.clear()

        seen = []
        for name, _d, _t, _b in self.msgs:
            if name not in seen:
                seen.append(name)

        palette = gen_hex_colors(len(seen))

        heads = ["删除", "原昵称", "替换为", "颜色 HEX", "色块"]
        for c, text in enumerate(heads):
            ttk.Label(self.scroll.inner, text=text, font=("", 9, "bold")).grid(
                row=0, column=c, sticky="w",
                padx=(8 if c <= 1 else 0, 8), pady=(6, 4))

        for i, name in enumerate(seen, start=1):
            color = palette[i - 1]
            self.default_colors[name] = color

            # 勾选框
            dvar = tk.BooleanVar(value=False)
            ttk.Checkbutton(self.scroll.inner, variable=dvar).grid(
                row=i, column=0, sticky="w", padx=(8, 4), pady=3)
            self.del_vars[name] = dvar

            ttk.Label(self.scroll.inner, text=name, width=22, anchor="w").grid(
                row=i, column=1, sticky="w", padx=(0, 8), pady=3)

            # 名字框
            var = tk.StringVar(value=name)
            name_ent = ttk.Entry(self.scroll.inner, textvariable=var, width=26)
            name_ent.grid(row=i, column=2, sticky="we", padx=(0, 8), pady=3)
            self.entries[name] = var
            self.name_entries[name] = name_ent

            # 颜色框
            hvar = tk.StringVar(value="#" + color)
            hex_ent = ttk.Entry(self.scroll.inner, textvariable=hvar, width=11)
            hex_ent.grid(row=i, column=3, sticky="w", padx=(0, 6), pady=3)
            self.hex_vars[name] = hvar
            self.hex_entries[name] = hex_ent

            # 色块
            sw = tk.Label(self.scroll.inner, text="", width=4, relief="solid",
                          borderwidth=1, background="#" + color, cursor="hand2")
            sw.grid(row=i, column=4, sticky="w", padx=(0, 8), pady=3)
            sw.bind("<Button-1>", lambda e, n=name: self.pick_color(n))
            self.swatches[name] = sw

            hvar.trace_add("write", lambda *a, n=name: self.refresh_swatch(n))
            dvar.trace_add("write", lambda *a, n=name: self.on_toggle_delete(n))

        self.scroll.inner.columnconfigure(2, weight=1)

        # 统计每个角色多少条，方便删掉刷屏的
        cnt = {}
        for name, _d, _t, _b in self.msgs:
            cnt[name] = cnt.get(name, 0) + 1
        detail = "、".join("%s %d条" % (n, cnt[n]) for n in seen)
        self.count_var.set("共识别到 %d 个角色、%d 条消息：%s"
                           % (len(seen), len(self.msgs), detail))

    def reset_all(self):
        for name, var in self.entries.items():
            var.set(name)
        for name, dvar in self.del_vars.items():
            dvar.set(False)


    def on_toggle_delete(self, name):
        """勾选删除：禁用该行两个输入框、色块变灰；取消勾选恢复"""
        on = self.del_vars[name].get()
        state = "disabled" if on else "normal"
        self.name_entries[name].configure(state=state)
        self.hex_entries[name].configure(state=state)
        if on:
            self.swatches[name].configure(background="#CCCCCC")
        else:
            self.refresh_swatch(name)
            
    def refresh_swatch(self, name):
        """手打 HEX 时实时更新色块；非法就变灰，但不清空输入"""
        h = normalize_hex(self.hex_vars[name].get())
        self.swatches[name].configure(
            background="#" + h if h else "#E6E6E6")

    def get_hex(self, name):
        """最终取色：输入合法用输入，否则回退到初始随机色"""
        return (normalize_hex(self.hex_vars[name].get())
                or self.default_colors[name])

    def pick_color(self, name):
        """点色块弹系统调色板，初始色就是当前色"""
        rgb, hexstr = colorchooser.askcolor(
            color="#" + self.get_hex(name), title="选择颜色 — " + name)
        if hexstr:
            self.hex_vars[name].set(hexstr.upper())

    def reroll_colors(self):
        """整批重洗随机色"""
        if not self.entries:
            return
        names = list(self.entries.keys())
        for n, c in zip(names, gen_hex_colors(len(names))):
            self.default_colors[n] = c
            self.hex_vars[n].set("#" + c)

    # ---------- 执行转换 ----------
    def do_convert(self):
        if not self.src:
            messagebox.showwarning("提示", "先点『① 选择未处理记录』选文件。")
            return

        name_map, color_map = {}, {}
        for name, var in self.entries.items():
            if self.del_vars[name].get():        # 勾了删除，优先级最高
                name_map[name] = None
                continue
            v = var.get().strip()
            name_map[name] = None if v == "" else v
            color_map[name] = self.get_hex(name)

        # 全被删光时提醒一下，别生成一个空文档才发现
        if name_map and all(v is None for v in name_map.values()):
            if not messagebox.askyesno("确认", "所有角色都被删除了，输出会是一份空文档。继续？"):
                return

        src_path = Path(self.src)
        dst = filedialog.asksaveasfilename(
            title="保存为",
            defaultextension=".docx",
            initialdir=str(src_path.parent),
            initialfile=src_path.stem + "_已处理.docx",
            filetypes=[("Word 文档", "*.docx"), ("文本文件", "*.txt")])
        if not dst:
            return

        try:
            total, kept, n_ooc = convert(self.src, dst, name_map, color_map,
                                         self.ooc_var.get())
        except Exception as e:
            messagebox.showerror("转换失败", str(e))
            return

        messagebox.showinfo("完成",
                            "解析 %d 条，保留 %d 条，剔除场外 %d 条。\n已保存到：\n%s"
                            % (total, kept, n_ooc, dst))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
