#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
大连理工大学本科毕业论文格式自动审查工具
依据：《大连理工大学大学生毕业设计（论文）规范化要求》
"""

import sys
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Optional

# ──────────────────────────────────────────────
# 大连理工标准值（twips 换算：1cm = 566.929 twips）
# ──────────────────────────────────────────────
TWIPS_PER_CM = 566.929

DUT = {
    "page": {
        "width_twips":  11906,   # A4 210mm
        "height_twips": 16838,   # A4 297mm
        "top_cm":    3.5,        # 上边距
        "bottom_cm": 2.5,        # 下边距
        "left_cm":   2.5,        # 左边距
        "right_cm":  2.5,        # 右边距
        "top_twips":    1984,    # 3.5cm
        "bottom_twips": 1417,    # 2.5cm
        "left_twips":   1417,    # 2.5cm
        "right_twips":  1417,    # 2.5cm
        "tolerance":     57,     # ±0.1cm 容差
    },
    "body": {
        "size_halfpt": 24,       # 小四 = 12pt = 24 half-pt
        "line_twips":  300,      # 1.25倍行距 (240*1.25)
        "font_cn":     "宋体",
        "font_en":     "Times New Roman",
        "ind_chars":   2,        # 首行缩进2字符
    },
    "heading1": {
        "font":        "黑体",
        "size_halfpt": 30,       # 小三 = 15pt
        "align":       "left",
        "line_twips":  360,      # 1.5倍
    },
    "heading2": {
        "font":        "黑体",
        "size_halfpt": 28,       # 四号 = 14pt
        "line_twips":  360,
    },
    "heading3": {
        "font":        "黑体",
        "size_halfpt": 24,       # 小四 = 12pt
        "line_twips":  360,
    },
    "word_count_min": 20000,
    "ref_min": 15,
}

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ──────────────────────────────────────────────
# 数据结构
# ──────────────────────────────────────────────
@dataclass
class Item:
    ok:       bool
    category: str
    name:     str
    expected: str
    actual:   str
    level:    str = "error"   # error | warn | info


# ──────────────────────────────────────────────
# 主检查类
# ──────────────────────────────────────────────
class Checker:

    def __init__(self, path: str):
        self.path = Path(path)
        self.results: List[Item] = []
        self._load()

    # ── 加载 XML ──────────────────────────────
    def _load(self):
        with zipfile.ZipFile(self.path) as z:
            def read(name):
                try:
                    with z.open(name) as f:
                        return ET.parse(f).getroot()
                except KeyError:
                    return None

            self.doc     = read("word/document.xml")
            self.styles  = read("word/styles.xml")
            self.settings = read("word/settings.xml")

        # 预提取所有段落文本（带样式ID）
        self._paras = []
        if self.doc is not None:
            for p in self.doc.iter(f"{{{WNS}}}p"):
                pPr   = p.find(f"{{{WNS}}}pPr")
                sid   = ""
                if pPr is not None:
                    ps = pPr.find(f"{{{WNS}}}pStyle")
                    if ps is not None:
                        sid = ps.get(f"{{{WNS}}}val", "")
                text  = "".join(t.text or "" for t in p.iter(f"{{{WNS}}}t"))
                sp    = p.find(f".//{{{WNS}}}spacing")
                ind   = p.find(f".//{{{WNS}}}ind")
                jc    = p.find(f".//{{{WNS}}}jc") if pPr is not None else None
                # 检测 Word 自动列表编号（numPr），提取 numId
                num_id = ""
                if pPr is not None:
                    numPr = pPr.find(f"{{{WNS}}}numPr")
                    if numPr is not None:
                        numId_el = numPr.find(f"{{{WNS}}}numId")
                        if numId_el is not None:
                            num_id = numId_el.get(f"{{{WNS}}}val", "")
                self._paras.append({
                    "style": sid,
                    "text":  text,
                    "spacing": sp.attrib if sp is not None else {},
                    "indent":  ind.attrib if ind is not None else {},
                    "jc":      (jc.get(f"{{{WNS}}}val","") if jc is not None else ""),
                    "num_id":  num_id,
                })

    # ── 辅助 ──────────────────────────────────
    def _t2cm(self, twips: int) -> float:
        return round(twips / TWIPS_PER_CM, 2)

    def _add(self, ok, cat, name, exp, act, level="error"):
        self.results.append(Item(ok, cat, name, exp, act, level))

    def _texts(self) -> List[str]:
        return [p["text"] for p in self._paras]

    def _full_text(self) -> str:
        return "\n".join(self._texts())

    # ── 1. 页面设置 ────────────────────────────
    def check_page(self):
        if self.doc is None:
            self._add(False, "页面设置", "页面XML", "可读", "document.xml 不可读")
            return

        pgSz  = self.doc.find(f".//{{{WNS}}}pgSz")
        # 取正文节的页边距：文档可能有多节，封面节边距不同，取第二个及以后的节
        all_pgMar = self.doc.findall(f".//{{{WNS}}}pgMar")
        if len(all_pgMar) > 1:
            pgMar = all_pgMar[1]   # 跳过封面节，取正文节
        else:
            pgMar = all_pgMar[0] if all_pgMar else None

        # 纸张大小
        if pgSz is not None:
            w = int(pgSz.get(f"{{{WNS}}}w", 0))
            h = int(pgSz.get(f"{{{WNS}}}h", 0))
            is_a4 = abs(w - DUT["page"]["width_twips"]) <= 20 and \
                    abs(h - DUT["page"]["height_twips"]) <= 20
            self._add(is_a4, "页面设置", "纸张大小",
                      "A4 (11906×16838 twips)",
                      f"实测 {w}×{h} twips ({self._t2cm(w)}×{self._t2cm(h)} cm)",
                      "error" if not is_a4 else "info")

        # 页边距
        if pgMar is not None:
            tol = DUT["page"]["tolerance"]
            margins = {
                "上边距": (f"{{{WNS}}}top",    DUT["page"]["top_twips"],    DUT["page"]["top_cm"]),
                "下边距": (f"{{{WNS}}}bottom",  DUT["page"]["bottom_twips"], DUT["page"]["bottom_cm"]),
                "左边距": (f"{{{WNS}}}left",    DUT["page"]["left_twips"],   DUT["page"]["left_cm"]),
                "右边距": (f"{{{WNS}}}right",   DUT["page"]["right_twips"],  DUT["page"]["right_cm"]),
            }
            for label, (attr, std_twips, std_cm) in margins.items():
                val = int(pgMar.get(attr, 0))
                ok  = abs(val - std_twips) <= tol
                self._add(ok, "页面设置", label,
                          f"{std_cm} cm ({std_twips} twips)",
                          f"{self._t2cm(val)} cm ({val} twips)")
        else:
            self._add(False, "页面设置", "页边距", "可读取", "pgMar 节点未找到", "warn")

    # ── 2. 默认段落样式 ────────────────────────
    def check_default_style(self):
        if self.styles is None:
            self._add(False, "段落样式", "styles.xml", "可读", "无法读取", "warn")
            return

        # 找默认段落样式
        default_style = None
        for s in self.styles.iter(f"{{{WNS}}}style"):
            if s.get(f"{{{WNS}}}type") == "paragraph" and \
               s.get(f"{{{WNS}}}default") == "1":
                default_style = s
                break

        if default_style is None:
            self._add(False, "段落样式", "默认段落", "可找到", "未找到默认段落样式", "warn")
            return

        # 字号
        sz = default_style.find(f".//{{{WNS}}}sz")
        if sz is not None:
            val = int(sz.get(f"{{{WNS}}}val", 0))
            ok  = val == DUT["body"]["size_halfpt"]
            self._add(ok, "正文样式", "字号",
                      f"小四 (24 half-pt / 12pt)",
                      f"实测 {val} half-pt ({val//2}pt)")
        else:
            self._add(False, "正文样式", "字号", "小四(24 half-pt)", "未设置", "warn")

        # 行距
        sp = default_style.find(f".//{{{WNS}}}spacing")
        if sp is not None:
            line = int(sp.get(f"{{{WNS}}}line", 0))
            rule = sp.get(f"{{{WNS}}}lineRule", "")
            ok   = (line == DUT["body"]["line_twips"] and rule == "auto")
            self._add(ok, "正文样式", "行距",
                      f"1.25倍 (300 twips, auto)",
                      f"line={line} rule={rule}  ≈{round(line/240,2)}倍")
        else:
            self._add(False, "正文样式", "行距", "1.25倍(300)", "未设置", "warn")

    # ── 3. 标题样式 ────────────────────────────
    def check_heading_styles(self):
        if self.styles is None:
            return

        for level in (1, 2, 3):
            std = DUT[f"heading{level}"]
            style = None
            for s in self.styles.iter(f"{{{WNS}}}style"):
                if s.get(f"{{{WNS}}}styleId") == str(level):
                    style = s
                    break
            if style is None:
                self._add(False, "标题样式", f"标题{level}", "可找到", "未找到", "warn")
                continue

            # 字体
            fonts = style.find(f".//{{{WNS}}}rFonts")
            if fonts is not None:
                ea = fonts.get(f"{{{WNS}}}eastAsia", "")
                ok = ea == std["font"]
                self._add(ok, "标题样式", f"标题{level}中文字体",
                          std["font"], ea or "未设置")
            else:
                self._add(False, "标题样式", f"标题{level}字体", std["font"], "未设置", "warn")

            # 字号
            sz = style.find(f".//{{{WNS}}}sz")
            if sz is not None:
                val = int(sz.get(f"{{{WNS}}}val", 0))
                # 对标题1允许30或32（模板实测为30）
                if level == 1:
                    ok = val in (30, 32)
                    self._add(ok, "标题样式", f"标题{level}字号",
                              "三号(15-16pt, 30-32 half-pt)",
                              f"{val} half-pt ({val//2}pt)")
                else:
                    ok = val == std["size_halfpt"]
                    self._add(ok, "标题样式", f"标题{level}字号",
                              f"{std['size_halfpt']} half-pt ({std['size_halfpt']//2}pt)",
                              f"{val} half-pt ({val//2}pt)")

            # 行距
            sp = style.find(f".//{{{WNS}}}spacing")
            if sp is not None:
                line = int(sp.get(f"{{{WNS}}}line", 0))
                ok   = line >= 360   # 1.5倍及以上均可
                self._add(ok, "标题样式", f"标题{level}行距",
                          f"≥1.5倍(360 twips)",
                          f"{line} twips ≈{round(line/240,2)}倍",
                          "info" if ok else "warn")

    # ── 4. 字数统计 ────────────────────────────
    def check_word_count(self):
        full = self._full_text()
        cn   = len(re.findall(r"[一-鿿]", full))
        en   = len(re.findall(r"[a-zA-Z]+", full))
        total = cn + en
        ok    = total >= DUT["word_count_min"]
        self._add(ok, "字数", "总字数",
                  f"≥{DUT['word_count_min']}字",
                  f"约{total}字（中文{cn}字 + 英文{en}词）",
                  "error" if not ok else "info")

    # ── 5. 必要结构章节 ────────────────────────
    def check_structure(self):
        full = self._full_text()
        sections = {
            "摘要（中文）":   r"摘\s*要",
            "Abstract（英文）": r"Abstract|ABSTRACT",
            "关键词（中文）": r"关键词",
            "Keywords（英文）": r"Key\s*[Ww]ords?|KEYWORDS",
            "目录":           r"目\s*录",
            "绪论/引言":      r"绪论|引言",
            "相关工作/文献综述": r"相关工作|文献综述|Related",
            "方法章节":       r"方法|Method|框架|Framework",
            "实验与评估":     r"实验|评估|Experiment",
            "总结/结论":      r"总结|结论|Conclusion",
            "参考文献":       r"参考文献|References",
            "修改记录":       r"修改记录",
            "致谢":           r"致\s*谢|Acknowledgment",
        }
        for name, pattern in sections.items():
            found = bool(re.search(pattern, full))
            self._add(found, "论文结构", name,
                      "存在", "✓ 已找到" if found else "✗ 未找到",
                      "info" if found else "error")

    # ── 6. 图表编号 ────────────────────────────
    def check_figure_table_numbering(self):
        full = self._full_text()
        # 官方模板使用 X.X 格式（如图2.1），不是 X-X
        correct = re.compile(r"[图表]\s*\d+[.．]\d+")
        wrong_fig = re.compile(r"图\s*(\d+)(?![.．\d])")   # 图n 没有章节号
        wrong_tbl = re.compile(r"表\s*(\d+)(?![.．\d])")

        # 收集所有"图X.X"格式是否正确
        all_figs = re.findall(r"图\s*[\d]+[.．]?[\d]*", full)
        bad_figs = [f for f in all_figs if not re.search(r"\d+[.．]\d+", f)]
        all_tbls = re.findall(r"表\s*[\d]+[.．]?[\d]*", full)
        bad_tbls = [t for t in all_tbls if not re.search(r"\d+[.．]\d+", t)]

        # 去重
        bad_figs = list(dict.fromkeys(bad_figs))[:8]
        bad_tbls = list(dict.fromkeys(bad_tbls))[:8]

        ok_fig = len(bad_figs) == 0
        ok_tbl = len(bad_tbls) == 0

        self._add(ok_fig, "图表编号", "图编号格式",
                  "图X.X（如图3.1）",
                  "全部正确" if ok_fig else f"不合规示例：{'、'.join(bad_figs)}")
        self._add(ok_tbl, "图表编号", "表编号格式",
                  "表X.X（如表3.1）",
                  "全部正确" if ok_tbl else f"不合规示例：{'、'.join(bad_tbls)}")

        # 图在正文中是否有引用（简单检查：有 如图X.X 或 图X.X所示）
        fig_refs = re.findall(r"[如（(]?\s*图\s*\d+[.．]\d+", full)
        tbl_refs = re.findall(r"[如（(]?\s*表\s*\d+[.．]\d+", full)
        self._add(len(fig_refs) > 0, "图表编号", "图在正文中有引用",
                  "正文中出现 如图X.X 等引用",
                  f"找到{len(fig_refs)}处引用" if fig_refs else "未找到任何图引用",
                  "warn" if len(fig_refs) == 0 else "info")
        self._add(len(tbl_refs) > 0, "图表编号", "表在正文中有引用",
                  "正文中出现 如表X-X 等引用",
                  f"找到{len(tbl_refs)}处引用" if tbl_refs else "未找到任何表引用",
                  "warn" if len(tbl_refs) == 0 else "info")

    # ── 7. 章节编号格式 ────────────────────────
    def check_chapter_numbering(self):
        texts = self._texts()
        # 检测一级标题：应为 "N  标题" 或 "第N章 标题"
        h1_num = re.compile(r"^(\d+)\s{1,4}\S|^第[一二三四五六七八九十\d]+章\s")
        # 二级标题：N.N  标题
        h2_num = re.compile(r"^(\d+)\.(\d+)\s")
        # 三级标题：N.N.N  标题
        h3_num = re.compile(r"^(\d+)\.(\d+)\.(\d+)\s")

        h1_found, h2_found, h3_found = [], [], []
        for t in texts:
            t = t.strip()
            if h1_num.match(t) and len(t) < 30:
                h1_found.append(t[:25])
            elif h3_num.match(t) and len(t) < 40:
                h3_found.append(t[:30])
            elif h2_num.match(t) and len(t) < 40:
                h2_found.append(t[:30])

        self._add(len(h1_found) >= 3, "章节编号", "一级标题（章）",
                  "≥3章，格式：N  标题名",
                  f"找到{len(h1_found)}个：{h1_found[:3]}")
        self._add(len(h2_found) >= 5, "章节编号", "二级标题",
                  "≥5个，格式：N.N  标题名",
                  f"找到{len(h2_found)}个",
                  "info" if len(h2_found) >= 5 else "warn")
        self._add(len(h3_found) >= 3, "章节编号", "三级标题",
                  "≥3个，格式：N.N.N  标题名",
                  f"找到{len(h3_found)}个",
                  "info" if len(h3_found) >= 3 else "warn")

    # ── 8. 参考文献 ────────────────────────────
    def check_references(self):
        texts = self._texts()
        full  = self._full_text()

        # 定位参考文献节
        ref_idx = -1
        for i, t in enumerate(texts):
            if t.strip() in ("参考文献", "参 考 文 献", "References", "REFERENCES"):
                ref_idx = i
                break

        if ref_idx < 0:
            self._add(False, "参考文献", "参考文献章节", "存在", "未找到")
            return
        self._add(True, "参考文献", "参考文献章节", "存在", "✓ 已找到", "info")

        # 参考文献条目（含段落对象，用于检测 Word 自动编号）
        paras_after = self._paras[ref_idx+1:]
        stop_kw = re.compile(r"致\s*谢|修改记录|Acknowledgment")
        entry_paras = []
        for p in paras_after:
            if p["text"].strip() and stop_kw.search(p["text"]):
                break
            if p["text"].strip():
                entry_paras.append(p)

        entries = [p["text"].strip() for p in entry_paras]

        # 判断是否使用 Word 自动列表编号（numPr）
        auto_num_ids = {p["num_id"] for p in entry_paras if p.get("num_id")}
        using_auto_num = len(auto_num_ids) > 0

        # [n] 格式：显式文本匹配 或 Word 自动编号段落
        numbered_explicit = [e for e in entries if re.match(r"^\[\d+\]", e)]
        auto_num_entries  = [p["text"].strip() for p in entry_paras if p.get("num_id")]
        # 合并（显式优先）
        if numbered_explicit:
            numbered = numbered_explicit
            non_num  = [e for e in entries if e and not re.match(r"^\[\d+\]", e)]
            num_note = f"共{len(numbered)}条（显式[n]文本）"
        elif auto_num_entries:
            numbered = auto_num_entries
            non_num  = []
            num_note = f"共{len(numbered)}条（Word自动列表编号，[n]在numPr中，提取文本不含前缀）"
        else:
            numbered = []
            non_num  = entries
            num_note = "0条"

        self._add(len(numbered) >= DUT["ref_min"], "参考文献", "文献数量",
                  f"≥{DUT['ref_min']}条",
                  num_note,
                  "error" if len(numbered) < DUT["ref_min"] else "info")

        fmt_note = "格式统一" if not non_num else f"有{len(non_num)}条未用[n]格式：{non_num[:3]}"
        if using_auto_num and not numbered_explicit:
            fmt_note += "（Word自动编号，视觉正确，建议人工确认已显示[n]）"
        self._add(len(non_num) == 0, "参考文献", "编号格式统一",
                  "全部使用[n]格式",
                  fmt_note,
                  "warn" if (non_num and not using_auto_num) else "info")

        # 中英文分布
        cn_cnt = sum(1 for e in numbered if re.search(r"[一-鿿]", e))
        en_cnt = len(numbered) - cn_cnt
        self._add(True, "参考文献", "中英文分布",
                  "中英文文献均有",
                  f"中文{cn_cnt}条，英文{en_cnt}条",
                  "info")

        # GB/T 7714 典型格式检查（抽查：期刊[J]、会议[C]、EB/OL等标注）
        type_marks = re.findall(r"\[J\]|\[C\]|\[EB/OL\]|\[M\]|\[D\]|\[R\]", full)
        self._add(len(type_marks) >= 5, "参考文献", "文献类型标注[J][C]等",
                  "≥5条含[J][C][EB/OL]等类型标注 (GB/T 7714)",
                  f"找到{len(type_marks)}处类型标注",
                  "warn" if len(type_marks) < 5 else "info")

        # DOI / URL 格式（arXiv / https）
        url_refs = [e for e in numbered if re.search(r"https?://|arXiv|doi", e, re.I)]
        self._add(len(url_refs) > 0, "参考文献", "网络资源含URL/DOI",
                  "网络文献应附 URL 或 DOI",
                  f"{len(url_refs)}条含URL/DOI",
                  "info")

        # 顺序连续性（自动编号模式下无法从文本提取，跳过）
        nums = []
        if not (using_auto_num and not numbered_explicit):
            for e in numbered:
                m = re.match(r"^\[(\d+)\]", e)
                if m:
                    nums.append(int(m.group(1)))
        gaps = [nums[i] for i in range(1, len(nums)) if nums[i] != nums[i-1]+1]
        if using_auto_num and not numbered_explicit:
            cont_note = "Word自动编号，顺序由Word维护，人工确认"
        else:
            cont_note = "连续" if not gaps else f"跳号位置：{gaps[:5]}"
        self._add(len(gaps) == 0, "参考文献", "编号连续性",
                  "编号连续无跳号",
                  cont_note,
                  "warn" if gaps else "info")

    # ── 9. 摘要格式 ────────────────────────────
    def check_abstract(self):
        full = self._full_text()
        # 中文摘要
        has_cn = bool(re.search(r"摘\s*要", full))
        has_en = bool(re.search(r"Abstract|ABSTRACT", full))
        has_kw_cn = bool(re.search(r"关键词", full))
        has_kw_en = bool(re.search(r"Key\s*[Ww]ords?|KEYWORDS", full))

        for name, found in [("中文摘要", has_cn), ("英文摘要", has_en),
                             ("中文关键词", has_kw_cn), ("英文关键词", has_kw_en)]:
            self._add(found, "摘要", name,
                      "存在", "✓" if found else "✗ 缺失",
                      "info" if found else "error")

        # 关键词数量（建议3-8个）
        kw_match = re.search(r"关键词[：:]\s*(.+?)(?:\n|Abstract|$)", full, re.S)
        if kw_match:
            kw_text = kw_match.group(1).strip()[:200]
            kw_cnt  = len(re.split(r"[；;，,、\s]+", kw_text))
            ok = 3 <= kw_cnt <= 8
            self._add(ok, "摘要", "关键词数量",
                      "3-8个", f"{kw_cnt}个：{kw_text[:60]}",
                      "info" if ok else "warn")

    # ── 10. 页眉页脚 ───────────────────────────
    def check_header_footer(self):
        with zipfile.ZipFile(self.path) as z:
            headers = [n for n in z.namelist() if n.startswith("word/header")]
            footers = [n for n in z.namelist() if n.startswith("word/footer")]

        self._add(len(headers) > 0, "页眉页脚", "页眉",
                  "正文部分有页眉", f"找到{len(headers)}个页眉文件",
                  "info" if headers else "warn")
        self._add(len(footers) > 0, "页眉页脚", "页脚（页码）",
                  "正文部分有页脚/页码", f"找到{len(footers)}个页脚文件",
                  "info" if footers else "warn")

    # ── 11. 封面题目字数 ────────────────────────
    def check_title_length(self):
        texts = self._texts()
        # 封面题目通常是前几行中最长的中文行
        candidates = []
        for t in texts[:15]:
            t = t.strip()
            cn_len = len(re.findall(r"[一-鿿]", t))
            if 5 <= cn_len <= 40:
                candidates.append((cn_len, t))
        if candidates:
            ln, title = max(candidates)
            ok = ln <= 25   # 大连理工要求题目≤25字（含标点更宽松）
            self._add(ok, "封面", "论文题目长度",
                      "≤25个汉字",
                      f"{ln}字：{title[:40]}",
                      "warn" if not ok else "info")

    # ── 运行所有检查 ───────────────────────────
    def run(self) -> List[Item]:
        self.check_page()
        self.check_default_style()
        self.check_heading_styles()
        self.check_word_count()
        self.check_title_length()
        self.check_abstract()
        self.check_structure()
        self.check_chapter_numbering()
        self.check_figure_table_numbering()
        self.check_header_footer()
        self.check_references()
        return self.results

    # ── 人工核查清单（文本 / HTML 共用） ──────────
    MANUAL_CHECKS = [
        "正文字体：中文宋体 + 英文 Times New Roman，小四（12pt）",
        "行距：全文统一多倍行距 1.25（含表格内文字）",
        "首行缩进：每段首行缩进 2 字符（不可用空格代替）",
        "一级标题居左（黑体小三，非居中）",
        "图/表题注位置：图题在图下，表题在表上",
        "图/表题注字体：五号宋体，居中",
        "页眉：论文中文题目，宋体五号居中（封面/声明页无页眉）",
        "页码：正文从第1页起阿拉伯数字居中；摘要/目录用罗马数字",
        "打印：封面/声明/摘要单面；目录/正文/致谢双面",
        "原创性声明、使用授权声明：手写签字，日期已填",
        "修改记录：四项内容已填写，查重重复比已填入",
        "三线表：无竖线，只有顶线、栏目线、底线",
        "公式编号：(章号.序号) 居中+右对齐",
        "参考文献若使用 Word 自动列表编号：目视确认显示为 [1][2]… 格式",
        "外文翻译（原文+译文）已准备",
    ]

    # ── 文本报告 ───────────────────────────────
    def report(self) -> str:
        errors = [r for r in self.results if not r.ok and r.level == "error"]
        warns  = [r for r in self.results if not r.ok and r.level == "warn"]
        passed = [r for r in self.results if r.ok]

        lines = [
            "=" * 65,
            "  大连理工大学本科毕业论文  格式审查报告（自动版）",
            f"  文件：{self.path.name}",
            "=" * 65,
            f"\n  检查项：{len(self.results)}项  |  通过：{len(passed)}  |  "
            f"错误：{len(errors)}  |  警告：{len(warns)}",
            "",
        ]

        if errors:
            lines += ["─" * 65, "【✗ 必须修改】", "─" * 65]
            for r in errors:
                lines += [
                    f"  [{r.category}] {r.name}",
                    f"    要求：{r.expected}",
                    f"    实际：{r.actual}",
                    "",
                ]

        if warns:
            lines += ["─" * 65, "【△ 建议修改 / 请人工核对】", "─" * 65]
            for r in warns:
                lines += [
                    f"  [{r.category}] {r.name}",
                    f"    要求：{r.expected}",
                    f"    实际：{r.actual}",
                    "",
                ]

        lines += ["─" * 65, "【✓ 已通过】", "─" * 65]
        for r in passed:
            lines.append(f"  ✓ [{r.category}] {r.name}：{r.actual}")

        lines += [
            "",
            "=" * 65,
            "【需人工核对的项目（程序无法自动检查）】",
            "─" * 65,
        ]
        for item in self.MANUAL_CHECKS:
            lines.append(f"  □ {item}")
        lines.append("=" * 65)
        return "\n".join(lines)

    # ── HTML 报告 ──────────────────────────────
    def report_html(self) -> str:
        import html as _html
        from datetime import datetime

        errors = [r for r in self.results if not r.ok and r.level == "error"]
        warns  = [r for r in self.results if not r.ok and r.level == "warn"]
        passed = [r for r in self.results if r.ok]
        now    = datetime.now().strftime("%Y-%m-%d %H:%M")

        def esc(s): return _html.escape(str(s))

        def card(r, color):
            border = {"red": "#fca5a5", "amber": "#fcd34d", "green": "#86efac"}[color]
            bg     = {"red": "#fff5f5", "amber": "#fffbeb", "green": "#f0fdf4"}[color]
            icon   = {"red": "✗", "amber": "△", "green": "✓"}[color]
            ic     = {"red": "#c53030", "amber": "#b7791f", "green": "#276749"}[color]
            return f"""
            <div style="background:{bg};border:1px solid {border};border-radius:8px;
                        padding:1rem 1.2rem;margin-bottom:.7rem;">
              <div style="display:flex;align-items:baseline;gap:.6rem;margin-bottom:.35rem;">
                <span style="font-weight:700;color:{ic};font-size:1rem;">{icon}</span>
                <span style="font-size:.8rem;color:#6b7280;font-weight:500;">
                  [{esc(r.category)}]</span>
                <span style="font-weight:600;font-size:.92rem;">{esc(r.name)}</span>
              </div>
              <div style="font-size:.83rem;color:#374151;padding-left:1.4rem;">
                <span style="color:#6b7280;">要求：</span>{esc(r.expected)}<br>
                <span style="color:#6b7280;">实际：</span><strong>{esc(r.actual)}</strong>
              </div>
            </div>"""

        errors_html = "".join(card(r, "red")   for r in errors)
        warns_html  = "".join(card(r, "amber") for r in warns)
        passed_html = "".join(card(r, "green") for r in passed)

        manual_items = "".join(
            f'<li><label style="cursor:pointer;">'
            f'<input type="checkbox" style="margin-right:.5rem;">{esc(item)}'
            f'</label></li>\n'
            for item in self.MANUAL_CHECKS
        )

        n_all = len(self.results)
        n_ok  = len(passed)
        n_err = len(errors)
        n_wrn = len(warns)

        status_color = "#c53030" if n_err else ("#b7791f" if n_wrn else "#276749")
        status_text  = (f"{n_err} 项必须修改" if n_err else
                        (f"{n_wrn} 项建议核对" if n_wrn else "全部通过"))

        return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>格式审查报告 · {esc(self.path.name)}</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC",
       "Microsoft YaHei",sans-serif;background:#f5f7fa;color:#1a2433;line-height:1.7}}
  .header{{background:linear-gradient(135deg,#00335e,#004f8c);color:#fff;
           padding:2rem 2rem 1.5rem}}
  .header h1{{font-size:1.25rem;font-weight:700;margin-bottom:.3rem}}
  .header .meta{{font-size:.82rem;opacity:.75}}
  .status-banner{{background:#fff;border-bottom:3px solid {status_color};
                  padding:.75rem 2rem;font-size:.9rem;font-weight:600;
                  color:{status_color};display:flex;align-items:center;gap:.5rem}}
  .metrics{{display:flex;flex-wrap:wrap;background:#fff;
            border-bottom:1px solid #e8ecf1;padding:.5rem 1rem}}
  .metric{{flex:1;min-width:100px;text-align:center;padding:.8rem .5rem}}
  .metric-num{{font-size:1.8rem;font-weight:700;line-height:1}}
  .metric-label{{font-size:.73rem;color:#9aa5b4;margin-top:.2rem}}
  .body{{max-width:800px;margin:0 auto;padding:1.5rem}}
  .section{{margin-bottom:1.8rem}}
  .section-head{{font-size:1rem;font-weight:700;margin-bottom:.9rem;
                 display:flex;align-items:center;gap:.5rem}}
  .section-head .pill{{font-size:.72rem;font-weight:600;padding:.15rem .55rem;
                       border-radius:20px}}
  .pill-red   {{background:#fff0f0;color:#c53030;border:1px solid #fca5a5}}
  .pill-amber {{background:#fffbeb;color:#b7791f;border:1px solid #fcd34d}}
  .pill-green {{background:#f0fdf4;color:#276749;border:1px solid #86efac}}
  details{{background:#fff;border:1px solid #e8ecf1;border-radius:8px;overflow:hidden}}
  details>summary{{padding:.85rem 1.1rem;cursor:pointer;font-weight:600;
                   font-size:.92rem;list-style:none;display:flex;
                   align-items:center;justify-content:space-between}}
  details>summary::after{{content:"▸";transition:transform .2s}}
  details[open]>summary::after{{transform:rotate(90deg)}}
  details>div{{padding:.2rem 1.1rem 1rem}}
  .manual-list{{list-style:none;display:flex;flex-direction:column;gap:.5rem;
                padding:.2rem 1.1rem 1rem}}
  .manual-list li{{font-size:.87rem;color:#374151;
                   border-bottom:1px solid #f0f0f0;padding-bottom:.45rem}}
  .manual-list li:last-child{{border-bottom:none}}
  .footer{{text-align:center;font-size:.78rem;color:#9aa5b4;
           padding:1.5rem;border-top:1px solid #e8ecf1;background:#fff;
           margin-top:2rem}}
  @media(max-width:500px){{.metric-num{{font-size:1.3rem}}}}
</style>
</head>
<body>

<div class="header">
  <h1>大连理工大学本科毕业论文 · 格式审查报告</h1>
  <div class="meta">文件：{esc(self.path.name)} &nbsp;·&nbsp; 审查时间：{now}</div>
</div>

<div class="status-banner">
  {'⚠' if n_err or n_wrn else '✓'} &nbsp;{status_text}
</div>

<div class="metrics">
  <div class="metric">
    <div class="metric-num" style="color:#1a2433">{n_all}</div>
    <div class="metric-label">检查项</div>
  </div>
  <div class="metric">
    <div class="metric-num" style="color:#276749">{n_ok}</div>
    <div class="metric-label">已通过</div>
  </div>
  <div class="metric">
    <div class="metric-num" style="color:#c53030">{n_err}</div>
    <div class="metric-label">必须修改</div>
  </div>
  <div class="metric">
    <div class="metric-num" style="color:#b7791f">{n_wrn}</div>
    <div class="metric-label">建议核对</div>
  </div>
</div>

<div class="body">

  {f'''<div class="section">
    <div class="section-head">
      <span class="pill pill-red">✗ 必须修改</span>
    </div>
    {errors_html}
  </div>''' if errors else ''}

  {f'''<div class="section">
    <div class="section-head">
      <span class="pill pill-amber">△ 建议核对</span>
    </div>
    {warns_html}
  </div>''' if warns else ''}

  <div class="section">
    <details {"open" if not errors and not warns else ""}>
      <summary>
        <span>✓ 已通过 &nbsp;<span class="pill pill-green"
          style="font-size:.75rem">{n_ok} 项</span></span>
      </summary>
      <div>{passed_html}</div>
    </details>
  </div>

  <div class="section">
    <details>
      <summary>□ 需人工核对（打印前逐项确认）</summary>
      <ul class="manual-list">
        {manual_items}
      </ul>
    </details>
  </div>

</div>

<div class="footer">
  仅供辅助审查，最终以学院审核意见为准 &nbsp;·&nbsp;
  <a href="https://github.com/jackeyloveseven/dut-thesis-format-checker"
     style="color:#9aa5b4">dut-thesis-format-checker</a>
</div>

</body>
</html>"""


# ──────────────────────────────────────────────
# 入口
# ──────────────────────────────────────────────
def main():
    args = sys.argv[1:]
    gen_html = "--html" in args
    args = [a for a in args if a != "--html"]

    if not args:
        docx_files = list(Path(".").glob("*.docx"))
        if not docx_files:
            print("用法：python dut_format_checker.py <论文.docx> [--html]")
            sys.exit(1)
        docx_path = str(docx_files[0])
        print(f"自动选择：{docx_path}\n")
    else:
        docx_path = args[0]

    checker = Checker(docx_path)
    checker.run()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    print(checker.report())

    if gen_html:
        out = Path(docx_path).with_suffix("") .parent / (Path(docx_path).stem + "_格式报告.html")
        out.write_text(checker.report_html(), encoding="utf-8")
        print(f"\n✓ HTML 报告已生成：{out}")


if __name__ == "__main__":
    main()
