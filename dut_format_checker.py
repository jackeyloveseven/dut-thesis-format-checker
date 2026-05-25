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
        n_all  = len(self.results)
        n_ok   = len(passed)
        n_err  = len(errors)
        n_wrn  = len(warns)
        pct    = int(n_ok / n_all * 100) if n_all else 0

        def esc(s): return _html.escape(str(s))

        def issue_card(r, variant):
            cfg = {
                "error": ("#ef4444", "#fef2f2", "#fecaca", "✕", "必须修改"),
                "warn":  ("#f59e0b", "#fffbeb", "#fde68a", "！", "建议核对"),
            }[variant]
            accent, bg, border, icon, tag = cfg
            return f"""<div class="icard" style="--accent:{accent};--bg:{bg};--border:{border}">
  <div class="icard-left"><span class="icard-icon">{icon}</span></div>
  <div class="icard-body">
    <div class="icard-head">
      <span class="icard-tag" style="color:{accent}">{tag}</span>
      <span class="icard-cat">{esc(r.category)}</span>
      <span class="icard-name">{esc(r.name)}</span>
    </div>
    <div class="icard-row"><span class="lbl">要求</span>{esc(r.expected)}</div>
    <div class="icard-row"><span class="lbl">实际</span><strong>{esc(r.actual)}</strong></div>
  </div>
</div>"""

        def pass_row(r):
            return (f'<div class="prow"><span class="pcheck">✓</span>'
                    f'<span class="pcat">{esc(r.category)}</span>'
                    f'<span class="pname">{esc(r.name)}</span>'
                    f'<span class="pval">{esc(r.actual)}</span></div>')

        errors_html = "\n".join(issue_card(r, "error") for r in errors)
        warns_html  = "\n".join(issue_card(r, "warn")  for r in warns)
        passed_html = "\n".join(pass_row(r)             for r in passed)

        manual_html = "\n".join(
            f'<label class="mitem"><input type="checkbox"><span>{esc(item)}</span></label>'
            for item in self.MANUAL_CHECKS
        )

        overall_color = "#ef4444" if n_err else ("#f59e0b" if n_wrn else "#10a37f")
        overall_label = f"{n_err} 项必须修改" if n_err else (f"{n_wrn} 项建议核对" if n_wrn else "格式检查全部通过")
        overall_icon  = "✕" if n_err else ("！" if n_wrn else "✓")

        errors_section = f"""
<div class="section">
  <div class="section-label" style="color:#ef4444">
    <span class="dot" style="background:#ef4444"></span>必须修改 · {n_err} 项
  </div>
  {errors_html}
</div>""" if errors else ""

        warns_section = f"""
<div class="section">
  <div class="section-label" style="color:#f59e0b">
    <span class="dot" style="background:#f59e0b"></span>建议核对 · {n_wrn} 项
  </div>
  {warns_html}
</div>""" if warns else ""

        passed_open = "open" if not errors and not warns else ""

        return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>格式审查报告 · {esc(self.path.name)}</title>
<style>
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

body {{
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC",
               "Microsoft YaHei", sans-serif;
  background: #212121;
  color: #ececec;
  min-height: 100vh;
  line-height: 1.6;
}}

/* ── top bar ── */
.topbar {{
  background: #171717;
  border-bottom: 1px solid #2d2d2d;
  padding: .85rem 1.5rem;
  display: flex; align-items: center; justify-content: space-between;
  position: sticky; top: 0; z-index: 50;
}}
.topbar-title {{
  font-size: .88rem; font-weight: 600; color: #ececec;
  display: flex; align-items: center; gap: .6rem;
}}
.topbar-title .logo {{
  width: 22px; height: 22px; background: #10a37f; border-radius: 5px;
  display: flex; align-items: center; justify-content: center;
  font-size: .75rem; font-weight: 800; color: #fff; flex-shrink: 0;
}}
.topbar-meta {{ font-size: .75rem; color: #8e8ea0; }}

/* ── hero ── */
.hero {{
  max-width: 720px; margin: 3rem auto 0; padding: 0 1.25rem;
  text-align: center;
}}
.overall-badge {{
  display: inline-flex; align-items: center; gap: .5rem;
  background: #2a2a2a; border: 1px solid #3d3d3d;
  border-radius: 100px; padding: .35rem .9rem;
  font-size: .8rem; color: #8e8ea0; margin-bottom: 1rem;
}}
.hero-status {{
  font-size: 2.2rem; font-weight: 700;
  color: {overall_color}; line-height: 1.2; margin-bottom: .4rem;
}}
.hero-sub {{ font-size: .9rem; color: #8e8ea0; margin-bottom: 2rem; }}

/* ── progress ring ── */
.ring-wrap {{
  display: flex; justify-content: center; margin-bottom: 2.5rem;
}}
.ring-svg {{ transform: rotate(-90deg); }}
.ring-track {{ fill: none; stroke: #2d2d2d; stroke-width: 8; }}
.ring-fill  {{ fill: none; stroke: {overall_color}; stroke-width: 8;
               stroke-linecap: round;
               stroke-dasharray: 251.2;
               stroke-dashoffset: {251.2 * (1 - pct/100):.1f};
               transition: stroke-dashoffset .8s ease; }}
.ring-text  {{ font-size: .95rem; font-weight: 700; fill: #ececec; }}
.ring-sub   {{ font-size: .45rem; fill: #8e8ea0; }}

/* ── metric cards ── */
.metrics {{
  display: grid; grid-template-columns: repeat(4, 1fr); gap: .75rem;
  max-width: 720px; margin: 0 auto 2.5rem; padding: 0 1.25rem;
}}
.metric {{
  background: #2a2a2a; border: 1px solid #3d3d3d; border-radius: 12px;
  padding: 1rem .75rem; text-align: center;
}}
.metric-num {{ font-size: 1.9rem; font-weight: 700; line-height: 1; }}
.metric-label {{ font-size: .72rem; color: #8e8ea0; margin-top: .3rem; }}

/* ── main body ── */
.body {{ max-width: 720px; margin: 0 auto; padding: 0 1.25rem 4rem; }}

.section {{ margin-bottom: 1.5rem; }}
.section-label {{
  font-size: .78rem; font-weight: 600; letter-spacing: .04em;
  text-transform: uppercase; margin-bottom: .75rem;
  display: flex; align-items: center; gap: .5rem;
}}
.dot {{ width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; }}

/* ── issue card ── */
.icard {{
  background: var(--bg); border: 1px solid var(--border);
  border-left: 3px solid var(--accent);
  border-radius: 10px; padding: 1rem 1.1rem;
  display: flex; gap: .85rem; margin-bottom: .6rem;
}}
.icard-icon {{
  width: 24px; height: 24px; border-radius: 50%;
  background: var(--accent); color: #fff;
  font-size: .7rem; font-weight: 800;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0; margin-top: .1rem;
}}
.icard-body {{ flex: 1; min-width: 0; }}
.icard-head {{ display: flex; flex-wrap: wrap; align-items: center; gap: .4rem; margin-bottom: .4rem; }}
.icard-tag {{ font-size: .72rem; font-weight: 700; }}
.icard-cat {{
  font-size: .72rem; background: #2d2d2d; color: #8e8ea0;
  padding: .1rem .45rem; border-radius: 4px;
}}
.icard-name {{ font-size: .88rem; font-weight: 600; color: #ececec; }}
.icard-row {{ font-size: .82rem; color: #a8a8b3; line-height: 1.5; }}
.icard-row .lbl {{
  display: inline-block; width: 2.5rem; color: #5c5c6e;
  font-size: .75rem; flex-shrink: 0;
}}
.icard-row strong {{ color: #ececec; }}

/* ── passed rows ── */
.pass-block {{
  background: #2a2a2a; border: 1px solid #3d3d3d;
  border-radius: 10px; overflow: hidden;
}}
.prow {{
  display: grid;
  grid-template-columns: 1.2rem 5rem 1fr auto;
  gap: .6rem; align-items: center;
  padding: .6rem 1rem; border-bottom: 1px solid #333;
  font-size: .82rem;
}}
.prow:last-child {{ border-bottom: none; }}
.pcheck {{ color: #10a37f; font-weight: 700; font-size: .85rem; }}
.pcat {{
  background: #333; color: #8e8ea0; font-size: .7rem;
  padding: .1rem .4rem; border-radius: 4px; white-space: nowrap;
  overflow: hidden; text-overflow: ellipsis;
}}
.pname {{ color: #c8c8d0; }}
.pval {{ color: #8e8ea0; font-size: .78rem; text-align: right; max-width: 180px;
         overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}

/* ── collapsible ── */
.collapse {{ background: #2a2a2a; border: 1px solid #3d3d3d; border-radius: 10px; overflow: hidden; }}
.collapse-toggle {{
  width: 100%; background: none; border: none; cursor: pointer;
  padding: .85rem 1rem; display: flex; align-items: center;
  justify-content: space-between; color: #ececec;
  font-size: .88rem; font-weight: 600; font-family: inherit;
}}
.collapse-toggle:hover {{ background: #303030; }}
.collapse-toggle .chevron {{
  transition: transform .2s; color: #8e8ea0; font-size: .75rem;
}}
.collapse-toggle.open .chevron {{ transform: rotate(90deg); }}
.collapse-body {{ display: none; }}
.collapse-body.open {{ display: block; }}

/* ── manual checklist ── */
.mlist {{
  display: flex; flex-direction: column; padding: .25rem 0;
}}
.mitem {{
  display: flex; align-items: flex-start; gap: .75rem;
  padding: .65rem 1rem; cursor: pointer;
  border-bottom: 1px solid #333; font-size: .85rem; color: #c8c8d0;
  transition: background .15s;
}}
.mitem:last-child {{ border-bottom: none; }}
.mitem:hover {{ background: #303030; }}
.mitem input[type=checkbox] {{
  flex-shrink: 0; width: 16px; height: 16px;
  accent-color: #10a37f; cursor: pointer; margin-top: .2rem;
}}
.mitem input:checked + span {{ text-decoration: line-through; color: #5c5c6e; }}

/* ── footer ── */
.footer {{
  text-align: center; font-size: .75rem; color: #5c5c6e; padding: 2rem 1rem;
  border-top: 1px solid #2d2d2d; margin-top: 1.5rem;
}}
.footer a {{ color: #5c5c6e; text-decoration: none; }}
.footer a:hover {{ color: #8e8ea0; }}

@media (max-width: 500px) {{
  .metrics {{ grid-template-columns: repeat(2, 1fr); }}
  .prow {{ grid-template-columns: 1.2rem 1fr; }}
  .pcat, .pval {{ display: none; }}
}}
</style>
</head>
<body>

<div class="topbar">
  <div class="topbar-title">
    <div class="logo">D</div>
    大连理工大学毕业论文格式审查
  </div>
  <div class="topbar-meta">{esc(self.path.name)} · {now}</div>
</div>

<div class="hero">
  <div class="overall-badge">
    <span style="color:{overall_color}">{overall_icon}</span>
    {overall_label}
  </div>
  <div class="hero-status">{overall_label}</div>
  <div class="hero-sub">共 {n_all} 项自动检查 · {n_ok} 项通过 · 通过率 {pct}%</div>

  <div class="ring-wrap">
    <svg class="ring-svg" width="120" height="120" viewBox="0 0 90 90">
      <circle class="ring-track" cx="45" cy="45" r="40"/>
      <circle class="ring-fill" cx="45" cy="45" r="40"/>
      <g transform="rotate(90,45,45)">
        <text class="ring-text" x="45" y="47" text-anchor="middle" dominant-baseline="middle">{pct}%</text>
        <text class="ring-sub" x="45" y="57" text-anchor="middle">通过率</text>
      </g>
    </svg>
  </div>
</div>

<div class="metrics">
  <div class="metric">
    <div class="metric-num" style="color:#ececec">{n_all}</div>
    <div class="metric-label">检查项</div>
  </div>
  <div class="metric">
    <div class="metric-num" style="color:#10a37f">{n_ok}</div>
    <div class="metric-label">已通过</div>
  </div>
  <div class="metric">
    <div class="metric-num" style="color:#ef4444">{n_err}</div>
    <div class="metric-label">必须修改</div>
  </div>
  <div class="metric">
    <div class="metric-num" style="color:#f59e0b">{n_wrn}</div>
    <div class="metric-label">建议核对</div>
  </div>
</div>

<div class="body">

  {errors_section}
  {warns_section}

  <div class="section">
    <div class="section-label" style="color:#10a37f">
      <span class="dot" style="background:#10a37f"></span>已通过 · {n_ok} 项
    </div>
    <div class="collapse">
      <button class="collapse-toggle {passed_open}" onclick="toggle(this)">
        <span>查看通过项</span>
        <span class="chevron">▶</span>
      </button>
      <div class="collapse-body {passed_open}">
        <div class="pass-block" style="border-radius:0;border:none">
          {passed_html}
        </div>
      </div>
    </div>
  </div>

  <div class="section">
    <div class="section-label" style="color:#8e8ea0">
      <span class="dot" style="background:#5c5c6e"></span>人工核对清单 · {len(self.MANUAL_CHECKS)} 项
    </div>
    <div class="collapse">
      <button class="collapse-toggle" onclick="toggle(this)">
        <span>打印前逐项勾选确认</span>
        <span class="chevron">▶</span>
      </button>
      <div class="collapse-body">
        <div class="mlist">
          {manual_html}
        </div>
      </div>
    </div>
  </div>

</div>

<div class="footer">
  仅供辅助审查，最终以学院审核意见为准 ·
  <a href="https://github.com/jackeyloveseven/dut-thesis-format-checker" target="_blank">
    dut-thesis-format-checker
  </a>
</div>

<script>
function toggle(btn) {{
  btn.classList.toggle('open');
  const body = btn.nextElementSibling;
  body.classList.toggle('open');
}}
</script>
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
