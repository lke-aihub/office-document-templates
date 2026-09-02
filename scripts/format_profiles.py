"""Deterministic Word style profiles for office-document-templates.

Values in ``official-gbt9704`` follow GB/T 9704-2012. Other profiles are
deliberate business defaults because leave notes, email archives, and internal
reports do not share a single mandatory national typesetting standard.
"""

from __future__ import annotations


FORMAT_SOURCES = {
    "gbt9704_status": "https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=F3CC9BEF482524C895FDA7A08BB4A70E",
    "gbt9704_text": "https://www.nanyang.gov.cn/2021/05-08/53205.html",
    "leave_note": "https://gk.jyc.edu.cn/__local/2/49/BB/D10FBDE35B18E92E41136C466A9_15945CB3_CBB27.pdf",
    "formal_email": "https://houston.dlmu.edu.cn/info/1129/2380.htm",
    "travel_report": "https://www.jyvtc.edu.cn/gjjy/544403/728056/index.html",
    "official_writing": "https://www.ndrc.gov.cn/fggz/gbzj/wsgz/201508/t20150812_1105092_ext.html",
}


PROFILES = {
    "official-gbt9704": {
        "page": {
            "width_mm": 210,
            "height_mm": 297,
            "top_mm": 37,
            "bottom_mm": 35,
            "left_mm": 28,
            "right_mm": 26,
            "header_mm": 15,
            "footer_mm": 17,
        },
        "body": {
            "font": "仿宋",
            "ascii_font": "Times New Roman",
            "size_pt": 16,
            "alignment": "justify",
            "line_spacing_pt": 29,
            "space_before_pt": 0,
            "space_after_pt": 0,
            "first_line_chars": 2,
        },
        "title": {
            "font": "方正小标宋简体",
            "ascii_font": "Times New Roman",
            "size_pt": 22,
            "bold": False,
            "alignment": "center",
            "space_before_pt": 0,
            "space_after_pt": 16,
            "line_spacing_pt": 29,
        },
        "recipient": {"font": "仿宋", "size_pt": 16, "bold": False},
        "headings": {
            1: {"font": "黑体", "size_pt": 16, "bold": False},
            2: {"font": "楷体", "size_pt": 16, "bold": False},
            3: {"font": "仿宋", "size_pt": 16, "bold": True},
            4: {"font": "仿宋", "size_pt": 16, "bold": False},
        },
        "signature": {"font": "仿宋", "size_pt": 16, "alignment": "right"},
        "metadata": {"font": "仿宋", "size_pt": 16, "color": "000000"},
        "table": {"font": "仿宋", "size_pt": 14, "header_fill": "E7E6E6"},
        "footer": {"font": "宋体", "size_pt": 14, "style": "official"},
    },
    "business-formal": {
        "page": {
            "width_mm": 210,
            "height_mm": 297,
            "top_mm": 25.4,
            "bottom_mm": 25.4,
            "left_mm": 28,
            "right_mm": 28,
            "header_mm": 12.7,
            "footer_mm": 15,
        },
        "body": {
            "font": "宋体",
            "ascii_font": "Calibri",
            "size_pt": 12,
            "alignment": "justify",
            "line_spacing": 1.5,
            "space_before_pt": 0,
            "space_after_pt": 0,
            "first_line_chars": 2,
        },
        "title": {
            "font": "微软雅黑",
            "ascii_font": "Calibri",
            "size_pt": 18,
            "bold": True,
            "alignment": "center",
            "space_before_pt": 0,
            "space_after_pt": 14,
        },
        "recipient": {"font": "宋体", "size_pt": 12, "bold": False},
        "headings": {
            1: {"font": "微软雅黑", "size_pt": 14, "bold": True},
            2: {"font": "黑体", "size_pt": 12, "bold": False},
            3: {"font": "楷体", "size_pt": 12, "bold": False},
            4: {"font": "宋体", "size_pt": 12, "bold": True},
        },
        "signature": {"font": "宋体", "size_pt": 12, "alignment": "right"},
        "metadata": {"font": "微软雅黑", "size_pt": 10.5, "color": "666666"},
        "table": {"font": "宋体", "size_pt": 10.5, "header_fill": "EDEDED"},
        "footer": {"font": "宋体", "size_pt": 9, "style": "center"},
    },
    "business-report": {
        "page": {
            "width_mm": 210,
            "height_mm": 297,
            "top_mm": 25.4,
            "bottom_mm": 25.4,
            "left_mm": 25.4,
            "right_mm": 25.4,
            "header_mm": 12.7,
            "footer_mm": 15,
        },
        "body": {
            "font": "宋体",
            "ascii_font": "Calibri",
            "size_pt": 12,
            "alignment": "justify",
            "line_spacing": 1.5,
            "space_before_pt": 0,
            "space_after_pt": 6,
            "first_line_chars": 2,
        },
        "title": {
            "font": "微软雅黑",
            "ascii_font": "Calibri",
            "size_pt": 20,
            "bold": True,
            "alignment": "center",
            "space_before_pt": 0,
            "space_after_pt": 16,
        },
        "recipient": {"font": "宋体", "size_pt": 12, "bold": False},
        "headings": {
            1: {"font": "微软雅黑", "size_pt": 14, "bold": True},
            2: {"font": "黑体", "size_pt": 12, "bold": False},
            3: {"font": "楷体", "size_pt": 12, "bold": False},
            4: {"font": "宋体", "size_pt": 12, "bold": True},
        },
        "signature": {"font": "宋体", "size_pt": 12, "alignment": "right"},
        "metadata": {"font": "微软雅黑", "size_pt": 10.5, "color": "666666"},
        "table": {"font": "宋体", "size_pt": 10.5, "header_fill": "D9EAF7"},
        "footer": {"font": "宋体", "size_pt": 9, "style": "center"},
    },
    "email-archive": {
        "page": {
            "width_mm": 210,
            "height_mm": 297,
            "top_mm": 22,
            "bottom_mm": 22,
            "left_mm": 25.4,
            "right_mm": 25.4,
            "header_mm": 12.7,
            "footer_mm": 12.7,
        },
        "body": {
            "font": "微软雅黑",
            "ascii_font": "Calibri",
            "size_pt": 11,
            "alignment": "left",
            "line_spacing": 1.35,
            "space_before_pt": 0,
            "space_after_pt": 8,
            "first_line_chars": 0,
        },
        "title": {
            "font": "微软雅黑",
            "ascii_font": "Calibri",
            "size_pt": 16,
            "bold": True,
            "alignment": "left",
            "space_before_pt": 0,
            "space_after_pt": 10,
        },
        "recipient": {"font": "微软雅黑", "size_pt": 11, "bold": False},
        "headings": {
            1: {"font": "微软雅黑", "size_pt": 13, "bold": True},
            2: {"font": "微软雅黑", "size_pt": 11, "bold": True},
            3: {"font": "微软雅黑", "size_pt": 11, "bold": True},
            4: {"font": "微软雅黑", "size_pt": 11, "bold": True},
        },
        "signature": {"font": "微软雅黑", "size_pt": 11, "alignment": "left"},
        "metadata": {"font": "微软雅黑", "size_pt": 10, "color": "666666"},
        "table": {"font": "微软雅黑", "size_pt": 10, "header_fill": "F2F2F2"},
        "footer": {"font": "微软雅黑", "size_pt": 8.5, "style": "none"},
    },
    "simple-note": {
        "page": {
            "width_mm": 210,
            "height_mm": 297,
            "top_mm": 30,
            "bottom_mm": 30,
            "left_mm": 30,
            "right_mm": 30,
            "header_mm": 12.7,
            "footer_mm": 12.7,
        },
        "body": {
            "font": "宋体",
            "ascii_font": "Calibri",
            "size_pt": 12,
            "alignment": "justify",
            "line_spacing": 1.5,
            "space_before_pt": 0,
            "space_after_pt": 0,
            "first_line_chars": 2,
        },
        "title": {
            "font": "微软雅黑",
            "ascii_font": "Calibri",
            "size_pt": 18,
            "bold": True,
            "alignment": "center",
            "space_before_pt": 0,
            "space_after_pt": 16,
        },
        "recipient": {"font": "宋体", "size_pt": 12, "bold": False},
        "headings": {
            1: {"font": "微软雅黑", "size_pt": 14, "bold": True},
            2: {"font": "黑体", "size_pt": 12, "bold": False},
            3: {"font": "楷体", "size_pt": 12, "bold": False},
            4: {"font": "宋体", "size_pt": 12, "bold": True},
        },
        "signature": {"font": "宋体", "size_pt": 12, "alignment": "right"},
        "metadata": {"font": "宋体", "size_pt": 10.5, "color": "666666"},
        "table": {"font": "宋体", "size_pt": 10.5, "header_fill": "F2F2F2"},
        "footer": {"font": "宋体", "size_pt": 9, "style": "none"},
    },
}


TYPE_DEFAULT_PROFILE = {
    "leave-note": "simple-note",
    "request": "business-formal",
    "notice": "business-formal",
    "application": "business-formal",
    "formal-email": "email-archive",
    "travel-report": "business-report",
    "weekly-report": "business-report",
    "monthly-report": "business-report",
    "work-report": "business-report",
    "official-document": "official-gbt9704",
}


def resolve_profile(document_type: str, requested: str, payload: dict) -> tuple[str, dict]:
    """Resolve the requested profile without silently treating all work text as official."""
    if requested != "auto":
        if requested not in PROFILES:
            raise ValueError(f"Unknown profile: {requested}")
        return requested, PROFILES[requested]
    if payload.get("official") is True:
        return "official-gbt9704", PROFILES["official-gbt9704"]
    profile_name = TYPE_DEFAULT_PROFILE[document_type]
    return profile_name, PROFILES[profile_name]
