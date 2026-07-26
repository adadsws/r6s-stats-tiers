"""Render searchable PDF companions for the five leaderboard workbooks."""

from io import BytesIO
from pathlib import Path
from typing import Iterable, Mapping, Tuple
from xml.sax.saxutils import escape

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A2
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.platypus import (
    Image,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from .sources import ReportSources
from .tier_chart import CARDS_PER_ROW, OperatorCard, SOURCE_SHEETS, operator_key
from .tiers import TIER_COLORS


PDF_FONT = "R6Chinese"
PAGE_SIZE = A2
PAGE_MARGIN = 10 * mm
CONTENT_WIDTH = PAGE_SIZE[0] - 2 * PAGE_MARGIN
MIN_PAGE_HEIGHT = 100 * mm


def _register_fonts() -> None:
    if PDF_FONT not in pdfmetrics.getRegisteredFontNames():
        candidates = (
            Path("C:/Windows/Fonts/msyh.ttc"),
            Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
        )
        font_path = next((path for path in candidates if path.is_file()), None)
        if font_path is not None:
            pdfmetrics.registerFont(
                TTFont(PDF_FONT, str(font_path), subfontIndex=0)
            )
        else:
            pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
            globals()["PDF_FONT"] = "STSong-Light"


def _style(size=9, *, bold=False, color="#202327", align=TA_LEFT):
    return ParagraphStyle(
        "pdf-%s-%s-%s" % (size, bold, align),
        fontName=PDF_FONT,
        fontSize=size,
        leading=size * 1.35,
        textColor=colors.HexColor(color),
        alignment=align,
    )


def _p(text, style):
    return Paragraph(escape(str(text)).replace("\n", "<br/>"), style)


def _rpm_text(values: Tuple[int, ...]) -> str:
    return "/".join(str(value) for value in values) or "-"


def _card_flowable(
    card: OperatorCard,
    badge_dir: Path,
    gadget_icons: Mapping[str, Path],
    marker: str,
):
    badge_path = badge_dir / (operator_key(card.name) + ".png")
    badge = Image(str(badge_path), width=15 * mm, height=15 * mm)
    name = _p(card.name, _style(10, bold=True))
    tier = _p(
        "%s%s" % (card.tier, marker),
        _style(
            9,
            bold=True,
            color=(
                "#FFFFFF"
                if card.tier in ("S", "A", "D", "F")
                else "#1B1D20"
            ),
            align=TA_CENTER,
        ),
    )
    info = Table(
        [
            [name, tier, _p("%d速" % card.speed, _style(9, align=TA_CENTER))],
            [
                _p(
                    "副喷 %s" % ("是" if card.has_secondary_shotgun else "-"),
                    _style(),
                ),
                _p("主狙 %s" % ("是" if card.has_semiautomatic else "-"), _style()),
                "",
            ],
            [
                _p("副 %s" % _rpm_text(card.secondary_rpms), _style()),
                _p("主 %s" % _rpm_text(card.primary_rpms), _style()),
                "",
            ],
        ],
        colWidths=[30 * mm, 19 * mm, 12 * mm],
        rowHeights=[7 * mm, 6 * mm, 6 * mm],
    )
    info.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#CDD1D5")),
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F4F5F7")),
                (
                    "BACKGROUND",
                    (1, 0),
                    (1, 0),
                    colors.HexColor("#" + TIER_COLORS[card.tier]),
                ),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("SPAN", (1, 1), (2, 1)),
                ("SPAN", (1, 2), (2, 2)),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ]
        )
    )
    gadget_row = []
    for gadget in card.gadgets:
        path = gadget_icons[gadget.name]
        gadget_row.append(
            Table(
                [
                    [
                        Image(str(path), width=5 * mm, height=5 * mm),
                        _p(
                            gadget.name
                            + (
                                "×%d" % gadget.quantity
                                if gadget.quantity is not None
                                else ""
                            ),
                            _style(6.5),
                        ),
                    ]
                ],
                colWidths=[5 * mm, 14 * mm],
            )
        )
    gadgets = Table(
        [gadget_row[:4], gadget_row[4:8]] if len(gadget_row) > 4 else [gadget_row],
        colWidths=[19 * mm] * 4,
    )
    card_table = Table(
        [[badge, info], [gadgets, ""]],
        colWidths=[15 * mm, 62 * mm],
        rowHeights=[20 * mm, 13 * mm],
    )
    card_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F4F5F7")),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CDD1D5")),
                ("SPAN", (0, 1), (1, 1)),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ]
        )
    )
    return card_table


def _source_lines(sources: ReportSources):
    rating = sources.rating
    wiki = sources.wiki
    return (
        "评分来源：%s《%s》｜%s｜覆盖至 %s｜%s"
        % (
            rating.creator,
            rating.title,
            rating.season,
            rating.covered_patch,
            rating.url,
        ),
        "游戏数据：灰机 Wiki｜%s %s｜抓取 %s"
        % (wiki.season, wiki.patch, wiki.fetched_at.isoformat()),
        "补丁索引：%s" % sources.patch_index_url,
    )


def _content_fitted_height(page) -> float:
    bottoms = [
        float(item["bottom"])
        for items in page.objects.values()
        for item in items
        if "bottom" in item
    ]
    if not bottoms:
        return MIN_PAGE_HEIGHT
    return min(PAGE_SIZE[1], max(MIN_PAGE_HEIGHT, max(bottoms) + 12 * mm))


def _write_content_fitted_pdf(source: BytesIO, output: Path) -> None:
    payload = source.getvalue()
    reader = PdfReader(BytesIO(payload))
    writer = PdfWriter()
    with pdfplumber.open(BytesIO(payload)) as layout:
        for page_number, (page, measured) in enumerate(
            zip(reader.pages, layout.pages),
            start=1,
        ):
            height = _content_fitted_height(measured)
            writer.add_page(page)
            output_page = writer.pages[-1]
            bottom = float(output_page.mediabox.top) - height
            output_page.mediabox.bottom = bottom
            output_page.cropbox.bottom = bottom

            footer_stream = BytesIO()
            footer = pdf_canvas.Canvas(
                footer_stream,
                pagesize=(PAGE_SIZE[0], height),
            )
            footer.setFont(PDF_FONT, 7)
            footer.setFillColor(colors.HexColor("#595959"))
            footer.drawRightString(
                PAGE_SIZE[0] - 12 * mm,
                8 * mm,
                "第 %d 页" % page_number,
            )
            footer.save()
            footer_page = PdfReader(BytesIO(footer_stream.getvalue())).pages[0]
            output_page.merge_translated_page(footer_page, 0, bottom)

    with output.open("wb") as stream:
        writer.write(stream)


def write_leaderboard_pdf(
    path: Path,
    spec,
    cards: Mapping[str, Iterable[OperatorCard]],
    operator_icon_dir: Path,
    gadget_icons: Mapping[str, Path],
    report_sources: ReportSources,
) -> None:
    """Write one PDF containing attack, defense, and patch-note sections."""
    from .leaderboards import _band_color, group_cards, patch_markers

    _register_fonts()
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    normalized = {
        side: tuple(sorted(cards.get(side, ()), key=lambda card: card.source_order))
        for side in SOURCE_SHEETS
    }
    markers = patch_markers(report_sources)
    story = []
    for side_index, side in enumerate(SOURCE_SHEETS):
        if side_index:
            story.append(PageBreak())
        story.append(
            _p("%s · %s" % (side, spec.title), _style(19, bold=True))
        )
        story.append(Spacer(1, 3 * mm))
        groups = group_cards(normalized[side], spec.key, side)
        for band, band_cards in groups.items():
            band_header = Table(
                [[_p(band, _style(11, bold=True, color="#FFFFFF", align=TA_CENTER))]],
                colWidths=[CONTENT_WIDTH],
                rowHeights=[9 * mm],
            )
            band_header.setStyle(
                TableStyle(
                    [
                        (
                            "BACKGROUND",
                            (0, 0),
                            (-1, -1),
                            colors.HexColor("#" + _band_color(spec.key, band, side)),
                        ),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ]
                )
            )
            chunks = [
                band_cards[index:index + CARDS_PER_ROW]
                for index in range(0, len(band_cards), CARDS_PER_ROW)
            ] or [()]
            block = [band_header]
            for chunk in chunks:
                cells = [
                    _card_flowable(
                        card,
                        Path(operator_icon_dir),
                        gadget_icons,
                        markers.get(card.name, ""),
                    )
                    for card in chunk
                ]
                cells.extend([""] * (CARDS_PER_ROW - len(cells)))
                row = Table([cells], colWidths=[78 * mm] * CARDS_PER_ROW)
                row.setStyle(
                    TableStyle(
                        [
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 1),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 1),
                            ("TOPPADDING", (0, 0), (-1, -1), 1),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                        ]
                    )
                )
                block.append(row)
            story.append(KeepTogether(block))
            story.append(Spacer(1, 2 * mm))
        for line in _source_lines(report_sources):
            story.append(_p(line, _style(7, color="#595959")))

    story.append(PageBreak())
    story.append(_p("补丁说明", _style(19, bold=True)))
    story.append(
        _p(
            "%s 视频评分之后至 %s"
            % (
                report_sources.rating.covered_through.isoformat(),
                report_sources.wiki.fetched_at.date().isoformat(),
            ),
            _style(9, color="#595959"),
        )
    )
    scores = {
        card.name: card.score
        for side in SOURCE_SHEETS
        for card in normalized[side]
    }
    for patch in report_sources.patches:
        story.append(Spacer(1, 3 * mm))
        story.append(
            _p(
                "%s · %s · %s"
                % (patch.patch, patch.released.isoformat(), patch.season_name),
                _style(12, bold=True, color="#FFFFFF"),
            )
        )
        rows = [["方向", "对象", "视频评分", "更新内容"]]
        for change in patch.changes:
            rows.append(
                [
                    change.direction,
                    change.subject,
                    str(scores.get(change.subject, "-")),
                    change.detail,
                ]
            )
        if len(rows) == 1:
            rows.append(["-", "-", "-", "无影响本报告字段的变更"])
        table = Table(
            [
                [_p(value, _style(8, bold=index == 0)) for index, value in enumerate(row)]
                for row in rows
            ],
            colWidths=[22 * mm, 35 * mm, 28 * mm, 305 * mm],
            repeatRows=1,
        )
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#BFBFBF")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ]
            )
        )
        story.append(table)
        story.append(_p(patch.wiki_url, _style(7, color="#4472C4")))
        story.append(_p(patch.official_url, _style(7, color="#4472C4")))

    rendered = BytesIO()
    document = SimpleDocTemplate(
        rendered,
        pagesize=PAGE_SIZE,
        leftMargin=PAGE_MARGIN,
        rightMargin=PAGE_MARGIN,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=spec.title,
        author="R6 干员中文榜单",
    )
    document.build(story)
    _write_content_fitted_pdf(rendered, output)
