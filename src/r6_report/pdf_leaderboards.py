"""Render searchable PDF companions for the five leaderboard workbooks."""

from io import BytesIO
from pathlib import Path
from typing import Iterable, Mapping, Tuple
from xml.sax.saxutils import escape

import pdfplumber
import pypdfium2 as pdfium
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
    KeepInFrame,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
    XPreformatted,
)
from PIL import Image as PillowImage

from . import report_theme as theme
from .gadget_slots import arrange_gadgets
from .sources import ReportSources
from .tier_chart import (
    CARDS_PER_ROW,
    OperatorCard,
    SOURCE_SHEETS,
    WeaponItem,
    operator_key,
)
from .tiers import TIER_COLORS


PDF_FONT = "R6Chinese"
PDF_SYMBOL_FONT = "R6Symbols"
PAGE_SIZE = A2
PAGE_MARGIN = 10 * mm
CONTENT_WIDTH = PAGE_SIZE[0] - 2 * PAGE_MARGIN
MIN_PAGE_HEIGHT = 100 * mm
PNG_DPI = 144


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
    if PDF_SYMBOL_FONT not in pdfmetrics.getRegisteredFontNames():
        symbol_candidates = (
            Path("C:/Windows/Fonts/seguisym.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        )
        symbol_path = next(
            (path for path in symbol_candidates if path.is_file()),
            None,
        )
        if symbol_path is not None:
            pdfmetrics.registerFont(
                TTFont(PDF_SYMBOL_FONT, str(symbol_path))
            )
        else:
            globals()["PDF_SYMBOL_FONT"] = PDF_FONT


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


def _nobr_p(text, style):
    return XPreformatted(escape(str(text)), style)


def _feature_p(label: str, present: bool, style):
    if not present:
        return _p(theme.feature_text(label, False), style)
    return Paragraph(
        "%s <font name=\"%s\">✓</font>"
        % (escape(label), PDF_SYMBOL_FONT),
        style,
    )


def _fit_image_size(
    path: Path,
    box_width: float,
    box_height: float,
) -> Tuple[float, float]:
    """Fit the visible part of an image without changing its aspect ratio."""
    with PillowImage.open(path) as image:
        visible_bounds = image.convert("RGBA").getchannel("A").getbbox()
    if visible_bounds is None:
        raise ValueError("image has no visible pixels: %s" % path)
    width = visible_bounds[2] - visible_bounds[0]
    height = visible_bounds[3] - visible_bounds[1]
    if width <= 0 or height <= 0:
        raise ValueError("image dimensions must be positive: %s" % path)
    scale = min(box_width / width, box_height / height)
    return (width * scale, height * scale)


def _cropped_image_source(path: Path) -> BytesIO:
    """Return a PNG stream cropped to the source image's visible pixels."""
    with PillowImage.open(path) as image:
        icon = image.convert("RGBA")
        visible_bounds = icon.getchannel("A").getbbox()
        if visible_bounds is None:
            raise ValueError("image has no visible pixels: %s" % path)
        icon = icon.crop(visible_bounds)
        stream = BytesIO()
        icon.save(stream, format="PNG")
    stream.seek(0)
    return stream


def _patch_text_color(
    row_index: int,
    column_index: int,
    direction: str,
) -> str:
    if row_index == 0 or (
        column_index == 0
        and direction in theme.PATCH_DIRECTION_COLOURS
    ):
        return "#" + theme.COLOURS["white"]
    return "#" + theme.COLOURS["text"]


def _card_flowable(
    card: OperatorCard,
    badge_dir: Path,
    gadget_icons: Mapping[str, Path],
    weapon_icons: Mapping[str, Path],
    marker: str,
):
    badge_path = badge_dir / (operator_key(card.name) + ".png")
    badge = Image(str(badge_path), width=15 * mm, height=15 * mm)
    badge.hAlign = "CENTER"
    badge_name_size = theme.PDF_FONT_SIZES["gadget"]
    badge_name_width = pdfmetrics.stringWidth(
        card.name,
        PDF_FONT,
        badge_name_size,
    )
    if badge_name_width > 14 * mm:
        badge_name_size *= (14 * mm) / badge_name_width
    badge_name = _nobr_p(
        card.name,
        _style(
            badge_name_size,
            bold=True,
            color="#" + theme.COLOURS["text_strong"],
            align=TA_CENTER,
        ),
    )
    badge_panel = Table(
        [[badge], [badge_name]],
        colWidths=[15 * mm],
        rowHeights=[15 * mm, 4 * mm],
    )
    badge_panel.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    tier = _p(
        "%s%s" % (card.tier, marker),
        _style(
            theme.PDF_FONT_SIZES["body"],
            bold=True,
            color=(
                "#" + theme.COLOURS["white"]
                if card.tier in ("S", "A", "D", "F")
                else "#" + theme.COLOURS["text_strong"]
            ),
            align=TA_CENTER,
        ),
    )
    speed = _p(
        "%d速" % card.speed,
        _style(
            theme.PDF_FONT_SIZES["body"],
            color="#" + theme.COLOURS["text_strong"],
            align=TA_CENTER,
        ),
    )
    summary = Table(
        [[tier, speed]],
        colWidths=[15.5 * mm, 15.5 * mm],
        rowHeights=[7 * mm],
    )
    summary.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (0, 0),
                    colors.HexColor("#" + TIER_COLORS[card.tier]),
                ),
                (
                    "BACKGROUND",
                    (1, 0),
                    (1, 0),
                    colors.HexColor("#" + theme.COLOURS["card_fill"]),
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.55,
                    colors.HexColor("#" + theme.COLOURS["card_grid"]),
                ),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 1),
                ("RIGHTPADDING", (0, 0), (-1, -1), 1),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    secondary_shotgun_text = theme.feature_text(
        "副喷",
        card.has_secondary_shotgun,
    )
    primary_sniper_text = theme.feature_text(
        "主狙",
        card.has_semiautomatic,
    )
    primary_fields = [
        _weapon_field_flowable(
            _nobr_p(
                theme.rpm_text(
                    "主",
                    (weapon.firerate,) if weapon.firerate is not None else (),
                ),
                _style(theme.PDF_FONT_SIZES["body"]),
            ),
            (weapon,),
            weapon_icons,
            field_width_mm=30,
        )
        for weapon in card.primary_weapons[:3]
    ]
    primary_fields.extend([""] * (3 - len(primary_fields)))
    secondary_fields = [
        _weapon_field_flowable(
            _nobr_p(
                theme.rpm_text(
                    "副",
                    (weapon.firerate,) if weapon.firerate is not None else (),
                ),
                _style(theme.PDF_FONT_SIZES["body"]),
            ),
            (weapon,),
            weapon_icons,
            field_width_mm=30,
        )
        for weapon in card.secondary_weapons[:2]
    ]
    secondary_fields.extend([""] * (2 - len(secondary_fields)))
    primary_sniper = _weapon_field_flowable(
        _feature_p(
            "主狙",
            card.has_semiautomatic,
            _style(theme.PDF_FONT_SIZES["body"]),
        ),
        card.semiautomatic_weapons,
        weapon_icons,
        field_width_mm=30,
    )
    secondary_shotgun = _weapon_field_flowable(
        _feature_p(
            "副喷",
            card.has_secondary_shotgun,
            _style(theme.PDF_FONT_SIZES["body"]),
        ),
        card.secondary_shotguns,
        weapon_icons,
        field_width_mm=30,
    )
    details = Table(
        [
            [
                summary,
                "",
                "",
                primary_sniper,
                "",
                "",
            ],
            [primary_fields[0], "", "", secondary_shotgun, "", ""],
            [primary_fields[1], "", "", secondary_fields[0], "", ""],
            [primary_fields[2], "", "", secondary_fields[1], "", ""],
        ],
        colWidths=[(62 / 6) * mm] * 6,
        rowHeights=[
            7 * mm,
            theme.PDF_CARD_BODY_ROW_MM * mm,
            theme.PDF_CARD_BODY_ROW_MM * mm,
            theme.PDF_CARD_BODY_ROW_MM * mm,
        ],
    )
    details.setStyle(
        TableStyle(
            [
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.55,
                    colors.HexColor("#" + theme.COLOURS["card_grid"]),
                ),
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.85,
                    colors.HexColor("#" + theme.COLOURS["card_divider"]),
                ),
                (
                    "LINEBEFORE",
                    (3, 0),
                    (3, -1),
                    0.85,
                    colors.HexColor("#" + theme.COLOURS["card_divider"]),
                ),
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, -1),
                    colors.HexColor("#" + theme.COLOURS["card_fill"]),
                ),
                (
                    "LINEBELOW",
                    (3, 1),
                    (5, 1),
                    0.85,
                    colors.HexColor("#" + theme.COLOURS["card_divider"]),
                ),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("SPAN", (0, 0), (2, 0)),
                ("SPAN", (3, 0), (5, 0)),
                ("SPAN", (0, 1), (2, 1)),
                ("SPAN", (3, 1), (5, 1)),
                ("SPAN", (0, 2), (2, 2)),
                ("SPAN", (3, 2), (5, 2)),
                ("SPAN", (0, 3), (2, 3)),
                ("SPAN", (3, 3), (5, 3)),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                ("LEFTPADDING", (0, 0), (0, 0), 0),
                ("RIGHTPADDING", (0, 0), (0, 0), 0),
                ("TOPPADDING", (0, 0), (0, 0), 0),
                ("BOTTOMPADDING", (0, 0), (0, 0), 0),
            ]
            + [
                (
                    "BACKGROUND",
                    (column, row),
                    (column, row),
                    colors.HexColor("#" + theme.MISSING_FILL),
                )
                for column, row, value in (
                    (3, 0, primary_sniper_text),
                    (3, 1, secondary_shotgun_text),
                )
                if theme.is_missing_field(value)
            ]
        )
    )
    gadget_cells = []
    for gadget in arrange_gadgets(card.side, card.gadgets):
        if gadget is None:
            gadget_cells.append("")
            continue
        path = gadget_icons[gadget.name]
        icon_box = theme.PDF_GADGET_ICON_BOX_MM * mm
        image_width, image_height = _fit_image_size(
            path,
            icon_box,
            icon_box,
        )
        icon = Image(
            _cropped_image_source(path),
            width=image_width,
            height=image_height,
        )
        icon.hAlign = "CENTER"
        gadget_cells.append(icon)
    gadgets = Table(
        [gadget_cells],
        colWidths=[11 * mm] * 7,
        rowHeights=[8 * mm],
    )
    gadgets.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.55,
                    colors.HexColor("#" + theme.COLOURS["card_grid"]),
                ),
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.85,
                    colors.HexColor("#" + theme.COLOURS["card_divider"]),
                ),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    card_table = Table(
        [[badge_panel, details], [gadgets, ""]],
        colWidths=[15 * mm, 62 * mm],
        rowHeights=[25 * mm, 8 * mm],
    )
    card_table.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, -1),
                    colors.HexColor("#" + theme.COLOURS["card_fill"]),
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.55,
                    colors.HexColor("#" + theme.COLOURS["card_grid"]),
                ),
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.85,
                    colors.HexColor("#" + theme.COLOURS["card_divider"]),
                ),
                (
                    "LINEBEFORE",
                    (1, 0),
                    (1, 0),
                    0.85,
                    colors.HexColor("#" + theme.COLOURS["card_divider"]),
                ),
                (
                    "LINEABOVE",
                    (0, 1),
                    (-1, 1),
                    0.85,
                    colors.HexColor("#" + theme.COLOURS["card_divider"]),
                ),
                ("SPAN", (0, 1), (1, 1)),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                ("TOPPADDING", (0, 0), (-1, 0), 0),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 0),
                ("LEFTPADDING", (0, 1), (-1, 1), 0),
                ("RIGHTPADDING", (0, 1), (-1, 1), 0),
                ("TOPPADDING", (0, 1), (-1, 1), 0),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 0),
            ]
        )
    )
    return card_table


def _weapon_field_flowable(
    text,
    items: Tuple[WeaponItem, ...],
    weapon_icons: Mapping[str, Path],
    *,
    field_width_mm: float = 28,
    icon_area_mm: float = theme.PDF_WEAPON_ICON_AREA_MM,
):
    if not items:
        return text
    icon_area = icon_area_mm * mm
    slot_width = icon_area / len(items)
    icon_cells = []
    for item in items:
        path = weapon_icons.get(item.icon_key)
        if path is None:
            raise ValueError("找不到枪械图标：%s" % item.name)
        width, height = _fit_image_size(
            path,
            slot_width,
            theme.PDF_WEAPON_ICON_HEIGHT_MM * mm,
        )
        icon = Image(
            _cropped_image_source(path),
            width=width,
            height=height,
        )
        icon.hAlign = "CENTER"
        icon_cells.append(icon)
    icons = Table(
        [icon_cells],
        colWidths=[slot_width] * len(items),
        rowHeights=[theme.PDF_WEAPON_ICON_HEIGHT_MM * mm],
    )
    icons.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    field_width = field_width_mm * mm
    text_width = field_width - icon_area
    text_box = KeepInFrame(
        text_width,
        4 * mm,
        [text],
        mode="shrink",
        mergeSpace=False,
    )
    field = Table(
        [[text_box, icons]],
        colWidths=[text_width, icon_area],
        rowHeights=[4 * mm],
    )
    field.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    return field


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
            footer.setFont(PDF_FONT, theme.PDF_FONT_SIZES["page"])
            footer.setFillColor(
                colors.HexColor("#" + theme.COLOURS["text_muted"])
            )
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


def write_pdf_pages_as_png(
    pdf_path: Path,
    output_dir: Path,
    *,
    dpi: int = PNG_DPI,
) -> Tuple[Path, ...]:
    """Rasterize every PDF page as an unchanged, standalone PNG image."""
    if dpi <= 0:
        raise ValueError("dpi must be positive")

    source = Path(pdf_path)
    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    document = pdfium.PdfDocument(source)
    paths = []
    try:
        for page_index in range(len(document)):
            page = document[page_index]
            bitmap = None
            image = None
            try:
                bitmap = page.render(
                    scale=dpi / 72,
                    rotation=0,
                    may_draw_forms=True,
                )
                image = bitmap.to_pil()
                output = destination / ("第%d页.png" % (page_index + 1))
                image.save(output, format="PNG", dpi=(dpi, dpi))
                paths.append(output)
            finally:
                if image is not None:
                    image.close()
                if bitmap is not None:
                    bitmap.close()
                page.close()
    finally:
        document.close()
    return tuple(paths)


def write_leaderboard_pdf(
    path: Path,
    spec,
    cards: Mapping[str, Iterable[OperatorCard]],
    operator_icon_dir: Path,
    gadget_icons: Mapping[str, Path],
    weapon_icons: Mapping[str, Path],
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
            _p(
                "%s · %s" % (side, spec.title),
                _style(
                    theme.PDF_FONT_SIZES["title"],
                    bold=True,
                    color="#" + theme.COLOURS["text"],
                ),
            )
        )
        story.append(Spacer(1, 3 * mm))
        groups = group_cards(normalized[side], spec.key, side)
        for band, band_cards in groups.items():
            band_header = Table(
                [[
                    _p(
                        band,
                        _style(
                            theme.PDF_FONT_SIZES["band"],
                            bold=True,
                            color="#" + theme.COLOURS["white"],
                            align=TA_CENTER,
                        ),
                    )
                ]],
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
                        weapon_icons,
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
            story.append(
                _p(
                    line,
                    _style(
                        theme.PDF_FONT_SIZES["source"],
                        color="#" + theme.COLOURS["text_muted"],
                    ),
                )
            )

    story.append(PageBreak())
    story.append(
        _p(
            "补丁说明",
            _style(
                theme.PDF_FONT_SIZES["title"],
                bold=True,
                color="#" + theme.COLOURS["text"],
            ),
        )
    )
    story.append(
        _p(
            "%s 视频评分之后至 %s"
            % (
                report_sources.rating.covered_through.isoformat(),
                report_sources.wiki.fetched_at.date().isoformat(),
            ),
            _style(
                theme.PDF_FONT_SIZES["body"],
                color="#" + theme.COLOURS["text_muted"],
            ),
        )
    )
    scores = {
        card.name: card.score
        for side in SOURCE_SHEETS
        for card in normalized[side]
    }
    for patch in report_sources.patches:
        story.append(Spacer(1, 3 * mm))
        patch_header = Table(
            [[
                _p(
                    "%s · %s · %s"
                    % (
                        patch.patch,
                        patch.released.isoformat(),
                        patch.season_name,
                    ),
                    _style(
                        theme.PDF_FONT_SIZES["patch_header"],
                        bold=True,
                        color="#" + theme.COLOURS["white"],
                    ),
                )
            ]],
            colWidths=[CONTENT_WIDTH],
            rowHeights=[9 * mm],
        )
        patch_header.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, -1),
                        colors.HexColor(
                            "#" + theme.COLOURS["patch_header_fill"]
                        ),
                    ),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        story.append(patch_header)
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
        table_rows = []
        for row_index, row in enumerate(rows):
            table_row = []
            for column_index, value in enumerate(row):
                direction = row[0] if row_index else None
                table_row.append(
                    _p(
                        value,
                        _style(
                            theme.PDF_FONT_SIZES["patch_table"],
                            bold=(row_index == 0 or column_index == 0),
                            color=_patch_text_color(
                                row_index,
                                column_index,
                                direction,
                            ),
                        ),
                    )
                )
            table_rows.append(table_row)
        table = Table(
            table_rows,
            colWidths=[22 * mm, 35 * mm, 28 * mm, 305 * mm],
            repeatRows=1,
        )
        table_commands = [
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor("#" + theme.COLOURS["section_fill"]),
            ),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.35,
                colors.HexColor("#BFBFBF"),
            ),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]
        for row_index, row in enumerate(rows[1:], start=1):
            direction = row[0]
            if direction in theme.PATCH_DIRECTION_COLOURS:
                table_commands.append(
                    (
                        "BACKGROUND",
                        (0, row_index),
                        (0, row_index),
                        colors.HexColor(
                            "#" + theme.PATCH_DIRECTION_COLOURS[direction]
                        ),
                    )
                )
            for column_index, value in enumerate(row):
                if str(value).strip() == "-":
                    table_commands.append(
                        (
                            "BACKGROUND",
                            (column_index, row_index),
                            (column_index, row_index),
                            colors.HexColor("#" + theme.MISSING_FILL),
                        )
                    )
        table.setStyle(
            TableStyle(table_commands)
        )
        story.append(table)
        story.append(
            _p(
                patch.wiki_url,
                _style(
                    theme.PDF_FONT_SIZES["source"],
                    color="#" + theme.COLOURS["link"],
                ),
            )
        )
        story.append(
            _p(
                patch.official_url,
                _style(
                    theme.PDF_FONT_SIZES["source"],
                    color="#" + theme.COLOURS["link"],
                ),
            )
        )

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
