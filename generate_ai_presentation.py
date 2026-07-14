# -*- coding: utf-8 -*-
"""
================================================================================
  Генератор освітньої презентації (.pptx)
  Тема: «Як працює штучний інтелект: від нейрона до великих мовних моделей»
--------------------------------------------------------------------------------
  Аудиторія : широка публіка / новачки без технічного бекґраунду
  Джерела   : транскрипт відео (SRT) + плейсхолдери під кадри з оригіналу
  Стиль     : освітній / педагогічний
  Тон       : авторитетний / експертний
  Мова конт.: виключно українська
--------------------------------------------------------------------------------
  ВСТАНОВЛЕННЯ ЗАЛЕЖНОСТЕЙ:
      pip install python-pptx
  ЗАПУСК:
      python generate_ai_presentation.py
  РЕЗУЛЬТАТ:
      Як_працює_штучний_інтелект.pptx  (35 слайдів, детальні нотатки доповідача)
================================================================================
"""

from __future__ import annotations

import os
from typing import List, Tuple, Dict, Optional, Sequence

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# Дата-візуалізація
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION


# ==============================================================================
#  ПАЛІТРА ТА ТИПОГРАФІКА  (високий контраст, наукова стриманість)
# ==============================================================================
class Theme:
    """Централізована палітра та типографіка презентації.

    Принцип: темні заголовки зі світлим текстом + світлі тіла слайдів
    з темним текстом — для максимальної читабельності та академічної
    строгості.
    """

    # --- Кольори ---
    DARK: RGBColor = RGBColor(0x10, 0x24, 0x3B)        # глибокий морський синій (шапки)
    DARK_ALT: RGBColor = RGBColor(0x18, 0x33, 0x52)    # трохи світліший синій
    ACCENT: RGBColor = RGBColor(0x2E, 0x9C, 0xDB)      # акцентний блакитний
    ACCENT_2: RGBColor = RGBColor(0x17, 0xB2, 0x8A)    # смарагдовий (другий акцент)
    WARN: RGBColor = RGBColor(0xE8, 0x7A, 0x3E)        # теплий помаранчевий (акценти-ризики)

    LIGHT_BG: RGBColor = RGBColor(0xF4, 0xF6, 0xFA)    # світле тло тіла
    PANEL_BG: RGBColor = RGBColor(0xE7, 0xEE, 0xF6)    # тло візуальної панелі
    WHITE: RGBColor = RGBColor(0xFF, 0xFF, 0xFF)

    TEXT_DARK: RGBColor = RGBColor(0x1B, 0x24, 0x33)   # основний темний текст
    TEXT_LIGHT: RGBColor = RGBColor(0xFF, 0xFF, 0xFF)  # текст на темному
    MUTED: RGBColor = RGBColor(0x60, 0x70, 0x86)       # приглушений (підписи, футер)
    HAIRLINE: RGBColor = RGBColor(0xCF, 0xD8, 0xE3)    # тонкі лінії таблиць

    # --- Шрифти ---
    TITLE_FONT: str = "Georgia"      # серифний, підтримує кирилицю (заголовки)
    BODY_FONT: str = "Tahoma"        # без засічок, чудова кирилиця (тіло)
    MONO_FONT: str = "Consolas"

    # --- Кеглі ---
    SZ_TITLE: int = 32
    SZ_SUBTITLE: int = 20
    SZ_BODY: int = 24
    SZ_BODY_SM: int = 18
    SZ_FOOTER: int = 9
    SZ_QUESTION: int = 40


class EducationalPresentationBuilder:
    """Будівник академічної освітньої презентації на python-pptx.

    Клас інкапсулює єдину сітку, палітру та набір типів слайдів. Кожен
    змістовий слайд дотримується правила «≤30 % ширини під візуальний
    елемент» — решта (≥70 %) віддана під структурований текст.

    Геометрія (широкий формат 16:9):
        ширина слайда  = 13.333"
        30 % ширини    = 4.000"  →  візуальний блок обмежено 3.8"
    """

    # --- Сітка слайда (дюйми) ---
    SLIDE_W: float = 13.333
    SLIDE_H: float = 7.5
    MARGIN: float = 0.55
    HEADER_H: float = 1.15
    FOOTER_Y: float = 7.02

    # Візуальна панель (правий бік) — суворо ≤30 %
    VIS_W: float = 3.80
    VIS_X: float = SLIDE_W - MARGIN - 3.80   # = 8.983"
    VIS_Y: float = HEADER_H + 0.35
    VIS_H: float = 4.90

    # Текстова колонка (ліві ~70 %)
    TXT_X: float = MARGIN
    TXT_W: float = VIS_X - MARGIN - 0.35     # ≈ 7.55"

    def __init__(self, output_path: str = "Як_працює_штучний_інтелект.pptx") -> None:
        """Ініціалізує порожню презентацію 16:9.

        Args:
            output_path: шлях до вихідного .pptx файлу.
        """
        self.output_path: str = output_path
        self.prs: Presentation = Presentation()
        self.prs.slide_width = Inches(self.SLIDE_W)
        self.prs.slide_height = Inches(self.SLIDE_H)
        self._blank = self.prs.slide_layouts[6]  # порожній макет

    # ------------------------------------------------------------------ #
    #  НИЗЬКОРІВНЕВІ ДОПОМІЖНІ МЕТОДИ
    # ------------------------------------------------------------------ #
    def _new_slide(self, bg: RGBColor = Theme.LIGHT_BG):
        """Створює порожній слайд із суцільним тлом."""
        slide = self.prs.slides.add_slide(self._blank)
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = bg
        return slide

    def _rect(self, slide, x, y, w, h, color: RGBColor,
              shape=MSO_SHAPE.RECTANGLE, line_color: Optional[RGBColor] = None,
              line_w: float = 0.75):
        """Малює заповнену фігуру без тіні; повертає shape."""
        sp = slide.shapes.add_shape(shape, Inches(x), Inches(y),
                                    Inches(w), Inches(h))
        sp.fill.solid()
        sp.fill.fore_color.rgb = color
        if line_color is None:
            sp.line.fill.background()
        else:
            sp.line.color.rgb = line_color
            sp.line.width = Pt(line_w)
        sp.shadow.inherit = False
        return sp

    def _text(self, slide, x, y, w, h, runs, *, font=Theme.BODY_FONT,
              size=Theme.SZ_BODY, color=Theme.TEXT_DARK, bold=False,
              align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.08,
              autofit=True):
        """Універсальний текстовий блок.

        Args:
            runs: або str, або список абзаців (кожен — str чи (str, dict)).
        """
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        if autofit:
            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.vertical_anchor = anchor
        tf.margin_left = Pt(4)
        tf.margin_right = Pt(4)
        tf.margin_top = Pt(2)
        tf.margin_bottom = Pt(2)

        if isinstance(runs, str):
            runs = [runs]

        for i, item in enumerate(runs):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = line_spacing
            p.space_after = Pt(4)
            # елемент може нести власні перевизначення стилю
            txt, over = (item if isinstance(item, tuple) else (item, {}))
            r = p.add_run()
            r.text = txt
            f = r.font
            f.name = over.get("font", font)
            f.size = Pt(over.get("size", size))
            f.bold = over.get("bold", bold)
            f.color.rgb = over.get("color", color)
        return tb

    def _bullets(self, slide, x, y, w, h, items: Sequence, *,
                 size=Theme.SZ_BODY, color=Theme.TEXT_DARK,
                 marker="—", marker_color=Theme.ACCENT, line_spacing=1.1):
        """Марковані пункти з кольоровим маркером-тире."""
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.line_spacing = line_spacing
            p.space_after = Pt(8)
            txt, lvl = (item if isinstance(item, tuple) else (item, 0))
            m = p.add_run()
            m.text = ("    " * lvl) + marker + "  "
            m.font.name = Theme.BODY_FONT
            m.font.size = Pt(size)
            m.font.bold = True
            m.font.color.rgb = marker_color
            r = p.add_run()
            r.text = txt
            r.font.name = Theme.BODY_FONT
            r.font.size = Pt(size if lvl == 0 else size - 3)
            r.font.color.rgb = color
        return tb

    def _header(self, slide, title: str, kicker: Optional[str] = None,
                number: Optional[int] = None) -> None:
        """Темна шапка зі світлим заголовком, лівим акцентом та номером."""
        self._rect(slide, 0, 0, self.SLIDE_W, self.HEADER_H, Theme.DARK)
        # акцентна вертикальна риска
        self._rect(slide, 0, 0, 0.16, self.HEADER_H, Theme.ACCENT)
        ky = 0.14
        if kicker:
            self._text(slide, self.MARGIN, ky, self.SLIDE_W - 3.0, 0.34,
                       kicker.upper(), font=Theme.BODY_FONT,
                       size=11, color=Theme.ACCENT, bold=True, autofit=False)
        self._text(slide, self.MARGIN, 0.40 if kicker else 0.24,
                   self.SLIDE_W - 3.0, 0.72, title,
                   font=Theme.TITLE_FONT, size=Theme.SZ_TITLE,
                   color=Theme.TEXT_LIGHT, bold=True,
                   anchor=MSO_ANCHOR.MIDDLE, autofit=False)
        if number is not None:
            self._text(slide, self.SLIDE_W - 1.35, 0.30, 0.9, 0.6,
                       f"{number:02d}", font=Theme.TITLE_FONT, size=26,
                       color=Theme.ACCENT, bold=True,
                       align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE,
                       autofit=False)

    def _footer(self, slide, source_citation: str) -> None:
        """Дрібний підпис-цитата внизу слайда (окрема вимога)."""
        if not source_citation:
            return
        self._text(slide, self.MARGIN, self.FOOTER_Y, self.SLIDE_W - 2 * self.MARGIN,
                   0.36, "Джерело: " + source_citation,
                   font=Theme.BODY_FONT, size=Theme.SZ_FOOTER,
                   color=Theme.MUTED, autofit=False)

    def _visual_panel(self, slide, image_path: Optional[str],
                      caption: str, *, x: Optional[float] = None,
                      y: Optional[float] = None, w: Optional[float] = None,
                      h: Optional[float] = None) -> None:
        """Візуальний блок ≤30 % ширини: реальне фото або плейсхолдер-кадр.

        Якщо `image_path` існує на диску — вбудовує зображення. Інакше
        малює рамку-плейсхолдер із чіткою інструкцією, який саме кадр з
        оригінального відео сюди вставити (з таймкодом).
        """
        x = self.VIS_X if x is None else x
        y = self.VIS_Y if y is None else y
        w = self.VIS_W if w is None else w
        h = self.VIS_H if h is None else h

        # підкладка панелі
        self._rect(slide, x, y, w, h, Theme.PANEL_BG,
                   shape=MSO_SHAPE.ROUNDED_RECTANGLE, line_color=Theme.HAIRLINE)
        if image_path and os.path.isfile(image_path):
            try:
                pic = slide.shapes.add_picture(image_path, Inches(x + 0.12),
                                               Inches(y + 0.12), width=Inches(w - 0.24))
                # обмеження висоти, щоб не вилізти за панель
                if pic.height > Inches(h - 0.9):
                    pic.height = Inches(h - 0.9)
                    pic.width = int(pic.height * (pic.image.size[0] / pic.image.size[1]))
                    pic.left = Inches(x + (w - Emu(pic.width).inches) / 2)
            except Exception:
                self._placeholder_body(slide, x, y, w, h, caption)
        else:
            self._placeholder_body(slide, x, y, w, h, caption)

        # підпис під панеллю
        self._text(slide, x, y + h + 0.04, w, 0.5, caption,
                   font=Theme.BODY_FONT, size=10, color=Theme.MUTED,
                   align=PP_ALIGN.CENTER, autofit=False)

    def _placeholder_body(self, slide, x, y, w, h, caption: str) -> None:
        """Вміст плейсхолдера, коли реального зображення немає."""
        # «іконка» кадру
        icon = self._rect(slide, x + w / 2 - 0.55, y + h / 2 - 0.85, 1.10, 0.80,
                          Theme.DARK_ALT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        self._text(slide, x + w / 2 - 0.55, y + h / 2 - 0.85, 1.10, 0.80, "▣",
                   size=30, color=Theme.ACCENT, align=PP_ALIGN.CENTER,
                   anchor=MSO_ANCHOR.MIDDLE, autofit=False)
        self._text(slide, x + 0.2, y + h / 2 + 0.10, w - 0.4, h / 2 - 0.2,
                   "КАДР З ВІДЕО", size=11, color=Theme.MUTED, bold=True,
                   align=PP_ALIGN.CENTER, autofit=False)

    def _add_notes(self, slide, notes_text: str, source_citation: str = "") -> None:
        """Записує скрипт доповідача в нотатки слайда.

        Наприкінці приєднує професійну цитату (окрема вимога).
        """
        note = notes_text.strip()
        if source_citation:
            note += "\n\n———\nДжерело / Reference: " + source_citation
        slide.notes_slide.notes_text_frame.text = note

    def _style_table(self, table, headers: Sequence[str],
                     rows_data: Sequence[Sequence[str]]) -> None:
        """Сучасна таблиця: темна шапка, світлі рядки, тонкі роздільники."""
        # прибираємо «зебру» стандартного стилю
        tbl = table._tbl
        # заголовки
        for j, htext in enumerate(headers):
            cell = table.cell(0, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = Theme.DARK
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Pt(8); cell.margin_right = Pt(8)
            cell.margin_top = Pt(5); cell.margin_bottom = Pt(5)
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            r = p.add_run(); r.text = htext
            r.font.name = Theme.BODY_FONT; r.font.size = Pt(15)
            r.font.bold = True; r.font.color.rgb = Theme.TEXT_LIGHT
        # тіло
        for i, row in enumerate(rows_data, start=1):
            for j, val in enumerate(row):
                cell = table.cell(i, j)
                cell.fill.solid()
                cell.fill.fore_color.rgb = Theme.WHITE if i % 2 else Theme.LIGHT_BG
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = Pt(8); cell.margin_right = Pt(8)
                cell.margin_top = Pt(4); cell.margin_bottom = Pt(4)
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT
                r = p.add_run(); r.text = str(val)
                r.font.name = Theme.BODY_FONT; r.font.size = Pt(13)
                r.font.color.rgb = Theme.TEXT_DARK
                r.font.bold = (j == 0)

    # ================================================================== #
    #  ТИПИ СЛАЙДІВ
    # ================================================================== #
    def add_title_slide(self, title: str, subtitle: str, *,
                        presenter: str = "", notes_text: str = "",
                        source_citation: str = "", image_path: Optional[str] = None) -> None:
        """Титульний слайд: темне тло, велика серифна назва."""
        slide = self._new_slide(bg=Theme.DARK)
        # декоративні смуги
        self._rect(slide, 0, 0, 0.28, self.SLIDE_H, Theme.ACCENT)
        self._rect(slide, 0, self.SLIDE_H - 0.28, self.SLIDE_W, 0.28, Theme.ACCENT_2)
        self._text(slide, 1.1, 1.55, 11.2, 0.5, "ОСВІТНІЙ КУРС · ШТУЧНИЙ ІНТЕЛЕКТ",
                   font=Theme.BODY_FONT, size=15, color=Theme.ACCENT, bold=True, autofit=False)
        self._text(slide, 1.1, 2.15, 11.2, 2.3, title, font=Theme.TITLE_FONT,
                   size=48, color=Theme.TEXT_LIGHT, bold=True, line_spacing=1.02, autofit=False)
        self._text(slide, 1.12, 4.55, 11.0, 1.0, subtitle, font=Theme.BODY_FONT,
                   size=Theme.SZ_SUBTITLE, color=RGBColor(0xC7, 0xD6, 0xE6), autofit=False)
        if presenter:
            self._text(slide, 1.12, 6.4, 11.0, 0.5, presenter, font=Theme.BODY_FONT,
                       size=13, color=Theme.MUTED, autofit=False)
        self._add_notes(slide, notes_text, source_citation)

    def add_agenda_slide(self, items: Sequence[str], *, title: str = "Програма",
                        number: Optional[int] = None, notes_text: str = "",
                        source_citation: str = "", image_path: Optional[str] = None) -> None:
        """Слайд-порядок денний: двоколонковий список розділів."""
        slide = self._new_slide()
        self._header(slide, title, kicker="Огляд курсу", number=number)
        half = (len(items) + 1) // 2
        col1, col2 = items[:half], items[half:]
        y0 = self.HEADER_H + 0.5
        self._agenda_col(slide, self.MARGIN, y0, 5.9, col1, start=1)
        self._agenda_col(slide, 6.9, y0, 5.9, col2, start=half + 1)
        self._footer(slide, source_citation)
        self._add_notes(slide, notes_text, source_citation)

    def _agenda_col(self, slide, x, y, w, items, start=1):
        for k, it in enumerate(items):
            yy = y + k * 0.86
            self._rect(slide, x, yy, 0.62, 0.62, Theme.ACCENT,
                       shape=MSO_SHAPE.OVAL)
            self._text(slide, x, yy, 0.62, 0.62, f"{start + k}",
                       font=Theme.TITLE_FONT, size=18, color=Theme.WHITE, bold=True,
                       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, autofit=False)
            self._text(slide, x + 0.85, yy, w - 0.85, 0.62, it,
                       font=Theme.BODY_FONT, size=17, color=Theme.TEXT_DARK,
                       anchor=MSO_ANCHOR.MIDDLE, autofit=False)

    def add_concept_slide(self, concept: str, explanation: Sequence, *,
                        image_path: Optional[str] = None, visual_caption: str = "КАДР З ВІДЕО",
                        kicker: str = "Концепція", number: Optional[int] = None,
                        lead: str = "", notes_text: str = "",
                        source_citation: str = "") -> None:
        """Концептуальний слайд: текст ліворуч (≥70 %), візуал ≤30 % праворуч."""
        slide = self._new_slide()
        self._header(slide, concept, kicker=kicker, number=number)
        y = self.HEADER_H + 0.4
        if lead:
            self._text(slide, self.TXT_X, y, self.TXT_W, 0.9, lead,
                       font=Theme.BODY_FONT, size=17, color=Theme.ACCENT,
                       bold=True, line_spacing=1.12)
            y += 1.0
        self._bullets(slide, self.TXT_X, y, self.TXT_W, self.VIS_H - (y - (self.HEADER_H + 0.4)),
                      explanation, size=Theme.SZ_BODY - 2)
        self._visual_panel(slide, image_path, visual_caption)
        self._footer(slide, source_citation)
        self._add_notes(slide, notes_text, source_citation)

    def add_chart_slide(self, title: str, chart_type: str, chart_data: Dict, *,
                        insight: Sequence = (), kicker: str = "Дані",
                        number: Optional[int] = None, notes_text: str = "",
                        source_citation: str = "", image_path: Optional[str] = None) -> None:
        """Слайд із реальним графіком (pptx.chart). Fallback → текст."""
        slide = self._new_slide()
        self._header(slide, title, kicker=kicker, number=number)
        # ліворуч — короткий інсайт-текст (~40 %)
        if insight:
            self._bullets(slide, self.MARGIN, self.HEADER_H + 0.5, 4.7, 4.6,
                          insight, size=17)
        # праворуч — графік
        cx, cy, cw, ch = 5.5, self.HEADER_H + 0.35, 7.3, 4.9
        try:
            cd = CategoryChartData()
            cd.categories = chart_data["categories"]
            for name, vals in chart_data["series"].items():
                cd.add_series(name, vals)
            xl_type = {
                "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
                "barh": XL_CHART_TYPE.BAR_CLUSTERED,
                "line": XL_CHART_TYPE.LINE_MARKERS,
                "pie": XL_CHART_TYPE.PIE,
            }.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)
            gframe = slide.shapes.add_chart(xl_type, Inches(cx), Inches(cy),
                                            Inches(cw), Inches(ch), cd)
            chart = gframe.chart
            chart.has_title = False
            chart.has_legend = len(chart_data["series"]) > 1
            if chart.has_legend:
                chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                chart.legend.include_in_layout = False
                chart.legend.font.size = Pt(11)
            # кольори серій + підписи
            palette = [Theme.ACCENT, Theme.ACCENT_2, Theme.WARN, Theme.DARK_ALT]
            for si, plot_series in enumerate(chart.series):
                plot_series.format.fill.solid()
                plot_series.format.fill.fore_color.rgb = palette[si % len(palette)]
            try:
                plot = chart.plots[0]
                plot.has_data_labels = True
                plot.data_labels.font.size = Pt(11)
                plot.data_labels.font.bold = True
                plot.data_labels.number_format_is_linked = False
            except Exception:
                pass
            try:
                chart.category_axis.tick_labels.font.size = Pt(11)
                chart.value_axis.has_major_gridlines = True
                chart.value_axis.tick_labels.font.size = Pt(10)
            except Exception:
                pass
        except (ValueError, TypeError, KeyError, Exception):
            # ── Резервний текстовий слайд замість краху ──
            lines = []
            cats = chart_data.get("categories", [])
            for name, vals in chart_data.get("series", {}).items():
                for c, v in zip(cats, vals):
                    lines.append(f"{c}: {v} ({name})")
            self._bullets(slide, 5.5, self.HEADER_H + 0.5, 7.3, 4.8,
                          lines or ["Дані недоступні"], size=16)
        self._footer(slide, source_citation)
        self._add_notes(slide, notes_text, source_citation)

    def add_table_slide(self, title: str, headers: Sequence[str],
                        rows_data: Sequence[Sequence[str]], *,
                        image_path: Optional[str] = None,
                        visual_caption: str = "КАДР З ВІДЕО",
                        with_visual: bool = True, kicker: str = "Порівняння",
                        number: Optional[int] = None, notes_text: str = "",
                        source_citation: str = "") -> None:
        """Слайд-таблиця. За потреби — поряд візуальний блок ≤30 %."""
        slide = self._new_slide()
        self._header(slide, title, kicker=kicker, number=number)
        table_w = (self.TXT_W + 0.15) if with_visual else (self.SLIDE_W - 2 * self.MARGIN)
        rows, cols = len(rows_data) + 1, len(headers)
        gt = slide.shapes.add_table(rows, cols, Inches(self.MARGIN),
                                    Inches(self.HEADER_H + 0.5),
                                    Inches(table_w), Inches(0.5 + 0.55 * rows))
        # ширина першої колонки трохи більша
        try:
            gt.table.columns[0].width = Inches(table_w * 0.34)
            for j in range(1, cols):
                gt.table.columns[j].width = Inches(table_w * 0.66 / (cols - 1))
        except Exception:
            pass
        self._style_table(gt.table, headers, rows_data)
        if with_visual:
            self._visual_panel(slide, image_path, visual_caption,
                               y=self.HEADER_H + 0.5, h=4.6)
        self._footer(slide, source_citation)
        self._add_notes(slide, notes_text, source_citation)

    def add_infographic_slide(self, title: str, steps: Sequence[Tuple[str, str]], *,
                            kicker: str = "Процес", number: Optional[int] = None,
                            notes_text: str = "", source_citation: str = "",
                            image_path: Optional[str] = None) -> None:
        """Інфографіка-процес із базових фігур: пронумеровані блоки + стрілки."""
        slide = self._new_slide()
        self._header(slide, title, kicker=kicker, number=number)
        n = len(steps)
        y0 = self.HEADER_H + 0.7
        # до 4 кроків — горизонтально; більше — два ряди
        if n <= 4:
            self._info_row(slide, steps, y0, height=3.9, cols=n)
        else:
            top = (n + 1) // 2
            self._info_row(slide, steps[:top], y0, height=1.9, cols=top)
            self._info_row(slide, steps[top:], y0 + 2.35, height=1.9, cols=n - top)
        self._footer(slide, source_citation)
        self._add_notes(slide, notes_text, source_citation)

    def _info_row(self, slide, steps, y, height, cols):
        gap = 0.35
        total_w = self.SLIDE_W - 2 * self.MARGIN
        bw = (total_w - gap * (cols - 1)) / cols
        palette = [Theme.ACCENT, Theme.ACCENT_2, Theme.WARN, Theme.DARK_ALT,
                   Theme.ACCENT, Theme.ACCENT_2]
        for i, (stitle, sdesc) in enumerate(steps):
            x = self.MARGIN + i * (bw + gap)
            col = palette[i % len(palette)]
            self._rect(slide, x, y, bw, height, Theme.WHITE,
                       shape=MSO_SHAPE.ROUNDED_RECTANGLE, line_color=Theme.HAIRLINE)
            self._rect(slide, x, y, bw, 0.6, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            self._rect(slide, x, y + 0.3, bw, 0.3, col)  # прибрати нижнє заокруглення
            # номер-кружечок
            self._rect(slide, x + 0.18, y + 0.14, 0.34, 0.34, Theme.WHITE, shape=MSO_SHAPE.OVAL)
            self._text(slide, x + 0.18, y + 0.14, 0.34, 0.34, str(i + 1),
                       size=13, color=col, bold=True, align=PP_ALIGN.CENTER,
                       anchor=MSO_ANCHOR.MIDDLE, autofit=False)
            self._text(slide, x + 0.6, y + 0.1, bw - 0.7, 0.42, stitle,
                       font=Theme.BODY_FONT, size=13, color=Theme.WHITE, bold=True,
                       anchor=MSO_ANCHOR.MIDDLE, autofit=False)
            self._text(slide, x + 0.18, y + 0.72, bw - 0.36, height - 0.82, sdesc,
                       font=Theme.BODY_FONT, size=12, color=Theme.TEXT_DARK,
                       line_spacing=1.08)
            # стрілка між блоками
            if i < len(steps) - 1:
                self._text(slide, x + bw - 0.02, y, gap + 0.04, height, "›",
                           size=26, color=Theme.MUTED, align=PP_ALIGN.CENTER,
                           anchor=MSO_ANCHOR.MIDDLE, autofit=False)

    def add_engagement_slide(self, question: str, *, prompt: str = "",
                            number: Optional[int] = None, notes_text: str = "",
                            source_citation: str = "", image_path: Optional[str] = None) -> None:
        """Слайд залучення: темне тло, велике питання по центру."""
        slide = self._new_slide(bg=Theme.DARK)
        self._rect(slide, 0, 0, self.SLIDE_W, 0.18, Theme.ACCENT)
        self._rect(slide, 0, self.SLIDE_H - 0.18, self.SLIDE_W, 0.18, Theme.ACCENT_2)
        self._text(slide, 1.2, 2.0, self.SLIDE_W - 2.4, 0.6, "ПАУЗА НА РОЗДУМИ",
                   font=Theme.BODY_FONT, size=15, color=Theme.ACCENT, bold=True,
                   align=PP_ALIGN.CENTER, autofit=False)
        self._text(slide, 1.2, 2.7, self.SLIDE_W - 2.4, 2.2, question,
                   font=Theme.TITLE_FONT, size=Theme.SZ_QUESTION, color=Theme.TEXT_LIGHT,
                   bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                   line_spacing=1.05, autofit=False)
        if prompt:
            self._text(slide, 2.0, 5.4, self.SLIDE_W - 4.0, 1.0, prompt,
                       font=Theme.BODY_FONT, size=17, color=RGBColor(0xB9, 0xCB, 0xDD),
                       align=PP_ALIGN.CENTER, autofit=False)
        self._add_notes(slide, notes_text, source_citation)

    def add_summary_slide(self, takeaways: Sequence[str], *,
                        title: str = "Ключові висновки", number: Optional[int] = None,
                        notes_text: str = "", source_citation: str = "",
                        image_path: Optional[str] = None) -> None:
        """Підсумковий слайд: пронумеровані тези-висновки."""
        slide = self._new_slide()
        self._header(slide, title, kicker="Підсумок", number=number)
        y = self.HEADER_H + 0.5
        for i, t in enumerate(takeaways):
            yy = y + i * 0.82
            self._rect(slide, self.MARGIN, yy, 0.58, 0.58, Theme.ACCENT_2, shape=MSO_SHAPE.OVAL)
            self._text(slide, self.MARGIN, yy, 0.58, 0.58, f"{i + 1}",
                       font=Theme.TITLE_FONT, size=17, color=Theme.WHITE, bold=True,
                       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, autofit=False)
            self._text(slide, self.MARGIN + 0.8, yy, self.SLIDE_W - 2.2, 0.72, t,
                       font=Theme.BODY_FONT, size=17, color=Theme.TEXT_DARK,
                       anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05, autofit=False)
        self._footer(slide, source_citation)
        self._add_notes(slide, notes_text, source_citation)

    # ------------------------------------------------------------------ #
    def save(self) -> None:
        """Зберігає файл із коректною обробкою помилок доступу."""
        try:
            self.prs.save(self.output_path)
            print(f"[OK] Презентацію збережено: {self.output_path}")
            print(f"[OK] Кількість слайдів: {len(self.prs.slides._sldIdLst)}")
        except (PermissionError, IOError) as err:
            alt = "presentation_fallback.pptx"
            print(f"[!] Не вдалося зберегти у '{self.output_path}': {err}")
            print(f"[!] Спроба зберегти як '{alt}' ...")
            try:
                self.prs.save(alt)
                print(f"[OK] Збережено як: {alt}")
            except (PermissionError, IOError) as err2:
                print(f"[X] Критична помилка збереження: {err2}")


# ==============================================================================
#  ЗБІРКА ПРЕЗЕНТАЦІЇ  (35 слайдів, повний україномовний контент)
# ==============================================================================
# Уніфіковані цитати-джерела
SRC_VIDEO = "Метамодель. «Штучний інтелект за 30 хвилин» (YouTube, 2026). youtube.com/watch?v=o3NXZdTnYYk"
SRC_MCP = "McCulloch, W. S., & Pitts, W. (1943). A logical calculus of the ideas immanent in nervous activity. Bulletin of Mathematical Biophysics, 5, 115–133."
SRC_TURING = "Turing, A. M. (1950). Computing Machinery and Intelligence. Mind, 59(236), 433–460."
SRC_ALEXNET = "Krizhevsky, A., Sutskever, I., & Hinton, G. (2012). ImageNet Classification with Deep CNNs. NeurIPS."
SRC_TRANSF = "Vaswani, A. et al. (2017). Attention Is All You Need. arXiv:1706.03762."
SRC_SUTTON = "Sutton, R. (2019). The Bitter Lesson. incompleteideas.net/IncIdeas/BitterLesson.html"
SRC_CRAWL = "Common Crawl Foundation. commoncrawl.org (дата звернення: 2026)."
SRC_VELLUM = "Порівняння галюцинацій і контекстних вікон LLM: Vectara HHEM leaderboard; документація OpenAI, Anthropic, Google DeepMind (2025–2026)."


def build_presentation() -> None:
    """Створює повну презентацію з ≥30 слайдів і зберігає її."""
    b = EducationalPresentationBuilder("Як_працює_штучний_інтелект.pptx")

    # --- 01 · ТИТУЛ ---
    b.add_title_slide(
        "Як працює штучний інтелект",
        "Від математичної моделі нейрона — до великих мовних моделей. "
        "Пояснюємо простими словами за принципом «під капотом».",
        presenter="Освітній курс · за мотивами відео «Штучний інтелект за 30 хвилин»",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Вітаю вас. Тема сьогоднішньої зустрічі звучить масштабно, але я обіцяю: за "
            "наступні пів години ми розберемо штучний інтелект так, щоб це стало зрозуміло "
            "кожному, навіть без технічної освіти. Ми підемо шляхом самої історії — від "
            "простої математичної моделі одного нейрона, описаної ще у 1943 році, до "
            "гігантських мовних моделей, якими сьогодні користуються мільярди людей.\n\n"
            "Чому це важливо саме зараз? Людина — це тварина, що створює інструменти. Але "
            "штучний інтелект — це вже не просто інструмент на кшталт молотка чи навіть "
            "атомної бомби. Це перший в історії мислячий агент, який ми створили. У 2026 "
            "році капітальні витрати в галузі ШІ, за оцінками, сягнуть близько трильйона "
            "доларів — це більше, ніж уся світова нафтогазова індустрія. ШІ стає новою "
            "інфраструктурою, як колись електрика чи інтернет. І тут працює просте правило: "
            "ті, хто навчиться працювати з ШІ, отримають перевагу; ті, хто не зможе або не "
            "захоче, ризикують залишитися позаду.\n\n"
            "Практична порада для старту, щоб ви одразу могли експериментувати паралельно з "
            "лекцією. Якщо у вас ще немає жодного ШІ-асистента під рукою, зробіть так:\n"
            "- Відкрийте будь-який із безкоштовних сервісів — ChatGPT, Google Gemini або Claude.\n"
            "- Зареєструйтесь через пошту або акаунт Google.\n"
            "- Знайдіть поле вводу внизу екрана — саме туди ми згодом вводитимемо «промпти».\n\n"
            "Наприкінці курсу ви розумітимете, що відбувається всередині, коли ви натискаєте "
            "«надіслати». А розуміти, як влаштована машина, — означає сидіти за кермом, а не "
            "опинитися під її колесами."
        ),
    )

    # --- 02 · ПРОГРАМА ---
    b.add_agenda_slide(
        [
            "ШІ як нова інфраструктура: чому це важливо",
            "Коротка історія: 1943 → 1956 → 2012 → 2017",
            "Два підходи: символічний ШІ та навчання на даних",
            "Будова нейромережі: нейрони, шари, ваги",
            "Мова машини: токени, вектори, контекст",
            "Механізм уваги та архітектура «Трансформер»",
            "Як народжується відповідь: логіти й температура",
            "Як навчають моделі: датасет, втрати, градієнт",
            "Чому ШІ галюцинує та як цьому зарадити",
            "Інфраструктура: GPU, дата-центри, вартість",
            "Відкриті та закриті моделі: запуск удома",
            "Майбутнє: AGI, моделі світу, гонка систем",
        ],
        number=2,
        source_citation=SRC_VIDEO,
        notes_text=(
            "Ось наша дорожня карта на сьогодні. Ми свідомо рухаємось від простого до "
            "складного і від історії до сучасності — так кожен новий термін спиратиметься "
            "на попередній.\n\n"
            "Спочатку відповімо на питання «навіщо». Потім здійснимо коротку історичну "
            "подорож, щоб зрозуміти, звідки взялися нейромережі. Далі — найважливіший блок: "
            "ми буквально розберемо нейромережу на деталі й побачимо, як вона перетворює "
            "картинку з котиком або ваше речення на набір чисел. Після цього розберемося, як "
            "ці моделі навчають — це окрема захоплива інженерія. Обов’язково поговоримо про "
            "галюцинації, бо це найпоширеніша практична проблема. І завершимо інфраструктурою "
            "та поглядом у майбутнє — на AGI, штучний інтелект загального призначення.\n\n"
            "Практична організаційна порада для тих, хто конспектує:\n"
            "- У застосунку для нотаток створіть розділи за цими дванадцятьма пунктами.\n"
            "- Навпроти складних термінів залишайте місце — ми повертатимемось до них.\n\n"
            "Не намагайтеся запам’ятати всі терміни одразу. Мета — побачити загальну логіку, "
            "а деталі осядуть самі, коли ви почнете користуватися ШІ усвідомлено."
        ),
    )

    # --- 03 · ШІ ЯК ІНФРАСТРУКТУРА ---
    b.add_concept_slide(
        "ШІ — це вже не інструмент, а мислячий агент",
        [
            "Людина — тварина, що створює інструменти; навіть атомна бомба лишається "
            "безсловесним знаряддям.",
            "Штучний інтелект — якісно нове: не пасивний інструмент, а агент, здатний "
            "обробляти інформацію та ухвалювати рішення.",
            "Це нова інфраструктура — як електрика чи інтернет; нею вже користуються "
            "мільярди людей щодня.",
            "Розуміти, як він працює, стає базовою грамотністю ХХІ століття.",
        ],
        lead="Уперше в історії людство створило не знаряддя, а агента.",
        kicker="Вступ · Навіщо це знати",
        number=3,
        visual_caption="КАДР З ВІДЕО · ~00:10 — «мислячий агент»",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Почнімо з фундаментальної ідеї, яка задає тон усій розмові. Протягом усієї "
            "історії людина створювала інструменти — від кам’яної сокири до реактивного "
            "двигуна. Навіть найстрашніший винахід, атомна бомба, залишається безсловесним "
            "знаряддям: вона нічого не вирішує сама, вона лише виконує закладену функцію.\n\n"
            "Штучний інтелект уперше ламає цю логіку. Це не просто інструмент, який чекає, "
            "поки ним скористаються. Це агент, який отримує інформацію, обробляє її та формує "
            "власну відповідь чи дію. Різниця принципова: молоток не порадить вам, куди "
            "забити цвях, а мовна модель порадить, посперечається й запропонує альтернативу.\n\n"
            "Саме тому доречно говорити про ШІ як про нову інфраструктуру. Згадайте, як "
            "електрика колись перестала бути дивом і стала тлом нашого життя — вона просто є "
            "в кожній розетці. З інтернетом сталося те саме. Зараз на наших очах "
            "інфраструктурою стає штучний інтелект: він вбудовується в пошук, у застосунки, у "
            "робочі процеси.\n\n"
            "Що з цим робити практично вже сьогодні? Почніть свідомо помічати ШІ навколо себе:\n"
            "- Стрічка рекомендацій YouTube чи TikTok — Налаштування - Ваші дані - Історія "
            "переглядів: саме ці дані живлять нейромережу, що добирає вам відео.\n"
            "- Розумний фільтр пошти — Gmail - Налаштування - Фільтри та заблоковані адреси: "
            "спам відсіює модель машинного навчання.\n\n"
            "Головна теза слайда: розуміння того, як влаштований ШІ, перестає бути хобі для "
            "інженерів і стає практичною навичкою для кожного."
        ),
    )

    # --- 04 · ГРАФІК: КАПІТАЛЬНІ ВИТРАТИ ---
    b.add_chart_slide(
        "Масштаб інвестицій у ШІ, 2026 (оцінка)",
        "bar",
        {
            "categories": ["ШІ (капвитрати)", "Світова нафтогазова\nгалузь"],
            "series": {"Трильйони доларів США": [1.0, 0.8]},
        },
        insight=[
            "≈ 1 трильйон доларів — оцінка капітальних витрат у ШІ на 2026 рік.",
            "Це перевищує сукупні капвитрати всієї світової нафтогазової галузі.",
            "Гроші такого масштабу — сигнал: ринок бачить у ШІ не моду, а нову "
            "інфраструктуру.",
        ],
        kicker="Дані · Економіка",
        number=4,
        source_citation="Галузеві оцінки капітальних витрат у ШІ, 2026; " + SRC_VIDEO,
        notes_text=(
            "Щоб слово «інфраструктура» не звучало абстрактно, подивімося на цифри — вони "
            "вражають навіть скептиків. За оцінками, у 2026 році капітальні витрати у сфері "
            "штучного інтелекту складуть близько одного трильйона доларів. Щоб відчути "
            "масштаб: це більше, ніж капітальні витрати всієї світової нафтогазової галузі "
            "разом узятої — індустрії, яка десятиліттями вважалася кровоносною системою "
            "світової економіки.\n\n"
            "Про що це говорить? Коли найбільші компанії й держави вкладають такі суми, вони "
            "роблять довгострокову ставку. Це не бульбашка одного застосунку — це будівництво "
            "фундаменту: дата-центрів, чипів, енергомереж. Приблизно так само колись "
            "будували залізниці й електростанції.\n\n"
            "Практичний висновок для звичайної людини чи невеликого бізнесу:\n"
            "- Не потрібно будувати власний дата-центр, щоб скористатися цією хвилею.\n"
            "- Достатньо орендувати доступ до моделей через API або хмару — Панель "
            "провайдера - Створити ресурс - Обрати модель/сервер - Підтвердити.\n\n"
            "Зверніть увагу на сам графік: дві колонки навмисно поставлені поруч, щоб "
            "співвідношення читалося миттєво. Коли будуєте власні презентації, тримайтеся "
            "цього принципу — одна ключова думка, одне порівняння, жодного зайвого шуму."
        ),
    )

    # --- 05 · ІСТОРІЯ: ВИТОКИ ---
    b.add_concept_slide(
        "Витоки: ровесник комп’ютерів",
        [
            "1943 — Маккалок і Піттс описують математичну модель нейрона: приймає сигнали, "
            "обробляє, і якщо сила достатня — передає далі.",
            "1950 — Алан Тюринг ставить питання «Чи може машина мислити?» і пропонує "
            "«тест Тюринга» як практичну перевірку.",
            "1956 — на Дартмутській конференції вперше звучить сам термін «штучний інтелект».",
            "Ця модель нейрона стала прообразом усіх сучасних нейромереж.",
        ],
        lead="Ідея ШІ така ж стара, як і самі комп’ютери.",
        kicker="Історія · 1943–1956",
        number=5,
        visual_caption="КАДР З ВІДЕО · ~01:05 — схема штучного нейрона",
        source_citation=SRC_MCP + " | " + SRC_TURING,
        notes_text=(
            "Багато хто вважає, що штучний інтелект — винахід останніх років. Насправді ідея "
            "ровесниця самих комп’ютерів. Повернімося у 1943 рік. Двоє вчених — Воррен "
            "Маккалок і Волтер Піттс — описали математичну модель нейрона головного мозку. "
            "Логіка проста й геніальна: нейрон отримує сигнали від інших нейронів, "
            "підсумовує їх, і якщо загальна сила перевищує певний поріг — передає сигнал "
            "далі. Якщо ні — мовчить. Саме ця модель, якій уже понад вісімдесят років, стала "
            "прообразом усіх сьогоднішніх нейромереж.\n\n"
            "Наступна віха — 1950 рік. Британський математик Алан Тюринг ставить питання, "
            "яке досі не втратило гостроти: «Чи може машина мислити?» І пропонує практичний "
            "критерій — знаменитий тест Тюринга: якщо людина в переписці не може відрізнити "
            "машину від людини, машину умовно можна вважати «мислячою».\n\n"
            "І, нарешті, 1956 рік, Дартмутська конференція, де вперше пролунав сам термін "
            "«штучний інтелект». Відтоді галузь отримала ім’я.\n\n"
            "Практична вправа, щоб історія ожила:\n"
            "- Спробуйте влаштувати власний «тест Тюринга» — задайте моделі складне питання й "
            "оцініть, наскільки відповідь схожа на людську.\n"
            "- Увімкніть режим міркувань, якщо він є: ChatGPT - меню вибору моделі - оберіть "
            "варіант із позначкою «reasoning» / «thinking».\n\n"
            "Запам’ятайте цю трійку дат — 1943, 1950, 1956. Це фундамент, на якому стоїть усе "
            "інше."
        ),
    )

    # --- 06 · ТАБЛИЦЯ: ДВА ПІДХОДИ ---
    b.add_table_slide(
        "Два підходи до штучного інтелекту",
        ["Критерій", "Символічний ШІ", "Навчання на даних"],
        [
            ["Хто задає правила", "Людина пише правила «якщо–то»", "Модель сама шукає закономірності"],
            ["Джерело знань", "Явно прописана логіка", "Приклади з датасету"],
            ["Приклад", "Ігрові боти, ранні експертні системи", "Розпізнавання образів, LLM"],
            ["Плюс", "Прозорість, передбачуваність", "Гнучкість, робота з винятками"],
            ["Мінус", "Не описати всі винятки світу", "Потрібно багато даних і обчислень"],
        ],
        kicker="Порівняння · Дві парадигми",
        number=6,
        visual_caption="КАДР З ВІДЕО · ~01:38 — «правила проти прикладів»",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Від самого початку в штучному інтелекті боролися два підходи, і розуміти "
            "різницю між ними критично важливо — тоді решта картини складеться сама.\n\n"
            "Перший підхід — символічний. Тут людина власноруч пише правила, а машина точно "
            "їх виконує. Це логіка «якщо–то»: якщо гравець у полі зору — атакуй; якщо почув "
            "шум — перевір; якщо здоров’я низьке — відступай. Саме так працювали ранні "
            "експертні системи вісімдесятих і донині працюють персонажі в багатьох іграх. "
            "Величезний плюс — прозорість: код можна відкрити, прочитати й зрозуміти, чому "
            "система вчинила саме так.\n\n"
            "Другий підхід — навчання на даних, те, що ми сьогодні називаємо машинним "
            "навчанням. Тут правила не прописують. Замість цього моделі показують безліч "
            "прикладів і змушують її саму знаходити закономірності. Чому цей підхід "
            "переміг? Бо в реальному світі винятків і нюансів більше, ніж правил. Неможливо "
            "вручну описати геть усе.\n\n"
            "Практична порада — обирайте підхід під задачу:\n"
            "- Для жорстких, чітких бізнес-правил (знижки, доступи) досі краще працює "
            "символічна логіка — Налаштування - Автоматизація - Правила.\n"
            "- Для «розмитих» задач (тексти, зображення, настрій клієнта) підходить лише "
            "навчання на даних.\n\n"
            "Символічний ШІ не помер — він живий і корисний. Але прорив останнього "
            "десятиліття зробив саме другий підхід."
        ),
    )

    # --- 07 · ЗИМИ ШІ ---
    b.add_concept_slide(
        "«Зими» штучного інтелекту",
        [
            "Підходу «навчання на даних» довго бракувало двох речей: обсягу даних і "
            "потужності комп’ютерів.",
            "Через це прогрес ішов хвилями: сплеск ентузіазму — потім розчарування й "
            "заморозка фінансування.",
            "Такі періоди застою назвали «зимами штучного інтелекту».",
            "Урок: правильна ідея може випереджати технічні можливості на десятиліття.",
        ],
        lead="Дві «зими» показали: ідея без обчислень не злітає.",
        kicker="Історія · Цикли розвитку",
        number=7,
        visual_caption="КАДР З ВІДЕО · ~02:03 — «зими ШІ»",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Історія штучного інтелекту — це не рівна висхідна лінія, а радше гірські лижі: "
            "різкі підйоми змінювалися болючими спусками. Причина в тому, що другому "
            "підходу — навчанню на даних — довгий час бракувало двох ключових інгредієнтів: "
            "самих даних і обчислювальної потужності.\n\n"
            "Уявіть геніального кухаря без продуктів і без плити. Рецепт правильний, але "
            "приготувати страву неможливо. Приблизно так почувалися дослідники ШІ у "
            "сімдесятих і кінці вісімдесятих. Через це розвиток ішов хвилями: періодично "
            "науковців охоплювало розчарування, інвестори втрачали віру, фінансування "
            "заморожувалося. Ці періоди застою отримали влучну назву — «зими штучного "
            "інтелекту».\n\n"
            "Чому це важливо для нас сьогодні? Це урок про терпіння й про природу прогресу. "
            "Правильна ідея цілком може випереджати свій час на десятиліття. Модель нейрона "
            "1943 року чекала майже сімдесят років, поки з’явилися дані й «залізо», здатні її "
            "оживити.\n\n"
            "Практичне спостереження, яке варто перенести на власні проєкти:\n"
            "- Якщо ваша ідея з ШІ поки не працює, питання часто не в ідеї, а в ресурсах — "
            "перевірте якість та обсяг даних насамперед.\n"
            "- Оцініть, чи вистачає обчислень: спробуйте меншу задачу або орендуйте потужніший "
            "сервер — Хмара - Сервери з GPU - Обрати конфігурацію.\n\n"
            "Наступний слайд — про момент, коли зима нарешті скінчилася і почалася «весна» ШІ."
        ),
    )

    # --- 08 · ПЕРЕЛОМ 2012 ---
    b.add_concept_slide(
        "Перелом 2012 року: AlexNet",
        [
            "Конкурс комп’ютерного зору ImageNet: команда Джефрі Гінтона показує мережу "
            "AlexNet.",
            "Її навчили лише на двох відеокартах NVIDIA GTX 580 і мільйоні зображень.",
            "AlexNet різко обійшла суперників — усі зрозуміли: нейромережі не іграшка, а "
            "інструмент.",
            "Головне відкриття: якість зростає при масштабуванні даних і обчислень.",
        ],
        lead="Момент, коли нейромережі перестали бути іграшкою.",
        kicker="Історія · 2012",
        number=8,
        visual_caption="КАДР З ВІДЕО · ~02:11 — конкурс ImageNet / AlexNet",
        source_citation=SRC_ALEXNET,
        notes_text=(
            "Тепер — про момент, який без перебільшення змінив хід історії технологій. Рік "
            "2012, конкурс комп’ютерного зору під назвою ImageNet, де алгоритми змагалися, "
            "хто точніше розпізнає об’єкти на мільйоні зображень. Команда вченого Джефрі "
            "Гінтона представила нейромережу AlexNet, навчену лише на двох споживчих "
            "відеокартах NVIDIA GTX 580.\n\n"
            "За сьогоднішніми мірками це смішні потужності — такі відеокарти стояли у "
            "звичайних геймерів. Але тоді AlexNet різко, з великим відривом обійшла всіх "
            "суперників. І сталося прозріння: усі зрозуміли, що нейромережі — більше не "
            "академічна іграшка, а робочий інструмент, який стає тим кращим, чим більше "
            "даних і обчислень ви йому даєте.\n\n"
            "Це і є ключова ідея слайда — масштабованість. До 2012 року здавалося, що "
            "прогрес упирається в геніальність алгоритму. Після 2012 стало ясно: часто "
            "достатньо взяти той самий підхід і дати йому більше даних та більше "
            "обчислювальної сили.\n\n"
            "Практичний зв’язок із сучасністю:\n"
            "- Та сама NVIDIA, чиї ігрові карти живили AlexNet, сьогодні майже монополіст "
            "на ринку ШІ-прискорювачів.\n"
            "- Якщо захочете самі навчити невелику модель, перевірте наявність GPU — "
            "Диспетчер завдань - Продуктивність - GPU, або в хмарі оберіть сервер із "
            "відеокартою.\n\n"
            "Саме з 2012 року нейромережі почали рости, як то кажуть, «на стероїдах»."
        ),
    )

    # --- 09 · ТРАНСФОРМЕР 2017 ---
    b.add_concept_slide(
        "2017: архітектура «Трансформер»",
        [
            "Архітектура — це те, як розташовані нейрони і які функції вони виконують.",
            "«Трансформер» різко спрощує навчання на великих даних і чудово масштабується.",
            "Саме він став прологом до великих мовних моделей (LLM) — ChatGPT, Gemini, Claude.",
            "Ключова ідея — механізм уваги (розберемо далі окремо).",
        ],
        lead="Архітектура, що народила еру ChatGPT.",
        kicker="Історія · 2017",
        number=9,
        visual_caption="КАДР З ВІДЕО · ~02:51 — схема «Трансформера»",
        source_citation=SRC_TRANSF,
        notes_text=(
            "Якщо 2012 рік довів, що масштаб працює, то 2017 рік дав інструмент, який "
            "дозволив масштабувати мовні моделі до нинішніх висот. Цього року з’явилася "
            "архітектура під назвою «Трансформер».\n\n"
            "Спершу домовимося про слово «архітектура». У світі нейромереж це те, як саме "
            "розташовані нейрони і яку функцію кожен з них виконує — образно кажучи, "
            "планування будинку: де несучі стіни, де вікна, як кімнати з’єднані між собою. "
            "Різні архітектури по-різному «думають».\n\n"
            "Чим особливий «Трансформер»? Він різко спростив навчання на великих обсягах "
            "тексту й виявився винятково добре масштабованим. Простіше кажучи: чим більше "
            "даних і обчислень ви в нього вливаєте, тим розумнішим він стає, і робить це "
            "ефективніше за попередні підходи. Саме «Трансформер» став прологом до великих "
            "мовних моделей — тих самих LLM, на яких побудовані ChatGPT, Gemini та Claude.\n\n"
            "Серце «Трансформера» — механізм уваги. Ми присвятимо йому окремий слайд, бо "
            "саме він дозволяє моделі розуміти, які слова в реченні пов’язані між собою. "
            "Назва статті 2017 року говорить сама за себе — «Attention Is All You Need», "
            "тобто «Увага — це все, що вам потрібно».\n\n"
            "Практичний зв’язок:\n"
            "- Майже кожен сучасний асистент, яким ви користуєтесь, усередині — «Трансформер».\n"
            "- Довжину «пам’яті» такої моделі задає контекстне вікно — про нього теж поговоримо."
        ),
    )

    # --- 10 · ГІРКИЙ УРОК ---
    b.add_concept_slide(
        "«Гіркий урок» Річарда Саттона (2019)",
        [
            "Висновок: у довгій перспективі перемагають найзагальніші методи, що ефективно "
            "використовують зростання обчислень.",
            "Спроби вручну «вшити» знання про світ програють масштабуванню.",
            "Приклад: компанія ABBYY понад 20 років вручну описувала правила мов для "
            "перекладу — і ця праця виявилася марною перед статистичними моделями.",
            "Формула Саттона стала неформальним маніфестом масштабування.",
        ],
        lead="Масштаб перемагає ручні правила. Майже завжди.",
        kicker="Ідея · 2019",
        number=10,
        visual_caption="КАДР З ВІДЕО · ~03:05 — «The Bitter Lesson»",
        source_citation=SRC_SUTTON,
        notes_text=(
            "У 2019 році вчений Річард Саттон опублікував коротке есе під назвою «Гіркий "
            "урок» — The Bitter Lesson. Текст невеликий, але його вплив на галузь величезний. "
            "Головна теза така: у довгій перспективі перемагають найзагальніші методи, які "
            "ефективно використовують зростання обчислювальних потужностей, а не ті, де ми "
            "намагаємося вручну вшити в машину наші людські знання про світ.\n\n"
            "Чому урок «гіркий»? Бо він принизливий для дослідників. Дуже спокусливо вкласти "
            "у модель свою експертизу, свої елегантні правила. Але історія раз за разом "
            "показує: груба сила обчислень і даних зрештою перемагає красиві ручні рішення.\n\n"
            "Найяскравіший приклад — російська компанія ABBYY. Понад два десятиліття вона "
            "намагалася побудувати експертну систему для перекладу: наймала цілі команди "
            "лінгвістів, які вручну описували правила мови, винятки, граматику. І вся ця "
            "колосальна праця виявилася марною, коли з’явилися статистичні моделі, навчені "
            "просто на величезних обсягах тексту.\n\n"
            "Практичний висновок для ваших власних завдань:\n"
            "- Перш ніж витрачати місяці на ручні правила, перевірте, чи не вирішить задачу "
            "готова модель «з коробки».\n"
            "- Спробуйте спочатку простий промпт до наявної LLM — часто цього достатньо: "
            "поле вводу - опишіть задачу простими словами - оцініть результат.\n\n"
            "Формула Саттона стала неформальним маніфестом: масштабуй дані, моделі й "
            "обчислення — і якість зросте."
        ),
    )

    # --- 11 · СИМВОЛІЧНИЙ ШІ ЖИВИЙ ---
    b.add_concept_slide(
        "Символічний ШІ досі живий",
        [
            "10–15 років тому ШІ найчастіше зустрічався саме в іграх.",
            "Персонажі й досі діють за простими інструкціями «якщо–то»: бачить гравця — "
            "атакує, чує шум — перевіряє, мало здоров’я — відступає.",
            "Плюс: поведінка передбачувана, код можна прочитати й зрозуміти.",
            "Мінус: у реальному світі винятків більше, ніж можна прописати вручну.",
        ],
        lead="Прості правила теж працюють — там, де світ передбачуваний.",
        kicker="Концепція · Символічний підхід",
        number=11,
        visual_caption="КАДР З ВІДЕО · ~03:57 — ігровий бот «якщо–то»",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Ми сказали, що масштабування перемогло, але це не означає, що символічний "
            "підхід помер. Навпаки — він живий і щодня працює навколо нас, просто ми цього "
            "не помічаємо. Найзнайоміший приклад — відеоігри.\n\n"
            "Ще десять-п’ятнадцять років тому, коли говорили про «штучний інтелект», "
            "найчастіше мали на увазі поведінку персонажів у іграх. І донині більшість "
            "ігрових ботів діють за простими, жорстко прописаними інструкціями. Логіка "
            "буквально така: якщо гравець потрапив у поле зору — атакуй; якщо почув сторонній "
            "шум — піди перевір; якщо здоров’я впало нижче порога — відступи й сховайся.\n\n"
            "Це і є класичний символічний ШІ: людина описує світ набором правил «якщо–то», а "
            "машина точно, до літери, виконує ці інструкції. У такого підходу є вагомі "
            "переваги. Поведінка передбачувана — ви завжди знаєте, як бот відреагує. Код "
            "прозорий — його можна відкрити, прочитати й зрозуміти кожне рішення.\n\n"
            "Але є й фундаментальне обмеження. У реальному, живому світі винятків і нюансів "
            "значно більше, ніж правил, які можна прописати вручну. Спробуйте описати "
            "правилами, як виглядає «котик» на будь-якому можливому фото — ви здастеся на "
            "тисячному винятку.\n\n"
            "Практичний висновок:\n"
            "- Використовуйте символічну логіку там, де світ передбачуваний і правил "
            "небагато — Автоматизація - Створити правило - Якщо [умова] - То [дія].\n"
            "- Там, де панує різноманіття (образи, мова, емоції), переходьте до навчання на "
            "прикладах.\n\n"
            "Саме про машинне навчання й піде мова далі."
        ),
    )

    # --- 12 · МАШИННЕ НАВЧАННЯ ---
    b.add_concept_slide(
        "Машинне навчання: вчимо на прикладах",
        [
            "Ідемо «від зворотного»: не пояснюємо машині устрій світу, а показуємо приклади.",
            "Спочатку обирають модель: нейромережа, лінійна формула або дерево рішень.",
            "Моделі дають задачу: «це котик, а це — ні». Вона відповідає, помиляється і "
            "підкручує налаштування.",
            "Так — багато-багато разів, аж поки не навчиться впізнавати котика в будь-якому "
            "вигляді.",
        ],
        lead="Нейромережа — це стиснутий досвід, а не набір інструкцій.",
        kicker="Концепція · Як вчиться машина",
        number=12,
        visual_caption="КАДР З ВІДЕО · ~04:50 — «котик / не котик»",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Отже, підійшли до серця сучасного ШІ — машинного навчання. Ідея тут "
            "перевертає звичну логіку. Замість того щоб пояснювати машині, як улаштований "
            "світ, ми йдемо від зворотного: показуємо їй приклади й змушуємо самостійно "
            "знайти закономірності.\n\n"
            "Як це відбувається на практиці? Спочатку обирають модель. Важливо: модель — це "
            "не обов’язково нейромережа. Це може бути й проста лінійна формула, і «дерево "
            "рішень». Різниця між ними — в устрої, але логіка навчання одна.\n\n"
            "Далі — сам процес. Моделі дають задачу з відповіддю: ось картинка, це котик; а "
            "ось інша — це не котик. Модель робить припущення, найчастіше спершу помиляється, "
            "після чого трохи підкручує свої внутрішні налаштування так, щоб наступного разу "
            "помилятися рідше. Потім ще раз, і ще, і так багато-багато разів. Зрештою вона "
            "навчається впізнавати котика в будь-якій позі, за будь-якого освітлення й навіть "
            "у мультяшному стилі.\n\n"
            "Найточніша метафора звучить так: якщо символічний ШІ — це інструкції в явному "
            "вигляді, то нейромережа — це стиснутий досвід. Вона не зберігає правила; вона "
            "зберігає узагальнення тисяч прикладів.\n\n"
            "Практична порада, як застосувати цю логіку самому:\n"
            "- Якщо навчаєте сервіс власним прикладам, дайте якомога різноманітніші зразки — "
            "Налаштування - Навчальні дані / Приклади - Додати.\n"
            "- Пам’ятайте правило: якість моделі майже завжди впирається в якість прикладів.\n\n"
            "Далі розберемо, з чого фізично складається нейромережа."
        ),
    )

    # --- 13 · ІНФОГРАФІКА: ШАРИ ---
    b.add_infographic_slide(
        "Будова нейромережі: шари нейронів",
        [
            ("Вхідний шар", "Отримує числову інформацію: пікселі картинки або токени тексту, "
                            "перетворені на числа."),
            ("Приховані шари", "Нейрони-«калькулятори» переобчислюють числа, шукаючи ознаки: "
                               "спершу прості, далі складніші."),
            ("Вихідний шар", "Видає результат: текст, зображення чи звук, породжений "
                             "нейромережею."),
        ],
        kicker="Структура · Шари",
        number=13,
        source_citation=SRC_VIDEO,
        notes_text=(
            "Тепер зазирнемо всередину нейромережі й побачимо її будову. Пам’ятаєте модель "
            "нейрона 1943 року? У сучасних мережах діє той самий принцип. Нейрони тут — це "
            "наче маленькі калькулятори, що перераховують немислиму кількість чисел.\n\n"
            "Ці нейрони об’єднані в шари, і саме послідовність шарів формує «конвеєр "
            "мислення». Перший шар — вхідний. Він отримує числову інформацію: якщо це "
            "картинка — числові значення пікселів, якщо текст — токени, перетворені на "
            "числа. Вхідний шар передає ці дані далі.\n\n"
            "Потім працюють приховані шари — їх може бути від кількох до сотень. Вони "
            "переобчислюють числа знову й знову, витягуючи дедалі складніші ознаки. Уявіть "
            "конвеєр: перший прихований шар помічає прості речі — краї, плями кольору; "
            "наступні збирають із них форми — вухо, ніс, лапу; ще глибші — цілий образ "
            "«котика». Кожен шар стоїть на плечах попереднього.\n\n"
            "Нарешті вихідний шар видає готовий результат — текст, зображення чи звук, "
            "породжений мережею.\n\n"
            "Практичне спостереження:\n"
            "- Коли чуєте «глибоке навчання» (deep learning), «глибина» — це і є кількість "
            "прихованих шарів.\n"
            "- Більше шарів — потенційно розумніша модель, але й важча: їй потрібно більше "
            "пам’яті — перевірте вимоги перед локальним запуском: Властивості моделі - "
            "Розмір / VRAM.\n\n"
            "А тепер подивимося на живому прикладі, як шари перетворюють картинку з котиком "
            "на рішення «так/ні»."
        ),
    )

    # --- 14 · ПІКСЕЛІ ТА ФУНКЦІЯ АКТИВАЦІЇ ---
    b.add_concept_slide(
        "Як мережа «бачить» котика",
        [
            "Картинку 280×280 пікселів можна пронумерувати: кожен піксель отримує числове "
            "ім’я, зрозуміле мережі.",
            "Кожен шар заповнює умовний «контрольний лист» ознак: є шерсть — таке значення, "
            "є вуса — таке.",
            "Якщо у фіналі досягнуто порогового значення — перед нами котик; ні — не котик.",
            "Це рішення «спрацювати чи мовчати» і називають функцією активації.",
        ],
        lead="Слова замінюємо числами — і мережа починає «бачити».",
        kicker="Концепція · Ознаки",
        number=14,
        visual_caption="КАДР З ВІДЕО · ~05:44 — пікселі 280×280",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Розберемо на живому прикладі, як нейромережа перетворює зображення на рішення. "
            "Візьмемо картинку з котиком розміром, скажімо, 280 на 280 пікселів. Ми можемо "
            "пронумерувати кожен піксель — дати йому числове ім’я, яке зрозуміє мережа. І ось "
            "вона вже здатна працювати з цим пікселем: порівнювати з сусідніми, пов’язувати "
            "з об’єктами, що теж мають цифрові імена.\n\n"
            "Тепер завдання — зрозуміти, котик на картинці чи ні. Кожен шар нейронів немов "
            "заповнює умовний контрольний лист, але замість слів використовує числові "
            "значення пікселів. Є шерсть — це таке-то число. Є вуса — ось таке. Крок за "
            "кроком мережа накопичує «докази».\n\n"
            "Якщо у фіналі сумарне значення досягає певного порога — перед нами котик. Якщо "
            "не досягає — не котик. Оцей механізм «спрацювати чи промовчати» називають "
            "функцією активації, і вона працює для кожного шару. Уявіть турнікет: набрали "
            "достатньо жетонів — прохід відкривається; ні — стоїте на місці.\n\n"
            "Практичний висновок для розуміння генерації зображень:\n"
            "- Коли ви пишете запит генератору картинок, він робить зворотну операцію — "
            "від опису до пікселів.\n"
            "- Щоб отримати кращий результат, задавайте деталі явно — поле промпту - опишіть "
            "об’єкт, стиль, освітлення, ракурс - згенерувати.\n\n"
            "Ми розглянули лише наявність ознак. Але не всі ознаки однаково важливі — про це "
            "наступний слайд, де з’являться ваги."
        ),
    )

    # --- 15 · ВАГИ ТА ЗМІЩЕННЯ ---
    b.add_concept_slide(
        "Ваги та зміщення: що важливіше",
        [
            "Вуса є і в кота, і в миші — така ознака менш важлива, ніж, скажімо, форма кігтів.",
            "Тому нейрон множить значення ознак на «коефіцієнти важливості» — їх називають "
            "вагами.",
            "Уяви трафарет: нейрон прикладає його до картинки й питає «схоже?». Схоже — "
            "активується, ні — мовчить.",
            "Щоб урахувати різноманіття хвостів і вух, додають поправку — зміщення (bias).",
        ],
        lead="Ваги кажуть, на що дивитися; bias — коли робити виняток.",
        kicker="Концепція · Параметри",
        number=15,
        visual_caption="КАДР З ВІДЕО · ~06:46 — «трафарет» нейрона",
        source_citation=SRC_VIDEO,
        notes_text=(
            "На попередньому слайді ми рахували ознаки так, ніби всі вони рівноцінні. Але це "
            "не так, і саме тут криється магія нейромереж. Погляньмо уважніше.\n\n"
            "Вуса є і в кота, і в собаки, і в миші. Отже, сама лише наявність вусів — не дуже "
            "надійна ознака. А от кігті певної форми важать значно більше. Як мережа "
            "враховує цю різницю? Кожен нейрон множить значення ознаки на коефіцієнт "
            "важливості й підсумовує результати. Ці коефіцієнти важливості називають вагами.\n\n"
            "Найкраща метафора: у кожного нейрона є свій трафарет. Він прикладає його до "
            "картинки й питає — схоже чи ні? Якщо схоже, нейрон активується; якщо ні — мовчить. "
            "Ваги — це, по суті, форма трафарета: на що саме цей нейрон навчився звертати "
            "увагу.\n\n"
            "Але є ще нюанс. Хвости, вуха й кігті бувають дуже різні. Їхні числові значення "
            "можуть трохи не дотягувати до потрібного порога — або, навпаки, перескакувати "
            "його. Тому потрібна поправка. Її називають зміщенням, або англійською bias. "
            "Зміщення дозволяє нейрону бути трохи поблажливішим чи, навпаки, суворішим.\n\n"
            "Практична аналогія, яка стане в пригоді далі:\n"
            "- Ваги й зміщення — це і є те, що модель «підкручує» під час навчання.\n"
            "- Коли ви робите тонке налаштування моделі під свою галузь, змінюються саме ці "
            "числа — Проєкт - Fine-tuning - Завантажити приклади - Запустити.\n\n"
            "Разом ваги й зміщення утворюють параметри мережі. Про їхній масштаб — далі."
        ),
    )

    # --- 16 · ПАРАМЕТРИ, МАТРИЦІ, ТЕНЗОРИ ---
    b.add_concept_slide(
        "Параметри, матриці та тензори",
        [
            "Параметри нейромережі — це всі її ваги й зміщення разом.",
            "Їх зберігають у матрицях і тензорах. Матриця — це буквально таблиця чисел, як в "
            "Excel.",
            "Тензор — масив чисел будь-якої розмірності, по суті «пачка матриць».",
            "Обсяг гігантський: наприклад, модель DeepSeek R1 важить близько 700 ГБ, і "
            "більшість цього — параметри.",
        ],
        lead="Один нейрон — трафарет; шар — ціла пачка трафаретів.",
        kicker="Концепція · Масштаб",
        number=16,
        visual_caption="КАДР З ВІДЕО · ~07:16 — матриці й тензори",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Ми познайомилися з вагами й зміщеннями. Тепер дамо цьому загальну назву й "
            "відчуємо масштаб. Усі ваги й усі зміщення мережі разом називають параметрами. "
            "Це, фактично, вся «пам’ять» і весь «досвід» моделі, стиснуті в числа.\n\n"
            "Де їх зберігають? У матрицях і тензорах — і ці слова звучать страшніше, ніж є "
            "насправді. Матриця — це буквально таблиця чисел, точнісінько як аркуш Excel: "
            "рядки й стовпці, у кожній клітинці число. Тензор — це масив чисел будь-якої "
            "розмірності, по суті пачка матриць, складених разом. Якщо один нейрон — це один "
            "трафарет, то шар — це вже ціла пачка трафаретів, а вся мережа — величезна "
            "бібліотека таких пачок.\n\n"
            "Наскільки все це велике? Цифри вражають. Модель DeepSeek R1 важить близько "
            "семисот гігабайтів, і переважна частина цього обсягу — саме параметри. Для "
            "порівняння: це сотні тисяч звичайних книжок, тільки замість літер — числа.\n\n"
            "Практичний висновок, дуже корисний на практиці:\n"
            "- Число параметрів (наприклад, «7B», «70B» — тобто мільярди) підказує, "
            "наскільки модель важка й потужна.\n"
            "- Перед завантаженням локальної моделі оцініть, чи влізе вона в пам’ять: "
            "Картка моделі - параметри / розмір файлу - порівняйте з вашим обсягом RAM та "
            "VRAM.\n"
            "- Менші моделі (2B–8B) працюють навіть на домашньому ноутбуці, великі — "
            "вимагають серверів.\n\n"
            "З картинками розібралися. Але як мережа працює зі словами? Про це — далі."
        ),
    )

    # --- 17 · ТОКЕНИ ТА ЕМБЕДИНГИ ---
    b.add_concept_slide(
        "Мова машини: токени та вектори",
        [
            "Мовна модель не бачить слів — вона ріже текст на шматочки, які звуться токенами "
            "(слово, частина слова або навіть символ).",
            "Кожен токен перетворюється на набір чисел — вектор, або ембединг.",
            "Замість речення всередині моделі з’являється стовпчик чисел.",
            "Вектор — зручна форма пам’яті: зберігає одразу стиль, зміст, роль у фразі та "
            "натяк на тему.",
        ],
        lead="Речення → токени → числа, з якими вміє працювати мережа.",
        kicker="Концепція · Токенізація",
        number=17,
        visual_caption="КАДР З ВІДЕО · ~07:47 — токени й ембединги",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Ми побачили, що з картинкою все зрозуміло: у кожного пікселя є числове значення. "
            "А як бути зі словами? Адже мережа розуміє лише числа. Тут криється один із "
            "найважливіших механізмів мовних моделей.\n\n"
            "Річ у тім, що мовна модель не бачить текст як слова. Вона ріже його на "
            "шматочки — токени. Іноді токен — це ціле слово, іноді частина слова, а іноді "
            "навіть один символ. Наприклад, довге слово може розпастися на два-три токени.\n\n"
            "А далі — майже те саме, що з котиком. Кожен токен перетворюється на набір "
            "чисел, який називають вектором, або ембедингом. Тобто замість речення на кшталт "
            "«Як переконати кота платити за оренду» всередині моделі з’являється довгий "
            "стовпчик чисел.\n\n"
            "Навіщо так складно? Бо вектор — це надзвичайно зручна форма пам’яті. Один "
            "вектор може одночасно зберігати безліч ознак слова: його стиль, зміст, "
            "граматичну роль у фразі, натяк на тему. Уявіть, що кожне слово отримує "
            "багатовимірну «картку характеристик», і слова зі схожим змістом опиняються "
            "поряд у цьому числовому просторі.\n\n"
            "Практична порада, що заощадить вам гроші й час:\n"
            "- Оплата за API рахується саме в токенах, а не в словах — 1000 токенів "
            "приблизно дорівнює 750 словам англійською.\n"
            "- Хочете оцінити вартість запиту заздалегідь? Скористайтеся токенайзером: "
            "сайт провайдера - Tokenizer / Playground - вставте текст - подивіться "
            "кількість токенів.\n\n"
            "Але самих токенів замало — моделі важливо знати ще й їхній порядок. Про це далі."
        ),
    )

    # --- 18 · ПОЗИЦІЯ ТА КОНТЕКСТ ---
    b.add_concept_slide(
        "Порядок слів і контекстне вікно",
        [
            "Токени самі не кажуть, яке слово було раніше — тож до кожного додають вектор "
            "позиції (позиційні ембединги).",
            "Приклад «поганий лук»: лук — це зброя, овоч чи «образ»? Модель визначає зміст із "
            "діалогу.",
            "Розмір доступного контексту звуть контекстним вікном.",
            "Раніше 8–32 тис. токенів вважалося багато; у нових Gemini воно перевищує "
            "мільйон.",
        ],
        lead="Контекстне вікно — це те, що модель «тримає перед очима».",
        kicker="Концепція · Контекст",
        number=18,
        visual_caption="КАДР З ВІДЕО · ~08:43 — приклад «лук»",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Уявіть, що ви отримали слова речення в довільному порядку. Зрозуміти зміст буде "
            "важко, адже порядок вирішує все: «собака вкусила людину» й «людина вкусила "
            "собаку» складаються з однакових слів. Токени самі по собі не кажуть, яке слово "
            "стоїть раніше, а яке пізніше. Тому до кожного токена додають ще один вектор — з "
            "указанням його позиції в тексті. Це називають позиційними ембедингами.\n\n"
            "Тепер про зміст. Розгляньмо приклад із діалогу: «Чому в мене поганий лук?» Слово "
            "«лук» може означати зброю, овоч або, у сучасному сленгу, зовнішній вигляд, "
            "образ. Шар за шаром модель проганяє слово «лук» через матриці й з контексту "
            "діалогу розуміє, про що йдеться. Одне слово — три сенси, і лише контекст "
            "підказує правильний.\n\n"
            "Тут з’являється ключове поняття — контекстне вікно. Це обсяг тексту, який модель "
            "здатна тримати «перед очима» просто зараз. Чим воно більше, тим більше токенів "
            "модель враховує одночасно. У перших мовних моделях великим вважалося вікно від "
            "восьми до тридцяти двох тисяч токенів. А в останніх версіях Gemini воно "
            "перевищує мільйон токенів — це вже цілі книжки.\n\n"
            "Практична порада, як цим користуватися:\n"
            "- Для аналізу довгого документа обирайте модель із великим вікном — Вибір "
            "моделі - подивіться параметр «context window».\n"
            "- Якщо модель «забуває» початок довгої розмови, почніть новий чат або коротко "
            "повторіть ключові умови — вікно переповнилося.\n\n"
            "Але навіть мільйона токенів іноді замало. Як модель вирішує, на що дивитися "
            "уважно? Це робота механізму уваги — наступний слайд."
        ),
    )

    # --- 19 · ГРАФІК: КОНТЕКСТНЕ ВІКНО ---
    b.add_chart_slide(
        "Зростання контекстного вікна (токенів)",
        "barh",
        {
            "categories": ["Gemini (перші версії)", "GPT-4", "Claude 3", "GPT-5",
                           "Gemini 3.x Pro", "Claude Fable 5"],
            "series": {"Тисяч токенів": [32, 8, 200, 1000, 1000, 1000]},
        },
        insight=[
            "За кілька років вікно зросло від десятків тисяч до мільйона токенів.",
            "Довгий контекст дозволяє «згодовувати» цілі документи й тримати історію "
            "діалогу.",
            "Це основа для агентів, що читають, порівнюють і планують.",
        ],
        kicker="Дані · Контекст",
        number=19,
        source_citation=SRC_VELLUM,
        notes_text=(
            "Щоб зростання контекстного вікна не залишалося абстракцією, погляньмо на "
            "конкретні числа. Ще нещодавно перші версії Gemini мали вікно близько тридцяти "
            "двох тисяч токенів, а GPT-4 стартував із восьми тисяч. Сьогодні GPT-5, нові "
            "Gemini та Claude оперують вікнами близько мільйона токенів. Це стрибок у "
            "десятки й сотні разів за кілька років.\n\n"
            "Що це означає на практиці? Мільйон токенів — це приблизно кілька товстих книжок, "
            "які модель здатна тримати «перед очима» одночасно. Саме довгий контекст робить "
            "можливими справжніх агентів: асистента, який читає стосорінковий договір, тримає "
            "в пам’яті всю історію вашого діалогу, порівнює документи й не забуває інструкцію "
            "на другій хвилині розмови.\n\n"
            "Але велике вікно — це не безкоштовно, і про це важливо пам’ятати. Чим більше "
            "токенів у контексті, тим більше пам’яті й обчислень витрачає система, а отже, "
            "тим дорожчий запит.\n\n"
            "Практичні поради:\n"
            "- Не завантажуйте в контекст усе поспіль — давайте лише релевантні фрагменти.\n"
            "- Якщо працюєте через API, стежте за лічильником вхідних токенів — Панель - "
            "Usage / Використання - Tokens.\n"
            "- Для дуже великих баз знань замість «запхати все у вікно» використовуйте пошук "
            "по документах (RAG) — про це поговоримо згодом.\n\n"
            "Зверніть увагу: на графіку я свідомо використав горизонтальні смуги — коли назв "
            "багато й вони довгі, так підписи читаються краще, ніж на вертикальних стовпцях."
        ),
    )

    # --- 20 · ІНФОГРАФІКА: ШЛЯХ У ТРАНСФОРМЕРІ ---
    b.add_infographic_slide(
        "Шлях запиту всередині «Трансформера»",
        [
            ("Механізм уваги", "Вирішує, на які токени дивитися пильно, а які майже "
                               "ігнорувати. Він багатоголовий: кожна «голова» ловить свій тип "
                               "зв’язків."),
            ("MLP", "Багатошаровий перцептрон — міні-мережа, що підсилює сильні зв’язки й "
                    "послаблює слабкі."),
            ("Залишкові зв’язки\nта нормалізація", "Зберігають головну інформацію й не дають "
                                                    "«міркуванням» розсипатися на фрагменти."),
            ("Вихідний шар", "Оцінює кожне можливе продовження токена — далі обирається "
                             "найімовірніше."),
        ],
        kicker="Процес · Архітектура",
        number=20,
        source_citation=SRC_TRANSF,
        notes_text=(
            "Тепер зберемо все докупи й простежимо, який шлях проходить ваш запит усередині "
            "«Трансформера». Це чотири ключові станції одного конвеєра.\n\n"
            "Перша станція — механізм уваги. Він у кожному шарі вирішує, на які токени "
            "контексту дивитися пильно, а які можна майже ігнорувати. Уявіть читача, який "
            "виділяє маркером головні слова в тексті. Важливо, що цей механізм "
            "багатоголовий: кожна «голова» уваги чіпляється за свій тип зв’язків — одна "
            "стежить за граматикою, інша за змістом, третя за порядком слів.\n\n"
            "Друга станція — MLP, багатошаровий перцептрон. Це просто міні-мережа "
            "всередині великої. Вона допомагає зберігати сильні зв’язки між шарами й "
            "послаблювати слабкі. Згадайте приклад із «луком»: навіть коли з контексту "
            "зрозуміло, що йдеться про овоч, модель лишає крихітну ймовірність, що це "
            "зброя. MLP допомагає цю зайву ймовірність зменшити.\n\n"
            "Третя станція — залишкові зв’язки й нормалізація. Вони зберігають основну "
            "інформацію й не дають мережі кидатися в крайнощі, щоб її «міркування» не "
            "розсипалися на незрозумілі уламки. Це наче страхувальний трос.\n\n"
            "Четверта станція — вихідний шар, де мережа оцінює кожне можливе продовження. "
            "Про те, як саме робиться фінальний вибір, — наступний слайд.\n\n"
            "Практичний висновок:\n"
            "- Вам не треба керувати цими станціями вручну — вони працюють автоматично.\n"
            "- Але розуміння, що модель зважує ймовірності, а не «знає істину», убереже вас "
            "від сліпої довіри до відповідей."
        ),
    )

    # --- 21 · ЛОГІТИ ТА ТЕМПЕРАТУРА ---
    b.add_concept_slide(
        "Як обирається наступне слово",
        [
            "На виході мережа дає оцінку кожному можливому продовженню — це логіти.",
            "Температура керує «сміливістю»: низька — передбачувано й сухо; висока — "
            "різноманітно, але з ризиком помилок.",
            "Обмежувачі вибору: Top-K (бери з N найімовірніших) і Top-P (бери з набору із "
            "заданою сумарною ймовірністю).",
            "Термін «температура» прийшов із термодинаміки, де вона задає хаотичність "
            "системи.",
        ],
        lead="Модель не «знає» відповідь — вона обирає найімовірніше продовження.",
        kicker="Концепція · Семплінг",
        number=21,
        visual_caption="КАДР З ВІДЕО · ~10:40 — температура моделі",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Ми дійшли до фінального кроку генерації — моменту, коли модель обирає наступне "
            "слово. На вихідному шарі мережа буквально дає оцінку кожному можливому "
            "продовженню токена. Ці оцінки називають логітами. За ними модель і обирає "
            "найімовірніше продовження. Ключова думка: модель не «знає» правильну "
            "відповідь — вона обирає статистично найімовірніше продовження.\n\n"
            "Але вибір залежить від кількох налаштувань, і найважливіше з них — температура. "
            "Температура відповідає за міру сміливості моделі. Якщо вона низька, модель іде "
            "найімовірнішим шляхом — сухо, точно й передбачувано. Якщо висока — частіше "
            "обирає менш очевидні варіанти: відповіді стають різноманітнішими й "
            "креативнішими, але росте й ризик помилок. Цікаво, що сам термін прийшов із "
            "термодинаміки, де температура задає хаотичність системи.\n\n"
            "Є ще два обмежувачі вибору. Top-K каже моделі: «бери лише з певної кількості "
            "найімовірніших варіантів». Top-P каже інакше: «бери з найменшого набору "
            "варіантів, чия сумарна ймовірність складає стільки-то відсотків».\n\n"
            "А тепер — найкорисніше, практика. Ці параметри реально можна крутити:\n"
            "- Через ігровий майданчик API: Playground - панель Parameters праворуч - "
            "повзунок Temperature - Top P.\n"
            "  - Для точних задач (код, факти, юридичний текст) ставте температуру низькою, "
            "близько 0,2.\n"
            "  - Для мозкового штурму й креативу підіймайте до 0,8–1,0.\n"
            "- У звичайному чаті прямого повзунка немає, але ефект досяжний словами: додайте "
            "в промпт «відповідай стисло й точно» або «запропонуй сміливі нестандартні "
            "ідеї».\n\n"
            "Тепер ви розумієте весь шлях: від запиту до відповіді. Використання вже "
            "навченої мережі має назву — інференс. Про нього далі."
        ),
    )

    # --- 22 · ІНФЕРЕНС ---
    b.add_concept_slide(
        "Інференс: коли модель відповідає",
        [
            "Інференс — це використання вже навченої нейромережі для відповіді на запит.",
            "Це не лише математика всередині моделі: щоб вона відповідала, потрібен сервер.",
            "Чим потужніша модель і чим більше користувачів — тим серйозніші вимоги до "
            "інфраструктури.",
            "Хочете власний сайт, бота чи агента — вам теж потрібна інфраструктура.",
        ],
        lead="Навчання — це створення мозку; інференс — його щоденна робота.",
        kicker="Концепція · Робота моделі",
        number=22,
        visual_caption="КАДР З ВІДЕО · ~11:31 — інференс і сервер",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Ми простежили весь шлях, який проходить навчена нейромережа від запиту до "
            "відповіді. У цього процесу є спеціальна назва — інференс, від англійського "
            "inference. Запам’ятайте це слово: коли ви щось питаєте в ChatGPT і отримуєте "
            "відповідь, у цей момент відбувається інференс.\n\n"
            "Але інференс — це не лише математика всередині моделі. Щоб модель фізично "
            "відповіла, потрібен сервер, який виконає всі ці обчислення. І тут працює просте "
            "правило: чим потужніша модель і чим більше користувачів звертається до неї "
            "одночасно, тим серйозніші вимоги до інфраструктури. Один запит — дрібниця. "
            "Мільярд запитів на тиждень — колосальне навантаження.\n\n"
            "Проведімо межу, яка знадобиться далі. Навчання — це разове, дороге створення "
            "«мозку» моделі. Інференс — це щоденна, повторювана робота цього мозку на "
            "мільйони людей. Саме інференс, як ми побачимо, створює найбільші поточні "
            "витрати.\n\n"
            "Практична частина — а що, як вам самим потрібна інфраструктура? Скажімо, ви "
            "хочете запустити власний сайт, чат-бота чи розгорнути ШІ-агента. Тоді вам "
            "знадобиться сервер. Найпростіший шлях — оренда:\n"
            "- Орендуйте VPS — віддалений комп’ютер у дата-центрі, доступний з інтернету "
            "цілодобово.\n"
            "- Типовий порядок дій у будь-якого хмарного провайдера: Реєстрація - Панель "
            "керування - Створити сервер - Оберіть образ системи (Ubuntu) - Оберіть "
            "конфігурацію (CPU/RAM, за потреби GPU) - Створити.\n"
            "- Для ШІ-задач обирайте тариф із відеокартою — Тип сервера - з GPU.\n\n"
            "Далі повернемося до того, як ці моделі взагалі навчають — це окрема захоплива "
            "історія."
        ),
    )

    # --- 23 · ІНФОГРАФІКА: ДАТАСЕТ ---
    b.add_infographic_slide(
        "Підготовка даних для навчання",
        [
            ("Збір датасету", "Гігантські масиви тексту: книги, статті, сайти, код, діалоги. "
                              "Якість моделі впирається в якість даних."),
            ("Очищення", "Прибирають сміття, биті символи, повтори й персональні дані. Потім "
                         "текст ріжуть на токени."),
            ("Поділ на 3 частини", "Тренувальні (навчання), валідаційні (проміжна перевірка), "
                                    "тестові (фінальний іспит)."),
        ],
        kicker="Процес · Дані",
        number=23,
        source_citation=SRC_CRAWL,
        notes_text=(
            "Ми багато говорили про навчання «на прикладах». Тепер розберемо, звідки ці "
            "приклади беруться і як їх готують. Насамперед потрібен датасет — набір "
            "прикладів, на яких вчиться модель. Це надзвичайно важливо: якість моделі майже "
            "завжди впирається в якість навчальних даних. Датасет — це єдиний світ, який "
            "бачила модель. Іншого світу в неї немає. Якщо в даних була упередженість чи "
            "помилки — вони перейдуть у модель.\n\n"
            "Для мовної моделі на кшталт ChatGPT датасети — це гігантські масиви тексту: "
            "книжки, статті, сайти, програмний код, діалоги, урочисті промови й панічні "
            "чутки — усе розмаїття людської писемності. Генератори картинок натомість "
            "навчаються на зображеннях, генератори музики — на звуках.\n\n"
            "Перший крок — очищення. Датасет вичищають від сміття, битих символів, повторів, "
            "персональних даних. Потім текст ріжуть на токени, як ми вже обговорювали.\n\n"
            "Далі важливий момент — датасет ділять на три частини. На тренувальних даних "
            "модель навчається. На валідаційних проходить проміжну перевірку — їх можна "
            "використовувати багато разів. На тестових складає фінальний іспит, і їх краще "
            "не використовувати повторно. Чому? Якщо все змішати, модель частково побачить "
            "відповіді заздалегідь. Вона здаватиметься геніальною, а насправді просто "
            "запам’ятає шматки тексту — як студент, який зазубрив відповіді до тесту, не "
            "зрозумівши суті.\n\n"
            "Практична порада, якщо навчаєте власну модель:\n"
            "- Ніколи не змішуйте тестові дані з тренувальними — Датасет - Split - тримайте "
            "окремі файли train / validation / test.\n"
            "- Витратьте більшість часу на очищення даних, а не на налаштування моделі — це "
            "дає найбільший приріст якості."
        ),
    )

    # --- 24 · ІНФОГРАФІКА: ЦИКЛ НАВЧАННЯ ---
    b.add_infographic_slide(
        "Цикл навчання: як мережа стає розумнішою",
        [
            ("Прямий прохід", "Модель проганяє запит крізь усі шари й намагається вгадати "
                              "наступний токен."),
            ("Функція втрат", "Порівнює відповідь із правильною й рахує, наскільки "
                              "промахнулася (loss)."),
            ("Зворотний прохід", "Помилку «розносять» назад по всіх шарах — це backpropagation."),
            ("Градієнтний спуск", "Обчислює, як саме змінити ваги, щоб втрати зменшити. І так "
                                  "мільярди разів."),
        ],
        kicker="Процес · Навчання",
        number=24,
        source_citation=SRC_VIDEO,
        notes_text=(
            "Дані готові — час навчати. Цей процес схожий на замкнене коло, яке "
            "повторюється мільярди разів, і складається з чотирьох кроків.\n\n"
            "Крок перший — прямий прохід. Модель бере приклад, проганяє його крізь усі свої "
            "шари й намагається вгадати наступний токен. Це рівно те саме, що відбувається "
            "під час інференсу.\n\n"
            "Крок другий — функція втрат. Модель порівнює свій результат із правильною "
            "відповіддю й обчислює, наскільки сильно вона промахнулася. Цю міру помилки "
            "називають функцією втрат, або loss. З першого разу жодна модель не дає "
            "правильної відповіді — і це нормально.\n\n"
            "Крок третій — зворотний прохід, англійською backpropagation. Після обчислення "
            "помилки буквально всі параметри переглядають, рухаючись шарами у зворотному "
            "напрямку — від виходу до входу. Модель ніби питає кожен нейрон: «наскільки саме "
            "ти винен у цій помилці?»\n\n"
            "Крок четвертий — градієнтний спуск. Це функція, яка обчислює корегувальні "
            "матриці для кожного шару: грубо кажучи, призначає штраф і підказує, у який бік "
            "змінити ваги, щоб зменшити втрати. Уявіть, що ви спускаєтеся з гори в тумані, "
            "намацуючи ногою найкрутіший схил униз — це і є градієнтний спуск.\n\n"
            "Потім усе повторюється — ще раз, і ще, мільярди разів. Важлива й дещо "
            "несподівана думка: модель не прагне до сенсу. Вона прагне лише мінімізувати "
            "функцію втрат. Сенс виникає як побічний ефект.\n\n"
            "Практичний висновок:\n"
            "- Саме тому навчання таке дороге — це мільярди повторів цього циклу.\n"
            "- Для більшості завдань вам не треба навчати з нуля: беріть готову модель і "
            "лише трохи донавчайте — Проєкт - Fine-tuning."
        ),
    )

    # --- 25 · ШВИДКІСТЬ, БАТЧІ, ЕПОХИ ---
    b.add_concept_slide(
        "Швидкість навчання, батчі та епохи",
        [
            "Швидкість навчання = розмір кроку: наскільки сильно змінюються параметри за "
            "прохід.",
            "Крок завеликий — модель «скаче» й втрачає стійкість; замалий — повзе й вчиться "
            "вічність.",
            "Тому навчання — завжди компроміс між швидкістю та стабільністю.",
            "Приклади обробляють не поштучно, а пачками — батчами; один прохід по всьому "
            "датасету звуть епохою.",
        ],
        lead="Правильний крок — золота середина між хаосом і равликом.",
        kicker="Концепція · Гіперпараметри",
        number=25,
        visual_caption="КАДР З ВІДЕО · ~15:11 — розмір кроку навчання",
        source_citation=SRC_VIDEO,
        notes_text=(
            "У циклі навчання є тонке місце, від якого залежить успіх усього процесу, — "
            "швидкість навчання. Вона визначається величиною кроку: тобто тим, наскільки "
            "сильно змінюються параметри за кожен новий прохід.\n\n"
            "Тут важливий баланс, і його легко відчути на аналогії. Якщо крок завеликий, "
            "модель скаче туди-сюди й втрачає стійкість — уявіть людину, яка намагається "
            "спуститися сходами, стрибаючи через п’ять сходинок за раз. Якщо крок замалий — "
            "модель повзе й вчиться майже вічність, обережно переставляючи ногу на "
            "міліметр. Тому навчання — це завжди компроміс між швидкістю й стабільністю.\n\n"
            "Ще один практичний нюанс: у реальності обробляють не один приклад за раз, а "
            "цілі шматки датасету — їх називають батчами. Без батчів навчання було б "
            "нестерпно довгим. А один повний прохід по всьому датасету іноді називають "
            "епохою. У великих моделях, утім, важливіші не епохи, а загальний обсяг даних та "
            "обчислень.\n\n"
            "Прикинемо масштаб. Навіть на скромних ста тисячах прикладів, якщо повне "
            "засвоєння одного прикладу займе лише дві секунди, на все підуть близько "
            "п’ятдесяти п’яти годин. А в реальності прохід триває значно довше й залежить "
            "від «заліза». Тож навчання нейромереж — справа справді довга.\n\n"
            "Практична порада, якщо колись налаштовуватимете навчання:\n"
            "- Параметр швидкості навчання шукайте в конфігурації як «learning rate» — "
            "Налаштування навчання - learning rate.\n"
            "  - Почніть із типового значення (наприклад, 0,0001) і змінюйте поступово.\n"
            "  - Якщо втрати «стрибають» хаотично — зменшіть крок; якщо майже не "
            "зменшуються — збільшіть.\n"
            "- Розмір батча (batch size) підбирайте під обсяг пам’яті GPU."
        ),
    )

    # --- 26 · ГРАФІК: ВАРТІСТЬ НАВЧАННЯ ---
    b.add_chart_slide(
        "Вартість навчання флагманських моделей",
        "bar",
        {
            "categories": ["GPT-4", "GPT-5"],
            "series": {"Млн доларів США (оцінка)": [100, 2000]},
        },
        insight=[
            "GPT-4 — близько 100 млн доларів; GPT-5 — близько 2 млрд доларів.",
            "І це лише робота з фінальною моделлю — без зарплат, збору датасетів і закупівлі "
            "«заліза».",
            "Відкритий датасет Common Crawl містить сотні мільярдів сторінок.",
        ],
        kicker="Дані · Економіка навчання",
        number=26,
        source_citation="Публічні оцінки вартості навчання GPT-4 / GPT-5; " + SRC_CRAWL,
        notes_text=(
            "Ми говорили, що навчання — процес довгий. Тепер побачимо, наскільки він "
            "дорогий, і цифри вражають навіть підготовлену людину. За оцінками, на навчання "
            "GPT-4 пішло близько ста мільйонів доларів. А на GPT-5 — уже близько двох "
            "мільярдів. Зверніть увагу на масштаб стрибка: приблизно у двадцять разів між "
            "поколіннями.\n\n"
            "І тут критично важлива примітка: ці суми — лише робота з фінальною моделлю. "
            "Сюди не входять зарплати інженерів, збір і очищення датасетів, розробка "
            "проміжних моделей і закупівля «заліза». Реальні сукупні витрати ще вищі. Плюс "
            "постійні витрати енергії та амортизація обладнання.\n\n"
            "Щоб відчути обсяг даних: відкритий датасет Common Crawl для порівняння містить "
            "сотні мільярдів сторінок. Це зріз значної частини всього публічного інтернету.\n\n"
            "Який висновок робить із цього звичайна людина чи невеликий бізнес? Дуже "
            "оптимістичний. Вам ніколи не доведеться навчати таку модель самостійно — це "
            "прерогатива кількох найбільших компаній світу. Ваше завдання — розумно "
            "користуватися плодами цих інвестицій.\n\n"
            "Практична порада:\n"
            "- Замість навчання з нуля використовуйте готові моделі через API — платите лише "
            "за фактичні запити: Панель - Billing - оплата за токени.\n"
            "- Для вузької задачі достатньо тонкого налаштування готової моделі на кількох "
            "сотнях прикладів — це тисячі разів дешевше.\n\n"
            "У великих моделей навчання складається з двох великих етапів. Про них — далі."
        ),
    )

    # --- 27 · ІНФОГРАФІКА: ЕТАПИ НАВЧАННЯ LLM ---
    b.add_infographic_slide(
        "Два етапи навчання великих мовних моделей",
        [
            ("Предтренування", "Модель учиться передбачати наступний токен на гігантських "
                              "обсягах тексту. Розмітка людиною не потрібна: відповідь уже в "
                              "тексті."),
            ("Інструктивне\nдонавчання (SFT)", "Показують пари «запит → бажана відповідь». "
                                                "Модель учиться відповідати по суті, а не просто "
                                                "продовжувати текст."),
            ("Навчання з відгуком\nлюдини (RLHF)", "Люди порівнюють відповіді й обирають кращі. "
                                                    "Модель підлаштовується під тон, форму й "
                                                    "корисність."),
            ("Fine-tuning\nта RAG", "Донавчання під галузь (право, медицина) або підключення "
                                     "зовнішніх знань — пошуку й документів."),
        ],
        kicker="Процес · Етапи LLM",
        number=27,
        source_citation=SRC_VIDEO,
        notes_text=(
            "У великих мовних моделей навчання зазвичай має два великі етапи, і на кожному "
            "працює вже знайомий нам цикл. Розберемо їх послідовно — це пояснює, чому "
            "ChatGPT не просто продовжує ваш текст, а справді відповідає.\n\n"
            "Перший етап — предтренування. Модель учиться передбачати наступний токен на "
            "гігантських обсягах тексту. Тут не потрібна людина для розмітки відповідей: "
            "текст сам містить правильну відповідь, адже наступне слово вже написане "
            "автором. На цьому етапі модель набирає базові навички: мову, стиль, факти, "
            "шаблони міркувань, масу статистики про світ. Вона стає потужним "
            "«продовжувачем тексту», але ще не обов’язково зручним співрозмовником — може "
            "відхилятися від теми й бути надміру самовпевненою.\n\n"
            "Тому проводять другий етап — посттренування, і в ньому два кроки. Перший — "
            "інструктивне донавчання, або SFT: моделі показують «ось запит, а ось бажана "
            "відповідь». Її вчать не продовжувати текст узагалі, а відповідати по задачі. "
            "Другий крок — навчання за зворотним зв’язком від людини, RLHF: люди порівнюють "
            "кілька відповідей і обирають кращу. Так модель підлаштовується під те, що люди "
            "вважають адекватним за тоном, формою й корисністю.\n\n"
            "А далі, за потреби, модель донавчають під конкретну область — юриспруденцію, "
            "медицину, фінанси. Це називають файнтюнінгом, тонким налаштуванням. Іноді ж "
            "змінювати ваги взагалі не треба: простіше підключити зовнішні знання — пошук "
            "або документи, які модель читає перед відповіддю. Цей підхід називають RAG.\n\n"
            "Практична порада, як обрати:\n"
            "- Потрібні свіжі чи приватні дані — використовуйте RAG: Асистент - Завантажити "
            "документи / База знань - Додати файли.\n"
            "- Потрібен стабільний стиль чи вузька експертиза — робіть fine-tuning."
        ),
    )

    # --- 28 · ГАЛЮЦИНАЦІЇ: ЧОМУ ---
    b.add_concept_slide(
        "Чому ШІ впевнено «вигадує»",
        [
            "Модель не знає — вона вгадує правдоподібне продовження тексту.",
            "Довго тести нагороджували за точність: якщо не впевнена, вигадати вигідніше, "
            "ніж сказати «не знаю».",
            "Частину фактів вона бачить лише в тексті, без прив’язки «це — правда»; рідкісні "
            "факти неможливо надійно вгадати з форми.",
            "Нові моделі (GPT-5) галюцинують менше: їх учать визнавати невизначеність, а не "
            "брехати.",
        ],
        lead="Галюцинація — не збій, а природний наслідок «вгадування».",
        kicker="Концепція · Обмеження",
        number=28,
        visual_caption="КАДР З ВІДЕО · ~18:07 — приклади галюцинацій",
        source_citation=SRC_VELLUM,
        notes_text=(
            "Ми розібрали, які потужні ці моделі. Але вони мають серйозну ваду, про яку "
            "мусить знати кожен користувач: навіть найпросунутіші моделі можуть упевнено "
            "нести відверту нісенітницю. Це явище називають галюцинаціями. Розберемо три "
            "причини — тоді ви навчитеся їх передбачати.\n\n"
            "Причина перша й головна: модель не знає — вона вгадує. Її завдання — видати "
            "правдоподібне продовження тексту. На предтренуванні вона бачила тонни "
            "правильної за формою мови, але не бачила табличок «це правда, а це брехня». "
            "Тому деякі факти вона може лише приблизно відновити за шаблонами.\n\n"
            "Причина друга: довгий час навчання й тести нагороджували за точність. Якщо "
            "модель не впевнена, у неї два варіанти — сказати «не знаю» або спробувати "
            "вгадати. А коли оцінюють лише за точністю, вгадувати вигідніше: іноді "
            "пощастить влучити. Ми самі привчили моделі блефувати.\n\n"
            "Причина третя: не на все можна відповісти з контексту. Деякі факти рідкісні, "
            "випадкові й не виводяться зі статистики мови. Їх неможливо надійно вгадати з "
            "форми тексту, хай якою потужною буде модель.\n\n"
            "Добра новина: нові моделі, як-от GPT-5, галюцинують помітно менше. Їм дають "
            "установку — краще визнати невизначеність, ніж збрехати; посилюють "
            "самоперевірку й жорсткіше контролюють датасети.\n\n"
            "А тепер найважливіше — практика захисту від галюцинацій:\n"
            "- Просіть модель позначати невпевненість і наводити джерела — додайте в промпт: "
            "«якщо не впевнений — так і скажи; наведи посилання».\n"
            "- Умикайте пошук для перевірки фактів — ChatGPT - рядок повідомлення - "
            "значок «глобус» / Search - увімкнути.\n"
            "- Критичні факти (імена, дати, цифри, юридичні норми) завжди перевіряйте "
            "самостійно старим добрим пошуком."
        ),
    )

    # --- 29 · ГРАФІК: ЧАСТОТА ГАЛЮЦИНАЦІЙ ---
    b.add_chart_slide(
        "Частота галюцинацій сучасних моделей",
        "bar",
        {
            "categories": ["Gemini 2.5 Pro", "DeepSeek (нові версії)"],
            "series": {"Частка відповідей із галюцинаціями, %": [7, 6]},
        },
        insight=[
            "Навіть флагманські моделі помиляються у ~6–7 % випадків.",
            "Це не «іноді», а систематична властивість — плануйте перевірку фактів.",
            "Тенденція позитивна: нові покоління галюцинують дедалі рідше.",
        ],
        kicker="Дані · Надійність",
        number=29,
        source_citation=SRC_VELLUM,
        notes_text=(
            "Щоб проблема галюцинацій не звучала як абстрактне застереження, погляньмо на "
            "конкретні числа. Навіть флагманські моделі помиляються відчутно часто. За "
            "оцінками, Gemini 2.5 Pro галюцинує приблизно у семи відсотках випадків, а нові "
            "версії DeepSeek — приблизно у шести.\n\n"
            "Здавалося б, шість-сім відсотків — це небагато. Але переосмисліть це так: "
            "приблизно кожна п’ятнадцята відповідь може містити впевнено подану "
            "вигадку. Якщо ви ухвалюєте на основі відповідей важливі рішення — медичні, "
            "юридичні, фінансові — така частота абсолютно неприпустима без перевірки.\n\n"
            "Водночас тенденція обнадійлива. Кожне нове покоління моделей галюцинує рідше, "
            "бо розробники цілеспрямовано з цим борються: покращують дані, додають "
            "самоперевірку, вчать моделі визнавати невпевненість.\n\n"
            "Практичний протокол роботи, який варто зробити звичкою:\n"
            "- Ставтеся до відповіді ШІ як до чернетки розумного, але іноді неуважного "
            "стажера — корисно, але потребує перевірки.\n"
            "- Для важливих фактів застосовуйте правило двох джерел: відповідь моделі плюс "
            "незалежна перевірка пошуком.\n"
            "- Умикайте режим із посиланнями — Налаштування чату - Web search / Джерела - "
            "увімкнути, і перевіряйте самі посилання, а не лише текст.\n\n"
            "Далі перейдемо від «мозку» моделі до «тіла» — інфраструктури, на якій усе це "
            "працює."
        ),
    )

    # --- 30 · ЗАЛУЧЕННЯ ---
    b.add_engagement_slide(
        "Чому навіть найрозумніша модель\nможе впевнено помилятися?",
        prompt="Підказка: вона не «знає» фактів — вона обирає найімовірніше продовження "
               "тексту. Обговоримо у групі 2 хвилини.",
        number=30,
        source_citation=SRC_VIDEO,
        notes_text=(
            "Зробімо коротку паузу й перевіримо, чи склалася картина в голові. Питання на "
            "екрані: чому навіть найрозумніша модель може впевнено помилятися?\n\n"
            "Дайте аудиторії хвилину подумати мовчки, потім запросіть кілька відповідей. "
            "Найпоширеніша хибна відповідь — «бо їй не вистачило даних». Це лише частина "
            "правди. Підведіть до головного: причина глибша — у самій природі моделі.\n\n"
            "Ключова думка, яку варто почути від групи або сформулювати разом: модель не "
            "зберігає факти як істину з поміткою «це правда». Вона навчилася передбачати "
            "найімовірніше продовження тексту. Тому вона може згенерувати щось граматично "
            "бездоганне й упевнене за тоном, але фактично хибне. Впевненість тону й "
            "правдивість змісту — це дві абсолютно різні речі, і модель оптимізує саме "
            "перше.\n\n"
            "Гарний додатковий поштовх для дискусії:\n"
            "- Запитайте: чи довірили б ви такій моделі поставити медичний діагноз без "
            "лікаря? А написати чернетку листа? Відчуйте різницю в ціні помилки.\n\n"
            "Практичне закріплення: нагадайте групі три прийоми з попереднього слайда — "
            "просити джерела, вмикати пошук, перевіряти критичні факти самостійно. Коли "
            "відповіді пролунали, плавно переходьте до інфраструктури."
        ),
    )

    # --- 31 · ІНФРАСТРУКТУРА ---
    b.add_concept_slide(
        "Інфраструктура: тіло штучного інтелекту",
        [
            "Потрібні титанічні дата-центри: потужне живлення, серйозне охолодження, дуже "
            "швидка мережа.",
            "Обчислення виконують GPU — графічні процесори, що паралельно рахують матриці й "
            "тензори.",
            "CPU керує відеокартою, RAM — перевалочний пункт даних, а на SSD зберігають ваги "
            "й датасети.",
            "Моделі важкі: GPT-3 важила ~700 ГБ, оцінка ваг GPT-5 — близько 3600 ГБ, тому GPU "
            "об’єднують по 3–8 у сервері.",
        ],
        lead="Один GPU не вміщає велику модель — тому їх об’єднують у кластери.",
        kicker="Концепція · Залізо",
        number=31,
        visual_caption="КАДР З ВІДЕО · ~19:40 — дата-центр і GPU",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Ми розібралися з «мозком» — математикою моделі. Тепер погляньмо на «тіло» — "
            "фізичну інфраструктуру, без якої жоден ШІ не працює. Для роботи штучного "
            "інтелекту потрібні титанічні дата-центри: з потужним живленням, серйозним "
            "охолодженням, дуже швидкою мережею та величезною кількістю кластерів із "
            "потужних серверів.\n\n"
            "Головний герой тут — GPU, графічний процесор. Саме в ньому паралельно "
            "обчислюються матриці й тензори, з якими ми познайомилися раніше. Чому саме GPU? "
            "У іграх відеокарта теж має одночасно малювати безліч об’єктів — вона від "
            "природи «паралельна». Центральний процесор, CPU, так не вміє, зате саме він "
            "керує відеокартою й роздає їй задачі. Оперативна пам’ять, RAM, — це "
            "перевалочний пункт, звідки дані потрапляють у GPU. А на SSD зберігаються ваги, "
            "датасети та інші постійні файли.\n\n"
            "Навіщо ж кілька GPU? Бо нейромережі дуже важкі. Навіть «древня» GPT-3 важила "
            "близько семисот гігабайтів. А оцінка ваг GPT-5 — приблизно три тисячі шістсот "
            "гігабайтів. Жоден окремий GPU не має стільки внутрішньої пам’яті, щоб пропустити "
            "крізь себе такий обсяг. Тому їх об’єднують — найчастіше від трьох до восьми "
            "штук в одному сервері.\n\n"
            "Практична порада, якщо захочете запустити модель локально:\n"
            "- Дивіться не лише на потужність GPU, а насамперед на обсяг його пам’яті — "
            "VRAM. Саме він визначає, яка модель поміститься.\n"
            "- Перевірити наявну відеокарту: Диспетчер завдань - Продуктивність - GPU - "
            "Виділена пам’ять GPU.\n"
            "- Для великих моделей орендуйте кілька GPU в хмарі — Сервери з GPU - оберіть "
            "кількість карток."
        ),
    )

    # --- 32 · ТАБЛИЦЯ: GPU / TPU / LPU ---
    b.add_table_slide(
        "Спеціалізовані чипи для ШІ",
        ["Тип чипа", "Призначення", "Особливість"],
        [
            ["GPU", "Універсальні обчислення", "Гнучкий: і нейромережі, і графіка, і ігри"],
            ["TPU", "Спеціально під нейромережі", "Ефективніший за GPU на профільних задачах"],
            ["LPU (Groq)", "Швидкий інференс", "Заточений під миттєву генерацію відповідей"],
            ["Власні чипи", "Google, Amazon, OpenAI", "Мета — знизити собівартість токена"],
        ],
        kicker="Порівняння · Чипи",
        number=32,
        visual_caption="КАДР З ВІДЕО · ~21:06 — GPU, TPU, LPU",
        source_citation="Огляд ринку ШІ-прискорювачів (NVIDIA, Google TPU, Groq LPU), 2025–2026.",
        notes_text=(
            "GPU гарні своєю універсальністю: вони можуть і мізками для нейромережі стати, і "
            "інтер’єр відрендерити, і гру запустити. Але універсальність означає зайві "
            "витрати енергії. Тому з’явилися спеціалізовані рішення — і в цьому суть "
            "таблиці.\n\n"
            "TPU — це чипи, створені спеціально під нейромережі. На профільних задачах вони "
            "ефективніші за універсальні GPU. LPU від компанії Groq заточений під швидкий "
            "інференс — миттєву генерацію відповідей. До речі, не плутайте компанію Groq "
            "через «q» з моделлю Ілона Маска, чия назва закінчується на «k».\n\n"
            "Окремо варто сказати про ринок. NVIDIA тут майже монополіст — за різними "
            "оцінками, до вісімдесяти п’яти відсотків ринку прискорювачів. Саме тому такі "
            "гіганти, як Google, Amazon і OpenAI, не лише купують чипи NVIDIA, а й "
            "розробляють власні. Мета в усіх одна — знизити собівартість токена, тобто "
            "вартість кожної одиниці згенерованого тексту.\n\n"
            "Що з цього корисно вам як користувачеві?\n"
            "- Вам не треба обирати чип — провайдери роблять це за вас. Але тип «заліза» "
            "впливає на швидкість і ціну.\n"
            "- Якщо для вашого сценарію критична швидкість відповіді (наприклад, голосовий "
            "бот), шукайте провайдерів на LPU/TPU — Документація API - Infrastructure / "
            "Hardware.\n"
            "- Стежте за ціною за мільйон токенів — це прямий наслідок ефективності чипів. "
            "Про динаміку цін — наступний слайд."
        ),
    )

    # --- 33 · ГРАФІК: ВАРТІСТЬ ТОКЕНА ---
    b.add_chart_slide(
        "Обвал вартості токенів (за 1 млн)",
        "line",
        {
            "categories": ["к.2022 (GPT-3.5)", "2023", "2024", "2025"],
            "series": {"Доларів за 1 млн токенів": [20.0, 4.0, 0.5, 0.04]},
        },
        insight=[
            "Кінець 2022: ~$20 за мільйон токенів. За 3 роки — падіння у ~500 разів, до ~4 "
            "центів.",
            "Причина — ефективніші чипи, квантування, дистиляція, батчинг і кешування.",
            "Ціна за мільйон токенів — ключовий показник потужності дата-центру.",
        ],
        kicker="Дані · Вартість",
        number=33,
        source_citation="Динаміка цін на інференс LLM, 2022–2025 (публічні прайси провайдерів).",
        notes_text=(
            "Ось графік, який найкраще передає темп цієї революції. Придивіться до "
            "вертикальної осі — вона показує вартість мільйона токенів. В епоху ChatGPT-3.5, "
            "наприкінці 2022 року, мільйон токенів коштував близько двадцяти доларів. За "
            "наступні три роки вартість упала приблизно у п’ятсот разів — до чотирьох центів "
            "за мільйон.\n\n"
            "Це один із найшвидших обвалів вартості технології в історії. Уявіть, що "
            "авіаквиток за три роки подешевшав би у п’ятсот разів. Саме тому ШІ так стрімко "
            "проникає всюди: те, що вчора було дорогою розкішшю, сьогодні — майже "
            "безкоштовне.\n\n"
            "Завдяки чому це сталося? Інференс максимально оптимізують. Роблять батчинг — "
            "склеюють запити пачками. Роблять квантування — стискають ваги, зменшуючи "
            "точність чисел без відчутної втрати якості. Проводять дистиляцію — роблять "
            "моделі компактнішими. І кешують усе, що можна кешувати. Плюс — ефективніші "
            "чипи, про які ми щойно говорили.\n\n"
            "Ціна за мільйон токенів — це один із головних показників потужності "
            "дата-центру, поряд із джоулями на токен, тобто енергією на обробку одного "
            "токена.\n\n"
            "Практична порада, як заощаджувати:\n"
            "- Не завжди беріть найдорожчу модель — для простих задач дешевша впорається не "
            "гірше: Вибір моделі - порівняйте ціну за 1M токенів.\n"
            "- Умикайте кешування повторюваних частин запиту, якщо провайдер це підтримує — "
            "Документація - Prompt caching.\n"
            "- Пам’ятайте: вартість часто визначає довжина входу, а не відповіді, — не "
            "роздувайте контекст без потреби."
        ),
    )

    # --- 34 · ВІДКРИТІ ТА ЗАКРИТІ МОДЕЛІ ---
    b.add_concept_slide(
        "Відкриті та закриті моделі",
        [
            "Закрита модель доступна лише через застосунок чи API: ви бачите відповідь, але "
            "ваг вам не дають (ChatGPT, Gemini).",
            "Чат і модель — не одне й те саме: чат — це кермо, модель — мотор.",
            "Відкриту модель можна завантажити й запустити локально (DeepSeek, Gemma).",
            "Від відкритості залежить усе: ціна, контроль, приватність і хто може будувати "
            "продукти.",
        ],
        lead="Локальний запуск: приватно, але повільніше за хмару.",
        kicker="Концепція · Доступ",
        number=34,
        visual_caption="КАДР З ВІДЕО · ~24:20 — «кермо і мотор»",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Важливе практичне питання: якщо сервер — це майже звичайний ПК, чи можна "
            "запустити нейромережу прямо в себе вдома? Відповідь залежить від двох речей: "
            "розміру моделі й того, закрита вона чи відкрита.\n\n"
            "Закрита модель доступна лише через застосунок або API. Ви бачите відповідь у "
            "чаті, але самі ваги вам не віддають. Так працюють ChatGPT, Gemini та остання "
            "версія Grok. Тут важливо не плутати два поняття: чат і модель — це не одне й те "
            "саме. Якщо порівняти з автомобілем, чат — це кермо, а модель — мотор. Ви "
            "керуєте через кермо, але двигун схований під капотом.\n\n"
            "Відкриту модель, навпаки, можна завантажити й запустити локально. Так працюють "
            "DeepSeek або Gemma. Від того, відкрита модель чи ні, залежить фактично все: "
            "ціна, контроль, приватність і навіть те, хто взагалі може будувати продукти "
            "поверх неї.\n\n"
            "А тепер — покрокова практика для тих, хто хоче спробувати запустити модель "
            "удома. Найпростіший шлях — програма LM Studio:\n"
            "- Встановлення LM Studio - завантажте інсталятор з офіційного сайту - "
            "встановіть як звичайну програму.\n"
            "- Пошук моделі - вкладка Search (лупа) - введіть назву (наприклад, Gemma або "
            "DeepSeek) - оберіть версію під ваш обсяг пам’яті - Download.\n"
            "- Запуск - вкладка Chat - Load model - оберіть завантажену модель - пишіть у "
            "поле, як у звичайному чаті.\n\n"
            "Альтернатива для тих, хто дружить із терміналом, — Ollama:\n"
            "- Термінал - команда `ollama run gemma` - зачекайте на завантаження - "
            "спілкуйтеся прямо в терміналі.\n\n"
            "Майте на увазі: локально швидкість буде помітно нижчою, ніж у хмарі, а зазирнути "
            "«в голову» моделі повністю не зможе навіть її творець."
        ),
    )

    # --- 35 · ЧОТИРИ ДРАЙВЕРИ МАСШТАБУ ---
    b.add_table_slide(
        "Чотири драйвери прогресу нейромереж",
        ["Драйвер", "Що сталося"],
        [
            ["Дані", "Обсяг для навчання LLM подвоювався приблизно кожні 4 місяці"],
            ["Залізо", "Продуктивність GPU з 2003 року зросла у ~7000 разів"],
            ["Розмір моделі", "Від 175 млрд параметрів (GPT-3) до 2–5 трлн (GPT-5)"],
            ["Контекстне вікно", "Від тисяч токенів до мільйона й більше"],
        ],
        with_visual=True,
        kicker="Порівняння · Масштаб",
        number=35,
        visual_caption="КАДР З ВІДЕО · ~25:31 — драйвери масштабу",
        source_citation=SRC_VIDEO,
        notes_text=(
            "Підіб’ємо технічний підсумок: за рахунок чого стався весь прогрес останніх "
            "п’яти років? Відповідь коротка — за рахунок зростання масштабу за чотирма "
            "напрямами одночасно. Ця таблиця — квінтесенція всієї нашої розмови.\n\n"
            "Драйвер перший — дані. Із 2010 року обсяг даних для навчання подвоювався "
            "приблизно кожні дев’ять-десять місяців у середньому. А якщо брати лише великі "
            "мовні моделі, зростання приголомшливе — кількість даних подвоювалася кожні "
            "чотири місяці.\n\n"
            "Драйвер другий — залізо. З 2003 року продуктивність графічних процесорів "
            "зросла приблизно в сім тисяч разів. З’явилися GPU, спеціально заточені під "
            "нейромережі.\n\n"
            "Драйвер третій — розмір моделей. У GPT-3 було сто сімдесят п’ять мільярдів "
            "параметрів. У GPT-5 — за оцінками, від двох до п’яти трильйонів. Це десятки "
            "разів більше.\n\n"
            "Драйвер четвертий — контекстне вікно. Останніми роками воно росте майже так "
            "само агресивно, як і розмір моделі. Довгий контекст дозволяє згодовувати великі "
            "документи, тримати історію діалогу й будувати агентів, які читають, "
            "порівнюють і планують.\n\n"
            "Практичний висновок для вас:\n"
            "- Прогрес не зупиняється — модель, застара сьогодні, за півроку зміниться "
            "кращою. Не прив’язуйтеся намертво до однієї.\n"
            "- Періодично переглядайте, чи не з’явилася дешевша чи розумніша модель під вашу "
            "задачу — Вибір моделі - порівняйте покоління.\n\n"
            "Але чи приведе саме лише масштабування до справжнього інтелекту? Про це — далі."
        ),
    )

    # --- 36 · AGI ТА МОДЕЛІ СВІТУ ---
    b.add_concept_slide(
        "Майбутнє: AGI та моделі світу",
        [
            "Головне питання галузі: коли ШІ стане кращим за людину в усьому? Це поява AGI — "
            "інтелекту загального призначення.",
            "Саме лише масштабування навряд чи приведе до AGI: моделі чудові з текстом, але "
            "погано розуміють реальний світ.",
            "Наступне покоління будують на моделях світу — агентах, що вчаться на власному "
            "досвіді, планують і узагальнюють (приклад: Genie від Google).",
            "Гонку за AGI порівнюють із ядерною: перемагає не одна геніальна ідея, а ціла "
            "система.",
        ],
        lead="Достатньо просунута технологія невідрізнна від магії. Але це — математика.",
        kicker="Концепція · Майбутнє",
        number=36,
        visual_caption="КАДР З ВІДЕО · ~27:45 — моделі світу / Genie",
        source_citation=SRC_VIDEO,
        notes_text=(
            "І нарешті — питання, яке хвилює всіх: що далі? Головне питання у світі ШІ звучить "
            "так: коли штучний інтелект стане кращим за людину в усьому? Це станеться з "
            "появою AGI — штучного інтелекту загального призначення, який володітиме всіма "
            "когнітивними здібностями людини.\n\n"
            "За першість у цій гонці точиться найжорсткіша конкуренція, що нагадує ядерні "
            "перегони середини ХХ століття, адже AGI дасть перевагу в усьому: науці, "
            "технологіях, економіці й навіть на війні.\n\n"
            "Чи приведе туди просте масштабування? Останніми роками прогрес ішов саме за "
            "рахунок зростання масштабу, і поки це працює. Але навряд чи цього достатньо для "
            "AGI. Нейромережі чудово впораються з текстом, непогано — з картинками, музикою "
            "та звуком, але й досі погано розуміють реальний світ і не мають прямого контакту "
            "з ним.\n\n"
            "Тому наступне покоління ШІ будують на основі моделей світу. Це агенти, які "
            "вчаться безперервно на власному досвіді, будують внутрішню модель світу, "
            "уміють планувати й добре узагальнюють. Як приклад — модель Genie від Google, "
            "здатна створювати віртуальні простори й моделювати в них фізичні та біологічні "
            "властивості реального світу.\n\n"
            "Ключова метафора наостанок: гонки такого масштабу — авіаційну, космічну, "
            "ядерну — ніколи не виграють однією вдалою машиною чи одним геніальним "
            "кресленням. Перемагає система: люди, дані, залізо, процеси.\n\n"
            "Практична настанова, яку варто запам’ятати:\n"
            "- Штучний інтелект здається магією, але це не магія, а математика. Ми поки не "
            "знайшли у Всесвіті нічого, що принципово не піддавалося б обчисленню.\n"
            "- Ставтеся до ШІ як до потужного інструмента, устрій якого можна зрозуміти. "
            "Краще опинитися за кермом цієї машини, ніж під її колесами."
        ),
    )

    # --- 37 · ЗАЛУЧЕННЯ 2 ---
    b.add_engagement_slide(
        "Що переможе в гонці за AGI —\nодна геніальна ідея чи ціла система?",
        prompt="Згадайте авіаційну, космічну та ядерну гонки. Аргументуйте свою позицію.",
        number=37,
        source_citation=SRC_VIDEO,
        notes_text=(
            "Друга пауза на роздуми — і водночас місток до фіналу. Питання на екрані: що "
            "переможе в гонці за AGI — одна геніальна ідея чи ціла система?\n\n"
            "Дайте групі висловитися. Спонукайте згадати історичні аналогії: авіаційну, "
            "космічну та ядерну гонки. Підводьте до думки з відео: такі перегони ніколи не "
            "виграються однією вдалою машиною, однією красивою ідеєю чи одним геніальним "
            "кресленням. Перемагає система — сукупність людей, даних, заліза, капіталу й "
            "відлагоджених процесів.\n\n"
            "Проведіть паралель, яку легко відчути: гарний YouTube-канал — це теж не один "
            "вдалий ролик, а концепція, формат, сценарії, монтаж, обкладинки, аналітика й "
            "постійне коригування курсу. Те саме з ШІ: за кожною «магічною» моделлю стоїть "
            "величезна злагоджена система.\n\n"
            "Практичне закріплення для аудиторії:\n"
            "- Запитайте: а у вашій власній роботі що важливіше — один геніальний хід чи "
            "система звичок? Нехай перенесуть висновок на себе.\n\n"
            "Коли дискусія вщухне, переходьте до підсумкового слайда й зберіть усі ключові "
            "думки воєдино."
        ),
    )

    # --- 38 · ПІДСУМОК ---
    b.add_summary_slide(
        [
            "ШІ — не інструмент, а мислячий агент і нова інфраструктура; розуміти його — "
            "базова грамотність.",
            "Нейромережа — це стиснутий досвід: нейрони, шари, ваги й зміщення, а не набір "
            "правил.",
            "Мовні моделі працюють із токенами й векторами; «пам’ять» задає контекстне вікно.",
            "Навчання — це мільярди циклів «прохід → втрати → зворотний прохід → "
            "градієнтний спуск».",
            "Галюцинації неминучі: модель вгадує; завжди перевіряйте критичні факти й "
            "просіть джерела.",
            "Прогрес живиться масштабом даних, заліза, розміру моделі й контексту — і "
            "коштує мільярди.",
            "Майбутнє — за моделями світу й AGI; краще бути за кермом цієї машини, ніж під "
            "її колесами.",
        ],
        number=38,
        source_citation=SRC_VIDEO,
        notes_text=(
            "Зберімо все воєдино — сім думок, які варто винести із сьогоднішньої зустрічі. "
            "Пройдімося по кожній, щоб картина остаточно склалася.\n\n"
            "Перше. Штучний інтелект — це вже не просто інструмент, а мислячий агент і нова "
            "інфраструктура, як електрика чи інтернет. Розуміти, як він працює, стає базовою "
            "грамотністю сучасної людини.\n\n"
            "Друге. У серці ШІ — нейромережа, і це стиснутий досвід, а не набір інструкцій. "
            "Нейрони, шари, ваги й зміщення разом навчаються впізнавати закономірності на "
            "прикладах.\n\n"
            "Третє. Мовні моделі не бачать слів — вони працюють із токенами й векторами, а "
            "їхню «пам’ять» у моменті задає контекстне вікно.\n\n"
            "Четверте. Навчання — це мільярди повторів циклу: прямий прохід, функція втрат, "
            "зворотний прохід і градієнтний спуск. Модель прагне не сенсу, а мінімізації "
            "помилки.\n\n"
            "П’яте, найпрактичніше. Галюцинації неминучі, бо модель вгадує правдоподібне "
            "продовження. Завжди перевіряйте критичні факти й просіть джерела.\n\n"
            "Шосте. Увесь прогрес живиться масштабом — даних, заліза, розміру моделей і "
            "контексту — і коштує мільярди доларів, тож користуйтеся готовими рішеннями.\n\n"
            "Сьоме. Майбутнє — за моделями світу й AGI. І головна настанова: краще опинитися "
            "за кермом цієї машини, ніж під її колесами.\n\n"
            "Практичний перший крок після лекції:\n"
            "- Оберіть один робочий процес і спробуйте делегувати його ШІ вже цього тижня — "
            "поле промпту - опишіть задачу - оцініть і доопрацюйте результат.\n\n"
            "Дякую за увагу. Тепер відкриваємо питання й обговорення."
        ),
    )

    b.save()


if __name__ == "__main__":
    build_presentation()
