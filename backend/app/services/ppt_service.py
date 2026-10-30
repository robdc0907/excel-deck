from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

from pptx import Presentation


class PPTService:
    def generate_presentation(
        self,
        rows: List[Dict[str, Any]],
        selected_columns: List[Dict[str, str]],
        title: str,
        file_name: str,
    ):
        prs = Presentation()
        prs.slide_width = 13.333 * 914400
        prs.slide_height = 7.5 * 914400

        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = title
        slide.placeholders[1].text = (
            f"Source file: {file_name}\n"
            f"Generated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}"
        )

        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "Data Summary"
        body = slide.shapes.placeholders[1].text_frame
        body.text = (
            f"Rows processed: {len(rows)}\n"
            f"Columns selected: {len(selected_columns)}\n"
            f"Formula types: {', '.join(column.get('formula', 'direct') for column in selected_columns)}"
        )

        for column in selected_columns:
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = column["name"]
            tf = slide.shapes.placeholders[1].text_frame
            tf.word_wrap = True
            tf.text = self._summarize_column(rows, column["name"], column["formula"])

        output_dir = Path(__file__).resolve().parent.parent.parent / "outputs"
        output_dir.mkdir(exist_ok=True)
        out_path = output_dir / f"{datetime.utcnow().strftime('%Y%m%d%H%M%S')}_{title.replace(' ', '_')}.pptx"
        prs.save(str(out_path))
        return out_path

    def _summarize_column(self, rows: List[Dict[str, Any]], column_name: str, formula: str):
        values = self._extract_numeric_values(rows, column_name)
        if not values:
            return "No numeric data available"

        if formula == "subtraction":
            if len(values) < 2:
                return "Subtraction requires at least 2 values"
            delta = sum(values[i] - values[i - 1] for i in range(1, len(values)))
            return f"Subtraction total: {delta:.2f}"

        if formula == "multiplication":
            total = 1.0
            for value in values:
                total *= value
            return f"Multiplication result: {total:.2f}"

        if formula == "growth_percent":
            if len(values) < 2:
                return "Growth percentage requires at least 2 points"
            start = values[0]
            end = values[-1]
            pct = ((end - start) / start * 100) if start else 0
            return f"Growth: {pct:.2f}%"

        if formula == "percentage_of_total":
            total = sum(values)
            if not total:
                return "Total is zero"
            share = (values[-1] / total) * 100 if values else 0
            return f"Share of total: {share:.2f}%"

        if formula == "percent_change":
            if len(values) < 2:
                return "Percent change requires at least 2 points"
            start = values[0]
            end = values[-1]
            change = ((end - start) / abs(start) * 100) if start else 0
            return f"Percent change: {change:.2f}%"

        average = sum(values) / len(values)
        minimum = min(values)
        maximum = max(values)
        return f"Average: {average:.2f}\nMin: {minimum:.2f}\nMax: {maximum:.2f}"

    def _extract_numeric_values(self, rows: List[Dict[str, Any]], column_name: str):
        values = []
        for row in rows:
            value = row.get(column_name)
            if value is None:
                continue
            try:
                numeric = float(value)
                values.append(numeric)
            except (TypeError, ValueError):
                continue
        return values
