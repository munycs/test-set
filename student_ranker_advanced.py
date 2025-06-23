from __future__ import annotations
import csv
import datetime as _dt
from pathlib import Path
from typing import List

import pandas as pd
from kivy.app import App
from kivy.clock import Clock
from kivy.lang import Builder
from kivy.metrics import dp
from kivy.properties import BooleanProperty, ListProperty
from kivymd.app import MDApp
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.button import MDFlatButton, MDRaisedButton
from kivymd.uix.dialog import MDDialog
from kivymd.uix.gridlayout import MDGridLayout
from kivymd.uix.textfield import MDTextField
from kivymd.uix.scrollview import MDScrollView
from kivymd.toast import toast

KV = '''
<RootWidget>:
    orientation: "vertical"

    MDBoxLayout:
        adaptive_height: True
        md_bg_color: app.theme_cls.primary_color
        padding: dp(6)
        spacing: dp(6)

        MDRaisedButton:
            text: "Add Subject"
            on_release: root.prompt_add_subject()
        MDRaisedButton:
            text: "Add Student"
            on_release: root.add_student()
        MDRaisedButton:
            text: "Compute"
            on_release: root.compute_scores()
        MDFlatButton:
            text: "Save CSV"
            on_release: root.save_csv()
        MDFlatButton:
            text: "Load CSV"
            on_release: root.load_csv()
        MDFlatButton:
            text: "Export Excel"
            on_release: root.export_excel()

    MDScrollView:
        MDGridLayout:
            id: table
            cols: 6
            size_hint: None, None
            width: self.minimum_width
            height: self.minimum_height
'''

Builder.load_string(KV)

class HeaderField(MDTextField):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.size_hint_x = None
        self.width = dp(90)
        self.readonly = True
        self.mode = "rectangle"
        self.halign = "center"
        self.text_color = (0, 0, 0, 1)

class ScoreField(MDTextField):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.size_hint_x = None
        self.width = dp(90)
        self.multiline = False
        self.mode = "rectangle"
        self.helper_text_mode = "on_error"
        Clock.schedule_once(self._bind_validation, 0)

    def _bind_validation(self, *_):
        root = App.get_running_app().root
        self.bind(text=lambda inst, val: root.validate_score(inst))

class RootWidget(MDBoxLayout):
    subjects = ListProperty([])
    student_rows = ListProperty([])
    stats_added = BooleanProperty(False)

    def __init__(self, **kw):
        super().__init__(**kw)
        Clock.schedule_once(self._first_build)

    def _first_build(self, *_):
        self.add_subject("Math", 1)
        self.add_subject("Science", 1)
        for _ in range(2):
            self.add_student()
        self._refresh_table()

    def _score_cell(self):
        return ScoreField()

    def _header_cell(self, text):
        return HeaderField(text=text)

    def prompt_add_subject(self):
        name_tf = MDTextField(hint_text="Subject name", mode="rectangle")
        weight_tf = MDTextField(hint_text="Weight (default 1)", mode="rectangle", input_filter="float")

        def _add(*_):
            name = (name_tf.text or "").strip() or f"Subj{len(self.subjects)+1}"
            weight = float(weight_tf.text or 1)
            self.add_subject(name, weight)
            dialog.dismiss()

        dialog = MDDialog(
            title="Add Subject",
            type="custom",
            content_cls=MDBoxLayout(orientation="vertical", spacing=dp(8), children=[weight_tf, name_tf]),
            buttons=[
                MDFlatButton(text="Cancel", on_release=lambda *_: dialog.dismiss()),
                MDRaisedButton(text="Add", on_release=_add),
            ],
        )
        dialog.open()

    def add_subject(self, name: str, weight: float = 1.0):
        self.subjects.append({"name": name, "weight": weight})
        for row in self.student_rows:
            cell = self._score_cell()
            row.score_cells.append(cell)
        self._refresh_table()

    def add_student(self):
        name_cell = self._score_cell()
        name_cell.hint_text = "Student name"
        score_cells = [self._score_cell() for _ in self.subjects]
        total = self._score_cell(); total.readonly = True
        avg = self._score_cell(); avg.readonly = True
        rank = self._score_cell(); rank.readonly = True; rank.text_color = (1, 0, 0, 1)
        self.student_rows.append(
            type("Row", (), {
                "name_cell": name_cell,
                "score_cells": score_cells,
                "total_field": total,
                "avg_field": avg,
                "rank_field": rank
            })()
        )
        self._refresh_table()

    def _update_subject_name(self, index, new_name):
        self.subjects[index]["name"] = new_name

    def _refresh_table(self):
        cols = 1 + len(self.subjects) + 3
        self.ids.table.clear_widgets()
        self.ids.table.cols = cols

        def header_field(text):
            return HeaderField(
                text=text, readonly=True, mode="rectangle",
                size_hint_x=None, width=dp(90), halign="center"
            )

        self.ids.table.add_widget(header_field("Name"))
        for i, subj in enumerate(self.subjects):
            editable_header = ScoreField()
            editable_header.text = subj["name"]
            editable_header.bind(text=lambda inst, val, idx=i: self._update_subject_name(idx, val))
            self.ids.table.add_widget(editable_header)
        self.ids.table.add_widget(header_field("Total"))
        self.ids.table.add_widget(header_field("Average"))
        self.ids.table.add_widget(header_field("Rank"))

        for row in self.student_rows:
            self.ids.table.add_widget(row.name_cell)
            for cell in row.score_cells:
                self.ids.table.add_widget(cell)
            self.ids.table.add_widget(row.total_field)
            self.ids.table.add_widget(row.avg_field)
            row.rank_field.text_color = (1, 0, 0, 1)
            self.ids.table.add_widget(row.rank_field)

        self.ids.table.width = cols * dp(90)
        self.ids.table.height = (len(self.student_rows) + 1) * dp(60)

    def validate_score(self, field: MDTextField):
        try:
            if field.text.strip() == "":
                raise ValueError
            float(field.text)
            field.error = False
        except ValueError:
            field.error = True

    def compute_scores(self):
        data = []
        for row in self.student_rows:
            name = row.name_cell.text.strip()
            scores = []
            for idx, subj in enumerate(self.subjects):
                cell = row.score_cells[idx]
                value = 0 if cell.error else float(cell.text or 0)
                scores.append(value * subj["weight"])
            total = sum(scores)
            tot_weight = sum(s["weight"] for s in self.subjects) or 1
            avg = total / tot_weight
            data.append((row, total, avg))

        data.sort(key=lambda t: t[2], reverse=True)
        for rank, (row, total, avg) in enumerate(data, 1):
            row.total_field.text = f"{total:.2f}"
            row.avg_field.text = f"{avg:.2f}"
            row.rank_field.text = str(rank)

        self._refresh_table()

    def save_csv(self):
        fp = Path.cwd() / "scores.csv"
        with fp.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Name", *[s["name"] for s in self.subjects]])
            w.writerow(["", *[s["weight"] for s in self.subjects]])
            for row in self.student_rows:
                w.writerow([row.name_cell.text.strip(), *[c.text for c in row.score_cells]])
        toast(f"Saved → {fp}")

    def load_csv(self):
        fp = Path.cwd() / "scores.csv"
        if not fp.exists():
            toast("scores.csv not found")
            return
        rows = list(csv.reader(fp.open("r", encoding="utf-8")))
        if len(rows) < 2:
            toast("CSV missing header rows")
            return
        names_row, weights_row, *student_data = rows
        self.subjects.clear()
        self.student_rows.clear()
        for n, w in zip(names_row[1:], weights_row[1:]):
            self.add_subject(n, float(w or 1))
        for r in student_data:
            self.add_student()
            row = self.student_rows[-1]
            row.name_cell.text = r[0]
            for idx, val in enumerate(r[1:]):
                row.score_cells[idx].text = val
        self._refresh_table()
        toast(f"Loaded {fp}")

    def export_excel(self):
        data = {"Name": [r.name_cell.text.strip() for r in self.student_rows]}
        for idx, subj in enumerate(self.subjects):
            data[subj["name"]] = [r.score_cells[idx].text for r in self.student_rows]
        data["Total"] = [r.total_field.text for r in self.student_rows]
        data["Average"] = [r.avg_field.text for r in self.student_rows]
        data["Rank"] = [r.rank_field.text for r in self.student_rows]
        df = pd.DataFrame(data)
        outfile = Path.cwd() / f"scores_{_dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        df.to_excel(outfile, index=False)
        toast(f"Excel → {outfile.name}")

class StudentRankerApp(MDApp):
    def build(self):
        self.theme_cls.primary_palette = "Blue"
        return RootWidget()

if __name__ == "__main__":
    StudentRankerApp().run()
