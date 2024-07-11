from .models import GridRange, MergeType, RGBColor


class SheetRequestBuilder:

    @staticmethod
    def set_col_width(sheet_id, start_idx: int, end_idx: int, size: int):
        body = {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": start_idx,
                    "endIndex": end_idx,
                },
                "properties": {"pixelSize": size},
                "fields": "pixelSize",
            }
        }
        return body

    @staticmethod
    def merge_cell(sheet_id, gridRange: GridRange, merge_type: str):
        body = {
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startColumnIndex": gridRange.startColumnIndex,
                    "endColumnIndex": gridRange.endColumnIndex,
                    "startRowIndex": gridRange.startRowIndex,
                    "endRowIndex": gridRange.endRowIndex,
                },
                "mergeType": merge_type,
            }
        }
        return body

    @staticmethod
    def set_cell_bg(sheet_id, gridRange: GridRange, color: RGBColor):
        body = {
            "updateCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startColumnIndex": gridRange.startColumnIndex,
                    "endColumnIndex": gridRange.endColumnIndex,
                    "startRowIndex": gridRange.startRowIndex,
                    "endRowIndex": gridRange.endRowIndex,
                },
                "rows": [
                    {
                        "values": [
                            {
                                "userEnteredFormat": {
                                    "backgroundColor": {
                                        "red": color.red,
                                        "green": color.green,
                                        "blue": color.blue,
                                    }
                                }
                            }
                        ]
                    }
                ],
                "fields": "userEnteredFormat.backgroundColor",
            }
        }
        return body

    @staticmethod
    def set_alternate_color(sheet_id, gridRange: GridRange):
        body = {
            "addBanding": {
                "bandedRange": {
                    "range": {
                        "sheetId": sheet_id,
                        "startColumnIndex": gridRange.startColumnIndex,
                        "endColumnIndex": gridRange.endColumnIndex,
                        "startRowIndex": gridRange.startRowIndex,
                        "endRowIndex": gridRange.endRowIndex,
                    },
                    "rowProperties": {
                        "headerColorStyle": {
                            "rgbColor": {
                                "red": 200 / 255,
                                "green": 200 / 255,
                                "blue": 200 / 255,
                            }
                        },
                        "firstBandColorStyle": {
                            "rgbColor": {
                                "red": 255 / 255,
                                "green": 255 / 255,
                                "blue": 255 / 255,
                            }
                        },
                        "secondBandColorStyle": {
                            "rgbColor": {
                                "red": 243 / 255,
                                "green": 243 / 255,
                                "blue": 243 / 255,
                            }
                        },
                    },
                },
            },
        }

        return body

    @staticmethod
    def format_title_cell(sheet_id: int, gridRange: GridRange):
        body = {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startColumnIndex": gridRange.startColumnIndex,
                    "endColumnIndex": gridRange.endColumnIndex,
                    "startRowIndex": gridRange.startRowIndex,
                    "endRowIndex": gridRange.endRowIndex,
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {
                            "fontSize": 14,
                        },
                    }
                },
                "fields": "userEnteredFormat",
            }
        }
        return body

    @staticmethod
    def set_halign(sheet_id: int, gridRange: GridRange, halign: str):
        body = {
            "repeatCell": {
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": halign,
                    }
                },
                "range": {
                    "sheetId": sheet_id,
                    "startColumnIndex": gridRange.startColumnIndex,
                    "endColumnIndex": gridRange.endColumnIndex,
                    "startRowIndex": gridRange.startRowIndex,
                    "endRowIndex": gridRange.endRowIndex,
                },
                "fields": "userEnteredFormat",
            }
        }

        return body

    @staticmethod
    def format_heading_cell(sheet_id: int, gridRange: GridRange):
        body = {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startColumnIndex": gridRange.startColumnIndex,
                    "endColumnIndex": gridRange.endColumnIndex,
                    "startRowIndex": gridRange.startRowIndex,
                    "endRowIndex": gridRange.endRowIndex,
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {"fontSize": 10, "bold": True},
                    }
                },
                "fields": "userEnteredFormat",
            }
        }
        return body

    @classmethod
    def format_title_cells(cls, sheet_id: str):
        requests = []

        range = GridRange(0, 2, 0, 3)
        x = cls.merge_cell(sheet_id, range, MergeType.MERGE_ALL)
        requests.append(x)

        x = cls.set_col_width(sheet_id, start_idx=0, end_idx=1, size=40)
        requests.append(x)

        x = cls.set_col_width(sheet_id, start_idx=2, end_idx=3, size=400)
        requests.append(x)

        range = GridRange(0, 1, 0, 3)
        x = cls.format_title_cell(sheet_id, range)
        requests.append(x)

        return requests

    @staticmethod
    def gen_cell_vals(terms: dict):
        row_idx = 2
        rows_need_merge = []
        values = []
        size = len(terms.keys())
        c = 1
        for term_name, term_dates in terms.items():
            values += [[f"🚧 {term_name} term class started 🚧"]]
            row_idx += 1
            rows_need_merge.append(row_idx)

            for class_no, class_date in enumerate(term_dates, start=1):
                values.append([class_no, class_date])

            row_idx += len(term_dates)

            values += [[f"🚧 {term_name} term class ended 🚧"]]
            row_idx += 1
            rows_need_merge.append(row_idx)

            if c < size:
                values += [["~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"]]
                row_idx += 1
                rows_need_merge.append(row_idx)
            c += 1

        return values, rows_need_merge

    @classmethod
    def format_heading_rows(cls, sheet_id, rows: list):
        reqs = []
        for row in rows:
            grid = GridRange(row - 1, row, 0, 3)
            reqs.append(cls.merge_cell(sheet_id, grid, merge_type=MergeType.MERGE_ALL))
            reqs.append(cls.format_heading_cell(sheet_id, grid))
        return reqs
