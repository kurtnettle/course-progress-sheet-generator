import calendar
from collections import OrderedDict
from datetime import timedelta
from time import sleep

from sheet_generator import LOGGER, conf_data
from sheet_generator.gsheet_api.client import GoogleSheetApi
from sheet_generator.gsheet_api.models import GridRange
from sheet_generator.gsheet_api.request_builder import SheetRequestBuilder

WeekDay = {
    "sun": calendar.SUNDAY,
    "mon": calendar.MONDAY,
    "tue": calendar.TUESDAY,
    "wed": calendar.WEDNESDAY,
    "thr": calendar.THURSDAY,
}


def get_all_course_class_dates(data: dict):
    dates = dict()

    terms = data["term"]

    # sort terms by the start date
    terms = dict(sorted(list(terms.items()), key=lambda x: x[1]["start"]))

    for term_name, term_info in data["term"].items():
        start_date = term_info["start"]
        end_date = term_info["end"]

        for n in range((end_date - start_date).days + 1):
            day = start_date + timedelta(days=n)
            date = day.strftime(r"%d/%m/%Y")

            for course_code, value in data["courses"].items():
                course_days = value["days"]

                for i in course_days:
                    weekday = WeekDay.get(i.lower())
                    if weekday == day.weekday():
                        if dates.get(course_code):
                            if dates[course_code].get(term_name):
                                dates[course_code][term_name].append(date)
                            else:
                                dates[course_code][term_name] = [date]
                        else:
                            dates[course_code] = OrderedDict()
                            dates[course_code][term_name] = [date]

    return dates


def create_course_tab(api, helper, data):
    course_name = data["name"]
    course_code = data["code"]
    tab_name = course_code

    if "lab" in course_name.lower():
        tab_name = f"{tab_name} (L)"

    LOGGER.info("%s: creating tab.", tab_name)
    sheet_id, tab_name = api.create_tab(tab_name)

    LOGGER.info("%s: setting course title.", tab_name)
    api.write_cell(f"{tab_name}!C1", [[course_name]])
    sleep(1)

    reqs = helper.format_title_cells(sheet_id)
    LOGGER.info("%s: formatting course title cells.", tab_name)
    api.send_batch_req(reqs)
    sleep(1)

    values, rows_need_merge = helper.gen_cell_vals(data["term"])
    start = rows_need_merge[0]
    end = rows_need_merge[len(rows_need_merge) - 1]

    LOGGER.info("%s: writing dates.", tab_name)
    api.write_cell(f"{tab_name}!A{start}:B{end}", values)
    sleep(1)

    reqs = helper.format_heading_rows(sheet_id, rows_need_merge)
    LOGGER.info("%s: formatting header rows.", tab_name)
    api.send_batch_req(reqs)
    sleep(1)

    grid = GridRange(0, end, 0, 3)
    reqs = helper.set_alternate_color(sheet_id, gridRange=grid)
    LOGGER.info("%s: setting alternative colors.", tab_name)
    api.send_batch_req(reqs)
    sleep(1)

    grid = GridRange(start, end + 1, 1, 2)
    reqs = helper.set_halign(sheet_id, grid, "RIGHT")
    LOGGER.info("%s: setting halign of date column to 'right'.", tab_name)
    api.send_batch_req(reqs)


def main():
    sheet_name = f'{conf_data["semester"]["name"]} [CW Tracks]'
    dates = get_all_course_class_dates(data=conf_data)
    course_codes = sorted(dates.keys())

    api = GoogleSheetApi()
    helper = SheetRequestBuilder()

    LOGGER.info("Creating sheet '%s'", sheet_name)
    api.create_sheet(sheet_name)
    sleep(2)

    for i in course_codes:
        data = {"name": conf_data["courses"][i]["name"], "code": i, "term": dates[i]}
        create_course_tab(api, helper, data)

        LOGGER.info("Sleeping for 3 seconds.")
        sleep(3)

    LOGGER.info("Deleting the first tab created by default.")
    api.delete_tab(0)


if __name__ == "__main__":
    main()

