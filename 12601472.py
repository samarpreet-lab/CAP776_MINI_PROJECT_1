from datetime import date, datetime
from openpyxl import load_workbook

SCRIPT_FOLDER = __file__.replace("\\", "/").rsplit("/", 1)[0]
FILE_NAME = SCRIPT_FOLDER + "/12601472.xlsx"
SHEET_NAME = "Daily Log"
START_ROW = 6
START_DATE = date(2026, 8, 13)
END_DATE = date(2026, 9, 21)
EXPECTED_DAYS = (END_DATE - START_DATE).days + 1
MINUTES_PER_DAY = 1440
CORRELATION_CUTOFF = 0.31  

TIME_COLUMNS = {"sleep": 1, "fitness": 2, "study": 3,
                "coding": 4, "class": 5, "other": 7}
TOTAL_COLUMN = 8
FREE_COLUMN = 9
FEELING_COLUMN = 10
SATISFACTION_COLUMN = 11
ENERGY_COLUMN = 12

FEELING_SCORE = {"Excellent": 5, "Good": 4, "Neutral": 3, "Low": 2, "Stressed": 1}
SATISFACTION_SCORE = {"Very Satisfied": 5, "Satisfied": 4, "Neutral": 3,
                      "Unsatisfied": 2, "Very Unsatisfied": 1}
ENERGY_SCORE = {"High": 3, "Medium": 2, "Low": 1}


def average(values):
    return sum(values) / len(values)


def get_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return datetime.strptime(str(value).strip()[:10], "%Y-%m-%d").date()


def get_minutes(value):
    minutes = float(value)
    if minutes != minutes or minutes < 0 or minutes > MINUTES_PER_DAY:
        raise ValueError("minutes must be between 0 and 1440")
    return minutes


def read_workbook():
    try:
        workbook = load_workbook(FILE_NAME, data_only=True)
    except FileNotFoundError:
        print("ERROR: Put", FILE_NAME, "in the same folder as this Python file.")
        return None

    if SHEET_NAME in workbook.sheetnames:
        return workbook[SHEET_NAME]
    return workbook.active


def read_records(sheet):
    names = ["sleep", "fitness", "study", "coding", "class", "other",
             "total", "free", "feeling", "satisfaction", "energy"]
    data = {name: [] for name in names}
    accepted_dates = set()
    problems = []

    for row_number, row in enumerate(
            sheet.iter_rows(min_row=START_ROW, values_only=True), START_ROW):
        if not row or row[0] is None:
            continue

        try:
            day = get_date(row[0])
            record = {}
            for name, column in TIME_COLUMNS.items():
                record[name] = get_minutes(row[column])

            record["feeling"] = FEELING_SCORE[str(row[FEELING_COLUMN]).strip()]
            record["satisfaction"] = SATISFACTION_SCORE[str(row[SATISFACTION_COLUMN]).strip()]
            record["energy"] = ENERGY_SCORE[str(row[ENERGY_COLUMN]).strip()]

            record["total"] = sum(record[name] for name in TIME_COLUMNS)
            record["free"] = MINUTES_PER_DAY - record["total"]

            excel_total = get_minutes(row[TOTAL_COLUMN])
            excel_free = get_minutes(row[FREE_COLUMN])
            if abs(excel_total - record["total"]) > 0.01:
                raise ValueError("Excel Total Tracked does not match activity entries")
            if abs(excel_free - record["free"]) > 0.01:
                raise ValueError("Excel Free Time does not match 1440 minus total")

        except (ValueError, TypeError, KeyError, IndexError) as error:
            problems.append("row " + str(row_number) + ": " + str(error))
            continue

        if day < START_DATE or day > END_DATE:
            problems.append("row " + str(row_number) + ": date is outside the period")
        elif day in accepted_dates:
            problems.append("row " + str(row_number) + ": duplicate date " + str(day))
        elif record["total"] > MINUTES_PER_DAY:
            problems.append("row " + str(row_number) + ": total time is over 1440 minutes")
        else:
            accepted_dates.add(day)
            for name in names:
                data[name].append(record[name])

    return data, accepted_dates, problems


def pearson_correlation(x_values, y_values):
    """Calculate Pearson r. Return None if either list does not vary."""
    x_average = average(x_values)
    y_average = average(y_values)
    top = sum((x - x_average) * (y - y_average)
              for x, y in zip(x_values, y_values))
    x_bottom = sum((x - x_average) ** 2 for x in x_values)
    y_bottom = sum((y - y_average) ** 2 for y in y_values)

    if x_bottom == 0 or y_bottom == 0:
        return None
    return top / (x_bottom * y_bottom) ** 0.5


def describe_correlation(r):
    if r is None:
        return "undefined (one value did not change)"
    size = abs(r)
    if size < 0.20:
        level = "negligible"
    elif size < 0.40:
        level = "weak"
    elif size < 0.70:
        level = "moderate"
    else:
        level = "strong"
    direction = "positive" if r > 0 else "negative"
    return level + " " + direction


def calculate_indexes(data, valid_days):
    academic = [data["study"][i] + data["class"][i]
                for i in range(valid_days)]
    experience = [(data["feeling"][i] + data["satisfaction"][i]
                   + data["energy"][i]) / 3 for i in range(valid_days)]

    indexes = {
        "TPI": average(data["coding"]),
        "AAI": average(academic),
        "PhAI": average(data["fitness"]),
        "SRI": average(data["sleep"]),
        "ABI": average(data["free"]),
        "TUI": average(data["total"]),
        "EI": average(experience),
        "DCI": valid_days / EXPECTED_DAYS * 100,
    }

    # Course-prescribed PAI weighted formula.
    indexes["PAI"] = (0.15 * indexes["TPI"] + 0.20 * indexes["AAI"]
                       + 0.15 * indexes["PhAI"] + 0.20 * indexes["SRI"]
                       + 0.15 * indexes["TUI"] + 0.10 * indexes["EI"]
                       + 0.05 * indexes["DCI"])
    return indexes


def show_results(data, accepted_dates, problems):
    print("CAP776 MINOR PROJECT #1 - MY DATA, MY STORY")
    print("Samarpreet Singh | 12601472 | MCA D1P2633")
    print("\n1. ACTIVITY DATA SUMMARY")
    print("Expected days:", EXPECTED_DAYS)
    print("Valid days:", len(accepted_dates))
    print("Missing days:", EXPECTED_DAYS - len(accepted_dates))
    print("Invalid / excluded records:", len(problems))
    for problem in problems:
        print("  ", problem)

    labels = {"sleep": "Sleep", "fitness": "Fitness", "study": "Study",
              "coding": "Coding", "class": "Class",
              "other": "Other Activities", "free": "Free / Unaccounted Time"}
    for name, label in labels.items():
        print("Average " + label + "/day: %.2f min/day" % average(data[name]))

    indexes = calculate_indexes(data, len(accepted_dates))
    units = {"PAI": "", "TPI": "min/day", "AAI": "min/day",
             "PhAI": "min/day", "SRI": "min/day", "ABI": "min/day",
             "TUI": "min/day", "EI": "/ 5", "DCI": "%"}
    print("\n2. INDEX VALUES")
    for name in ["PAI", "TPI", "AAI", "PhAI", "SRI", "ABI", "TUI", "EI", "DCI"]:
        print(("%s: %.2f %s" % (name, indexes[name], units[name])).rstrip())
    print("Check: TUI + ABI = %.2f minutes (should be 1440)"
          % (indexes["TUI"] + indexes["ABI"]))

    pairs = [("Sleep - Energy", "sleep", "energy", "required"),
             ("Study - Satisfaction", "study", "satisfaction", "required"),
             ("Coding - Energy", "coding", "energy", "required"),
             ("Total Tracked Time - Energy", "total", "energy", "extra"),
             ("Class Time - Satisfaction", "class", "satisfaction", "extra")]

    for heading, pair_type in [("3. REQUIRED CORRELATIONS", "required"),
                               ("4. ADDITIONAL CORRELATIONS", "extra")]:
        print("\n" + heading)
        for title, x_name, y_name, this_type in pairs:
            if this_type == pair_type:
                r = pearson_correlation(data[x_name], data[y_name])
                if r is None:
                    print(title + ": " + describe_correlation(r))
                else:
                    cut_off = "above" if abs(r) > CORRELATION_CUTOFF else "not above"
                    print("%s: %.2f (%s; %s the approximate %.2f cut-off)"
                          % (title, r, describe_correlation(r), cut_off,
                             CORRELATION_CUTOFF))


def main():
    sheet = read_workbook()
    if sheet is None:
        return
    data, accepted_dates, problems = read_records(sheet)
    if not accepted_dates:
        print("ERROR: No valid records found. Check the Excel workbook.")
        return
    show_results(data, accepted_dates, problems)


if __name__ == "__main__":
    main()
