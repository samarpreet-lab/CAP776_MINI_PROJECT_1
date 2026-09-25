import os
import openpyxl
from datetime import date, datetime, timedelta

# Recording period: 13 August 2026 to 20 September 2026
start_date = date(2026, 8, 13)
end_date = date(2026, 9, 20)

feeling_marks = {
    "Excellent": 5,
    "Good": 4,
    "Neutral": 3,
    "Low": 2,
    "Stressed": 1
}

satisfaction_marks = {
    "Very Satisfied": 5,
    "Satisfied": 4,
    "Neutral": 3,
    "Unsatisfied": 2,
    "Very Unsatisfied": 1
}

energy_marks = {
    "High": 3,
    "Medium": 2,
    "Low": 1
}


class InvalidRecordError(Exception):
    # used when one day in the sheet cannot be used
    pass


def average(numbers):
    total = 0
    count = 0
    for number in numbers:
        total = total + number
        count = count + 1
    if count == 0:
        return 0
    return total / count


def correlation(first_list, second_list):
    # tells if two lists move together or in opposite directions
    first_avg = average(first_list)
    second_avg = average(second_list)
    top = 0
    first_part = 0
    second_part = 0
    i = 0
    while i < len(first_list):
        first_diff = first_list[i] - first_avg
        second_diff = second_list[i] - second_avg
        top = top + (first_diff * second_diff)
        first_part = first_part + (first_diff * first_diff)
        second_part = second_part + (second_diff * second_diff)
        i = i + 1
    bottom = (first_part * second_part) ** 0.5
    if bottom == 0:
        return 0
    return top / bottom


def to_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        parts = value.strip().split("-")
        if len(parts) == 3:
            try:
                return date(int(parts[0]), int(parts[1]), int(parts[2]))
            except ValueError:
                return None
    return None


def to_number(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if number < 0:
        return None
    return number


def clean_text(value):
    if value is None:
        return ""
    text = str(value).replace("\n", " ")
    text = text.replace("’", "'")
    return text.strip()


def find_excel_file():
    folder = os.path.dirname(os.path.abspath(__file__))
    first_name = os.path.join(folder, "12601472.xlsx")
    second_name = os.path.join(folder, "12601472 (1).xlsx")
    if os.path.exists(first_name):
        return first_name
    if os.path.exists(second_name):
        return second_name
    return first_name


def find_header_row(sheet):
    for row in range(1, sheet.max_row + 1):
        for col in range(1, sheet.max_column + 1):
            if clean_text(sheet.cell(row, col).value) == "Date":
                return row
    return None


def column_map(sheet, header_row):
    columns = {}
    for col in range(1, sheet.max_column + 1):
        name = clean_text(sheet.cell(header_row, col).value)
        if name != "":
            columns[name] = col
    return columns


def read_days(sheet, header_row):
    columns = column_map(sheet, header_row)
    needed = [
        "Date",
        "Sleep (min)",
        "Fitness (min)",
        "Study (min)",
        "Coding (min)",
        "Class (min)",
        "Other Activities (min)",
        "Day's Feeling",
        "Satisfaction Level",
        "Energy Level"
    ]
    for name in needed:
        if name not in columns:
            raise KeyError(name)

    days = []
    row = header_row + 1
    while row <= sheet.max_row:
        raw_date = sheet.cell(row, columns["Date"]).value
        if raw_date is not None and raw_date != "":
            one_day = {}
            one_day["date"] = to_date(raw_date)
            one_day["sleep"] = to_number(sheet.cell(row, columns["Sleep (min)"]).value)
            one_day["fitness"] = to_number(sheet.cell(row, columns["Fitness (min)"]).value)
            one_day["study"] = to_number(sheet.cell(row, columns["Study (min)"]).value)
            one_day["coding"] = to_number(sheet.cell(row, columns["Coding (min)"]).value)
            one_day["class_min"] = to_number(sheet.cell(row, columns["Class (min)"]).value)
            one_day["other"] = to_number(sheet.cell(row, columns["Other Activities (min)"]).value)
            one_day["feeling"] = clean_text(sheet.cell(row, columns["Day's Feeling"]).value)
            one_day["satisfaction"] = clean_text(sheet.cell(row, columns["Satisfaction Level"]).value)
            one_day["energy"] = clean_text(sheet.cell(row, columns["Energy Level"]).value)
            if "Total Tracked (min)" in columns:
                one_day["tracked"] = to_number(sheet.cell(row, columns["Total Tracked (min)"]).value)
            else:
                one_day["tracked"] = None
            if "Free/Unaccounted (min)" in columns:
                one_day["free_cell"] = to_number(sheet.cell(row, columns["Free/Unaccounted (min)"]).value)
            else:
                one_day["free_cell"] = None
            days.append(one_day)
        row = row + 1
    return days


def check_date_sequence(days):
    # dates should move forward, not jump backward
    previous = None
    problems = 0
    for day in days:
        if day["date"] is None:
            continue
        if previous is not None and day["date"] < previous:
            print("Date out of order:", day["date"])
            problems = problems + 1
        previous = day["date"]
    if problems == 0:
        print("Date sequence: correct")
    else:
        print("Date sequence problems:", problems)


def check_day(day, used_dates):
    if day["date"] is None:
        raise InvalidRecordError("date missing")
    if day["date"] < start_date or day["date"] > end_date:
        raise InvalidRecordError("date outside the recording period")
    if day["date"] in used_dates:
        raise InvalidRecordError("same date written twice")

    time_names = ["sleep", "fitness", "study", "coding", "class_min", "other"]
    for name in time_names:
        if day[name] is None:
            raise InvalidRecordError("time value missing or not a number")

    total = day["sleep"] + day["fitness"] + day["study"] + day["coding"] + day["class_min"] + day["other"]
    if total > 1440:
        raise InvalidRecordError("total time is more than one day")
    if day["tracked"] is not None and abs(day["tracked"] - total) > 0.01:
        raise InvalidRecordError("total tracked does not match")
    if day["free_cell"] is not None and abs(day["free_cell"] - (1440 - total)) > 0.01:
        raise InvalidRecordError("free time does not match")

    if day["feeling"] not in feeling_marks:
        raise InvalidRecordError("feeling is not from the list")
    if day["satisfaction"] not in satisfaction_marks:
        raise InvalidRecordError("satisfaction is not from the list")
    if day["energy"] not in energy_marks:
        raise InvalidRecordError("energy is not from the list")

    day["total"] = total
    day["free"] = 1440 - total
    day["feeling_mark"] = feeling_marks[day["feeling"]]
    day["satisfaction_mark"] = satisfaction_marks[day["satisfaction"]]
    day["energy_mark"] = energy_marks[day["energy"]]


def expected_days():
    count = 0
    current = start_date
    while current <= end_date:
        count = count + 1
        current = current + timedelta(days=1)
    return count


def missing_dates(valid_days):
    recorded = []
    for day in valid_days:
        recorded.append(day["date"])
    missing = []
    current = start_date
    while current <= end_date:
        if current not in recorded:
            missing.append(current)
        current = current + timedelta(days=1)
    return missing


def values_of(days, key):
    result = []
    for day in days:
        result.append(day[key])
    return result


def calculate_tpi(days):
    return average(values_of(days, "coding"))


def calculate_aai(days):
    combined = []
    for day in days:
        combined.append(day["study"] + day["class_min"])
    return average(combined)


def calculate_phai(days):
    return average(values_of(days, "fitness"))


def calculate_sri(days):
    return average(values_of(days, "sleep"))


def calculate_abi(days):
    return average(values_of(days, "free"))


def calculate_tui(days):
    return average(values_of(days, "total"))


def calculate_ei(days):
    scores = []
    for day in days:
        one_day = (day["feeling_mark"] + day["satisfaction_mark"] + day["energy_mark"]) / 3
        scores.append(one_day)
    return average(scores)


def calculate_dci(valid_count, expected_count):
    return (valid_count / expected_count) * 100


def calculate_pai(tpi, aai, phai, sri, tui, ei, dci):
    pai = (0.15 * tpi) + (0.20 * aai) + (0.15 * phai) + (0.20 * sri)
    pai = pai + (0.15 * tui) + (0.10 * ei) + (0.05 * dci)
    return pai


def sleep_energy_text(days):
    sleep_list = values_of(days, "sleep")
    energy_list = values_of(days, "energy_mark")
    link = correlation(sleep_list, energy_list)

    low_sleep = []
    high_sleep = []
    for day in days:
        if day["energy"] == "Low":
            low_sleep.append(day["sleep"])
        if day["energy"] == "High":
            high_sleep.append(day["sleep"])

    line = "Low-energy days averaged " + str(round(average(low_sleep)))
    line = line + " min of sleep; High-energy days averaged "
    line = line + str(round(average(high_sleep))) + " min."
    if link < 0:
        line = line + " Longer sleep did not raise energy."
    return line


def study_satisfaction_text(days):
    study_list = values_of(days, "study")
    satisfaction_list = values_of(days, "satisfaction_mark")
    link = correlation(study_list, satisfaction_list)

    satisfied = []
    unsatisfied = []
    for day in days:
        if day["satisfaction_mark"] >= 4:
            satisfied.append(day["study"])
        if day["satisfaction_mark"] <= 2:
            unsatisfied.append(day["study"])

    line = "Satisfied days averaged " + str(round(average(satisfied)))
    line = line + " min of study; unsatisfied days averaged "
    line = line + str(round(average(unsatisfied))) + " min."
    if link <= 0:
        line = line + " More study did not raise satisfaction."
    return line


def coding_energy_text(days):
    coding_list = values_of(days, "coding")
    energy_list = values_of(days, "energy_mark")
    class_list = values_of(days, "class_min")
    coding_link = correlation(coding_list, energy_list)
    class_link = correlation(class_list, energy_list)

    august = []
    september = []
    for day in days:
        if day["date"].month == 8:
            august.append(day["coding"])
        if day["date"].month == 9:
            september.append(day["coding"])

    line = "Coding fell from " + str(round(average(august)))
    line = line + " min/day in August to " + str(round(average(september)))
    line = line + " min/day in September"
    if abs(class_link) > abs(coding_link):
        line = line + ", but energy followed class load, not coding time."
    else:
        line = line + "."
    return line


def show_results(valid_days, invalid_count):
    expected = expected_days()
    missing = missing_dates(valid_days)

    tpi = calculate_tpi(valid_days)
    aai = calculate_aai(valid_days)
    phai = calculate_phai(valid_days)
    sri = calculate_sri(valid_days)
    abi = calculate_abi(valid_days)
    tui = calculate_tui(valid_days)
    ei = calculate_ei(valid_days)
    dci = calculate_dci(len(valid_days), expected)
    pai = calculate_pai(tpi, aai, phai, sri, tui, ei, dci)

    print("Recording period: 13 August 2026 to 20 September 2026")
    print("Expected days:", expected)
    print("Valid days:", len(valid_days))
    print("Missing days:", len(missing))
    for one_date in missing:
        print("  missing:", one_date)
    print("Invalid records:", invalid_count)
    print()
    print("Averages")
    print("Sleep:", round(sri, 2), "min")
    print("Fitness:", round(phai, 2), "min")
    print("Study:", round(average(values_of(valid_days, "study")), 2), "min")
    print("Coding:", round(tpi, 2), "min")
    print("Class:", round(average(values_of(valid_days, "class_min")), 2), "min")
    print("Other:", round(average(values_of(valid_days, "other")), 2), "min")
    print("Free:", round(abi, 2), "min")
    print()
    print("Index values")
    print("PAI:", round(pai, 2))
    print("TPI:", round(tpi, 2), "min/day")
    print("AAI:", round(aai, 2), "min/day")
    print("PhAI:", round(phai, 2), "min/day")
    print("SRI:", round(sri, 2), "min/day")
    print("ABI:", round(abi, 2), "min/day")
    print("TUI:", round(tui, 2), "min/day")
    print("EI:", round(ei, 2), "/ 5")
    print("DCI:", round(dci, 2), "%")
    print()
    print("3. Key Findings from My Data")
    print()
    print("1. Sleep ↔ Energy")
    print(sleep_energy_text(valid_days))
    print()
    print("2. Study ↔ Satisfaction")
    print(study_satisfaction_text(valid_days))
    print()
    print("3. Coding ↔ Energy")
    print(coding_energy_text(valid_days))


def main():
    print("CAP776 - MY DATA, MY STORY")
    print("Name: Samarpreet")
    print("Registration No: 12601472")
    print("-" * 50)

    file_name = find_excel_file()
    try:
        workbook = openpyxl.load_workbook(file_name, data_only=True)
    except FileNotFoundError:
        print("Excel file not found:", file_name)
        print("Save the sheet as 12601472.xlsx in this folder.")
    except Exception as error:
        print("Could not open the Excel file.")
        print(error)
    else:
        try:
            sheet = None
            header_row = None
            for one_sheet in workbook.worksheets:
                header_row = find_header_row(one_sheet)
                if header_row is not None:
                    sheet = one_sheet
                    break
            if sheet is None:
                print("Date column was not found.")
            else:
                all_days = read_days(sheet, header_row)
                check_date_sequence(all_days)

                valid_days = []
                invalid_count = 0
                used_dates = []
                for day in all_days:
                    try:
                        check_day(day, used_dates)
                    except InvalidRecordError:
                        invalid_count = invalid_count + 1
                    else:
                        valid_days.append(day)
                        used_dates.append(day["date"])

                if len(valid_days) == 0:
                    print("No valid days found.")
                else:
                    show_results(valid_days, invalid_count)
        except KeyError as error:
            print("Missing column:", error)
        except Exception as error:
            print("Problem while reading the data.")
            print(error)
    finally:
        print("-" * 50)
        print("Finished.")


if __name__ == "__main__":
    main()
