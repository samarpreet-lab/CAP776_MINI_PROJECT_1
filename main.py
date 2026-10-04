from datetime import date, datetime
from math import sqrt
from openpyxl import load_workbook

FOLDER = __file__.replace("\\", "/")
FOLDER = FOLDER[:FOLDER.rfind("/") + 1]
FILE_NAME = FOLDER + "12601472.xlsx"

SHEET_NAME = "Daily Log"
START_ROW = 6
START_DATE = date(2026, 8, 13)
END_DATE = date(2026, 9, 21)
EXPECTED_DAYS = (END_DATE - START_DATE).days + 1
MINUTES = 1440

FEELING = {"Excellent": 5, "Good": 4, "Neutral": 3, "Low": 2, "Stressed": 1}
SATISFACTION = {"Very Satisfied": 5, "Satisfied": 4, "Neutral": 3,"Unsatisfied": 2, "Very Unsatisfied": 1}
ENERGY = {"High": 3, "Medium": 2, "Low": 1}

# Which column each value sits in (0 is the first column of the sheet).
TIME_COLS = {"sleep": 1, "fitness": 2, "study": 3,
             "coding": 4, "class": 5, "other": 7}
RATING_COLS = {"feeling": 10, "satisfaction": 11, "energy": 12}
SCALES = {"feeling": FEELING, "satisfaction": SATISFACTION, "energy": ENERGY}

# Weights used in the PAI formula (they add up to 1.00).
WEIGHTS = {"TPI": 0.15, "AAI": 0.20, "PhAI": 0.15, "SRI": 0.20,
           "TUI": 0.15, "EI": 0.10, "DCI": 0.05}

# The relationships to work out: title, first list, second list, kind.
PAIRS = [("Sleep - Energy", "sleep", "energy", "required"),
         ("Study - Satisfaction", "study", "satisfaction", "required"),
         ("Coding - Energy", "coding", "energy", "required"),
         ("Total Tracked Time - Energy", "total", "energy", "extra"),
         ("Class Time - Satisfaction", "class", "satisfaction", "extra"),
         ("Free Time - Feeling", "free", "feeling", "extra")]

# Names and units used when printing.
LABELS = {"sleep": "Sleep", "fitness": "Fitness", "study": "Study",
          "coding": "Coding", "class": "Class", "other": "Other Activities",
          "free": "Free / Unaccounted Time"}
INDEX_UNITS = {"PAI": "", "TPI": "min/day", "AAI": "min/day", "PhAI": "min/day",
               "SRI": "min/day", "ABI": "min/day", "TUI": "min/day",
               "EI": "/ 5", "DCI": "%"}

def average(values):
    """Return the mean of a list of numbers."""
    return sum(values) / len(values)

def correlation(x, y):
    """Return the Pearson correlation coefficient of two equal-sized lists."""
    x_mean = average(x)
    y_mean = average(y)
    top = sum((a - x_mean) * (b - y_mean) for a, b in zip(x, y))
    left = sum((a - x_mean) ** 2 for a in x)
    right = sum((b - y_mean) ** 2 for b in y)
    if left == 0 or right == 0:          # a column that never changes
        return 0.0
    return top / sqrt(left * right)

def strength(r):
    """Return a simple word describing how strong a correlation is."""
    size = abs(r)
    if size < 0.20:
        return "negligible"
    if size < 0.40:
        word = "weak"
    elif size < 0.70:
        word = "moderate"
    else:
        word = "strong"
    return word + (" positive" if r > 0 else " negative")

def read_date(value):
    """Turn a cell into a date. Raises an error if the cell is not a date."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return datetime.strptime(str(value).strip()[:10], "%Y-%m-%d").date()

def read_minutes(value):
    """Turn a cell into a number of minutes between 0 and 1440."""
    minutes = float(value)               # fails if the cell is not a number
    if minutes < 0 or minutes > MINUTES:
        raise ValueError("value is outside 0-1440 minutes")
    return minutes

def find_missing_days(recorded):
    """Return the expected dates that were not recorded."""
    missing = []
    for step in range(EXPECTED_DAYS):
        day = date.fromordinal(START_DATE.toordinal() + step)
        if day not in recorded:
            missing.append(day)
    return missing

# ------------------------------------------------------------- reading data

def open_sheet():
    """Open the workbook and return the Daily Log sheet, or None if it fails."""
    try:
        book = load_workbook(FILE_NAME, data_only=True)
    except FileNotFoundError:
        print("ERROR: Excel file not found ->", FILE_NAME)
        return None
    except Exception as error:
        print("ERROR: the Excel file could not be opened ->", error)
        return None
    if SHEET_NAME in book.sheetnames:
        return book[SHEET_NAME]
    return book.active

def read_data(sheet):
    """Read and check every row, keeping only rows whose values are all
    present and sensible, so problems are counted and never assumed."""
    names = list(TIME_COLS) + list(RATING_COLS) + ["total", "free"]
    data = {name: [] for name in names}
    accepted = set()
    problems = []
    row_number = START_ROW - 1

    for row in sheet.iter_rows(min_row=START_ROW, values_only=True):
        row_number = row_number + 1
        if row[0] is None:               # ignore empty rows
            continue

        try:
            day = read_date(row[0])
            record = {}
            for name in TIME_COLS:
                record[name] = read_minutes(row[TIME_COLS[name]])
            for name in RATING_COLS:
                record[name] = SCALES[name][str(row[RATING_COLS[name]]).strip()]
        except (TypeError, ValueError, KeyError, IndexError) as error:
            problems.append("row %d : %s" % (row_number, error))
            continue

        # Total and free time are worked out from the activity values.
        record["total"] = sum(record[name] for name in TIME_COLS)
        record["free"] = MINUTES - record["total"]

        if day < START_DATE or day > END_DATE:
            problems.append("row %d : date %s is outside the period" % (row_number, day))
        elif day in accepted:
            problems.append("row %d : date %s is repeated" % (row_number, day))
        elif record["total"] > MINUTES:
            problems.append("row %d : tracked time %.0f is over 1440 minutes"
                            % (row_number, record["total"]))
        else:
            accepted.add(day)
            for name in names:
                data[name].append(record[name])

    return data, accepted, problems

def calculate_indexes(data, valid_days):
    """Work out all nine indexes and return them in a dictionary."""
    academic = [data["study"][i] + data["class"][i] for i in range(valid_days)]
    experience = [(data["feeling"][i] + data["satisfaction"][i] +
                   data["energy"][i]) / 3 for i in range(valid_days)]
    indexes = {"TPI": average(data["coding"]),
               "AAI": average(academic),
               "PhAI": average(data["fitness"]),
               "SRI": average(data["sleep"]),
               "ABI": average(data["free"]),
               "TUI": average(data["total"]),
               "EI": average(experience),
               "DCI": valid_days / EXPECTED_DAYS * 100}
    indexes["PAI"] = sum(WEIGHTS[name] * indexes[name] for name in WEIGHTS)
    return indexes

def calculate_correlations(data):
    """Work out r and a strength word for every pair listed in PAIRS."""
    results = []
    for title, first, second, kind in PAIRS:
        r = correlation(data[first], data[second])
        results.append((title, r, strength(r), kind))
    return results

def print_summary(data, accepted, problems):
    """Print section 1 of the report."""
    missing = find_missing_days(accepted)
    print("\n1. ACTIVITY DATA SUMMARY")
    print("Expected number of days:", EXPECTED_DAYS)
    print("Valid days recorded:", len(accepted))
    print("Missing days:", len(missing))
    print("Invalid / excluded records:", len(problems))
    if missing:
        print("Missing dates:", ", ".join(str(day) for day in missing))
    for note in problems:
        print("   problem ->", note)
    for name in LABELS:
        print(f"Average {LABELS[name]}/day: {average(data[name]):.2f} min/day")

def print_indexes(indexes):
    """Print section 2 of the report."""
    print("\n2. INDEX VALUES")
    for name in INDEX_UNITS:               # already in the right order
        print(f"{name}: {indexes[name]:.2f} {INDEX_UNITS[name]}".rstrip())
    print(f"Check: TUI + ABI = {indexes['TUI'] + indexes['ABI']:.2f} "
          f"minutes (must be {MINUTES})")

def print_correlations(results):
    """Print sections 3 and 4 of the report."""
    for heading, wanted in [("\n3. REQUIRED CORRELATIONS", "required"),
                            ("\n4. ADDITIONAL CORRELATIONS", "extra")]:
        print(heading)
        for title, r, word, kind in results:
            if kind == wanted:
                print(f"{title}: {r:.2f} ({word})")

def main():
    """Run the complete analysis."""
    sheet = open_sheet()
    if sheet is None:
        return
    data, accepted, problems = read_data(sheet)
    if len(accepted) == 0:
        print("ERROR: no valid rows were found. Please check the Excel file.")
        return
    print("CAP776 MINOR PROJECT #1 - MY DATA, MY STORY")
    print("Samarpreet Singh | 12601472 | MCA D1P2633")
    print_summary(data, accepted, problems)
    print_indexes(calculate_indexes(data, len(accepted)))
    print_correlations(calculate_correlations(data))

if __name__ == "__main__":
    main()
