from textwrap import dedent as dd

import pandas as pd


def header(text: str, char: str='=') -> str:
    """Return a formatted text header.

    Args:
        text: The text to wrap (default='=').
        char: The character to wrap the text with.
    """
    wrap = char * 79
    return f"{wrap}\n{text.center(79)}\n{wrap}"


def results(metric: int, pos_hours) -> str:
    """Return the raw metric and its relative percentage in printable form.

    Args:
        metric: The value to compare (in hours).
        pos_hours: The maximum possible hours.
    """
    return f"{metric}, {(round((metric / pos_hours), 3) * 100)}%"


def remove_empty_fields(iterable: iter, removable: str) -> list:
    """Remove specified elements from an iterable and return a parsed list.

    Args:
        iterable: The iterable to parse.
        removable: The element to remove.
    """
    parsed_list = []
    
    for item in iterable:
        if str(item) != removable:
            parsed_list.append(item)
    return parsed_list


def parse_week(file_name: str, week: str) -> str:
    """Parse the specified week and return the delay totals.

    This works with a specially formatted excel file which contains an oven
    schedule broken down by weeks. Each worksheet in the excel file represents
    one week. Each week considers Monday as the first day and uses the format
    MM-DD-YY.

    Args:
        file_name: The file path to the excel file.
        week: The specified week [worksheet] that will be parsed.
    """
    
    pos_hours = 1529  # Production: 5am Monday -> 12am Sunday * 11 ovens
    # The following are sub-catagories which identify the reason for lost time
    start_up = 0
    waiting = 0
    delay = 0
    other = 0
    OTHER_ITEMS = ("Calibration", "Maintenance", "HOLIDAY", "WEATHER")
    # Create the excel data frame
    d_frame = pd.read_excel(file_name,
                            sheet_name=week,
                            usecols="C, D, E, F, G, H, I, J, K, L, M")
    clean_data = {}
    ovens = list(d_frame.keys())

    for oven in ovens:
        clean_data[oven] = remove_empty_fields(d_frame[oven].tolist(), "nan")

        for item in clean_data[oven]:
            if item == "Start-up":
                start_up += 1
            elif item == "Waiting":
                waiting += 1
            elif item == "…" or item == "...":  # unicode 8230 != "..."
                delay += 1
            elif item in OTHER_ITEMS:
                other += 1
    return dd(f"""
    Total Start-up: {results(start_up, pos_hours)}
    Total Waiting: {results(waiting, pos_hours)}
    Total Delay: {results(delay, pos_hours)}
    Total Other: {results(other, pos_hours)}
    Total [All]: {results(sum((start_up, waiting, delay, other)), pos_hours)}
    """)


def main() -> None:
    """Identify the schedule and worksheet to parse; run the script.

    Each worksheet in the excel file is named after the starting date for that
    week (starting with Monday). Users may also enter [Q] to quit.
    """
    print("Starting session...\n")
    schedule = input("Enter the file-path to the oven schedule: ")

    running = True
    while running:
        config = input("Enter the target week (MM-DD-YY) or [Q] to quit: ")

        if config.lower() == 'q':
            print("Ending session...\n")
            running = False
        else:
            try:
                print(header(f"Week of {config}"))
                print(parse_week(schedule, config))
                print('-' * 79)
            except ValueError("Enter a valid week or [Q] to quit."):
                continue


if __name__ == "__main__":
    main()
