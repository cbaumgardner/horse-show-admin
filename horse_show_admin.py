from enum import Enum
import pandas as pd
from styleframe import StyleFrame


"""
Load the show entries and create three output files
    1) Secretary.xlsx
    2) Schedule.xlsx
    3) Placings/Gate.xlsx
"""


class ShowType(Enum):
    HUNTER = "Hunter"
    JUMPER = "Jumper"


warmup_mapping = {
    "1-3 - Pre-Short Stirrup": "Pre-Short Stirrup Warmup",
    "4-7 - Short Stirrup Hunter*": "Short Stirrup Hunter Warmup",
    "8-10 - Novice Equitation*": "Novice Equitation Warmup",
    "15-18 - Long Stirrup Hunter": "Long Stirrup Hunter Warmup",
    "26-28 - Baby Green Hunter (2')": "Baby Green Hunter Warmup",
    "29-31 - Special Hunter - (2' or 2'6)": "Special Hunter Warmup",
    "32-34 - Open Equitation (2' or 2'6)": "Open Equitation Warmup",
    "35-37 - Pony Hunter* (2' or 2'3 2'6)": "Pony Hunter Warmup",
    "39-41 - Working Hunter* (2'6)": "Working Hunter Warmup",
    "42-44 - Green Pony/Horse* (2' or 2'3 or 2'6)": "Green Pony/Horse Warmup",
    "45-47 - Thoroughbred Hunter* (2'6)": "Thoroughbred Hunter Warmup",
    "48-50 - Child/Adult Hunter* (2'6)": "Child/Adult Hunter Warmup"
}


def load_entries() -> pd.DataFrame:
    """
    Loads input entries.xlsx file

    :return: DataFrame
    """
    starting_number = 1
    df = pd.read_excel('entries.xlsx', engine="openpyxl")
    df.insert(0, "Number", range(starting_number, starting_number + len(df)), allow_duplicates=False)
    df.replace("\"", "")
    df['Divisions'] = df['Divisions'].str.split(',')
    return df


def move_column_inplace(df, col, pos):
    col = df.pop(col)
    df.insert(pos, col.name, col)
    return df


def has_warmup(warmup: str, division: str) -> bool:
    try:
        return warmup == warmup_mapping[division]
    except KeyError:
        return False


def split_divisions(df: pd.DataFrame) -> pd.DataFrame:
    # Create one row per Division and reset index to get rid of duplicates
    df = move_column_inplace(df, 'Divisions', 0)
    df = df.explode('Divisions')
    df.reset_index(drop=True, inplace=True)
    df['Divisions'] = df['Divisions'].str.strip()
    df.sort_values('Divisions', inplace=True)
    return df


def create_secretary(entries: pd.DataFrame):
    """
    Do transformations for Secretary form and write to excel file.

    :param entries: DataFrame
    """

    secretary = entries.copy()
    secretary['Paid'] = ""
    secretary['Method'] = ""
    secretary = secretary.drop(columns=['#','Status', 'Coggins', 'Status', 'TIP Number', 'Phone',
                                        'Email', 'Date Submitted'])
    secretary['Divisions'] = secretary['Divisions'].apply(lambda x: ('\n'.join(x)))
    StyleFrame(secretary).to_excel('Secretary.xlsx', index=False).save()


def create_schedule(entries: pd.DataFrame):
    """
    Do transformations for Schedule form and write to excel file.

    :param entries: DataFrame
    """
    schedule = split_divisions(entries.copy())
    schedule = schedule.drop(columns=['#', 'Date Submitted', 'Status', 'Rider Age',
                                      'TIP Number', 'Phone', 'Email', 'Total', 'Coggins'])

    # Need to check Warmup value to warmup_mapping and change value to Yes if it matches
    for i, row in schedule.iterrows():
        warmup_val = ''
        if has_warmup(row['Warmup'],row['Divisions']):
            warmup_val = 'Yes'
        schedule.at[i, 'Warmup'] = warmup_val

    schedule.to_excel('Schedule.xlsx', index=False)


def create_placings(entries: pd.DataFrame):
    """
    Do transformations for Placings form and write to excel file.

    :param entries: DataFrame
    """
    placings = split_divisions(entries.copy())
    placings = placings.drop(columns=['#', 'Status', 'Date Submitted', 'Rider Age', 'Trainer',
                                      'TIP Number', 'Phone', 'Email', 'Warmup', 'Total', 'Coggins'])
    placings['Class 1'] = ""
    placings['Class 2'] = ""
    placings['Class 3'] = ""
    placings['Ch/Res'] = ""
    placings.to_excel('Placings.xlsx', index=False)

def main():
    entries = load_entries()
    create_secretary(entries)
    create_schedule(entries)
    create_placings(entries)


if __name__ == "__main__":
    main()




