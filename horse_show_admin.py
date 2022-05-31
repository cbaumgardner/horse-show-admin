from enum import Enum
import pandas as pd
from styleframe import StyleFrame

"""
Load the show entries and create three output files
    1) Secretary
    2) Schedule
    3) Placings/Gate
"""


class ShowType(Enum):
    HUNTER = "Hunter"
    JUMPER = "Jumper"


class HunterPrices(Enum):
    CLASS = 25
    DIVISION = 50
    WARMUP = 20
    OFFICE_FEE = 15


class JumperPrices(Enum):
    CLASS = 25
    DIVISION = 50
    WARMUP = 15
    OFFICE_FEE = 15
    UNJUDGED = 10000000


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
    #df.rename(columns={'Warm-Up Round ($15)' : 'Warmup', 'Non-Showing/Schooling ($25)' : 'Schooling'}, inplace=True)
    return df


def create_secretary(entries: pd.DataFrame, show_type: ShowType):
    """
    Creates the Secretary output.

    :param entries: DataFrame
    :param show_type: ShowType
    """

    secretary = entries.copy()
    #secretary['Total'] = secretary['Divisions_Amount'] + secretary['Warmups_Amount']
    secretary['Paid'] = ""
    secretary['Method'] = ""
    secretary = secretary.drop(columns=['#','Status', 'Coggins', 'Status', 'TIP Number', 'Phone', 'Email', 'Date Submitted'])
    secretary['Divisions'] = secretary['Divisions'].apply(lambda x: ('\n'.join(x)))
    StyleFrame(secretary).to_excel('Secretary.xlsx', index=False).save()


def create_schedule():
    pass


def create_placings():
    pass


def main():
    entries = load_entries()
    create_secretary(entries, 'Hunter')


if __name__ == "__main__":
    main()




