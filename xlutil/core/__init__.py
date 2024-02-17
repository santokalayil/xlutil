from pathlib import Path
from typing import Self
import pandas as pd


class NewExcelFile:
    """ExcelFile class to create a new excel with multiple sheets."""

    def __init__(self) -> None:
        """Initialize the ExcelFile"""
        self.sheets: dict[str, pd.DataFrame] = dict()
        self.index_info: dict[str, bool] = dict()
        self.index = 0

    def add_sheet(
        self, sheet_name: str, dataframe: pd.DataFrame, replace=True, index=True
    ) -> None:
        """Adds new sheet to the excel

        Args:
            sheet_name (str): name of the sheet
            dataframe (pd.DataFrame): data to be saved to the sheet
        """
        if not replace:
            if sheet_name in self.sheets.keys():
                raise Exception(
                    "Unable to add the new file since the sheet_name already exists"
                )

        if not isinstance(index, bool):
            raise ValueError("The index value should be True or False")

        self.index_info[sheet_name] = index
        self.sheets[sheet_name] = dataframe

    def view_sheet(self, sheet_name: str) -> pd.DataFrame:
        """view the sheet as dataframe if sheet name exists

        Args:
            sheet_name (str): name of the sheet

        Raises:
            KeyError: if sheet does not exist, raises error

        Returns:
            pd.DataFrame: sheet as DataFrame
        """
        if sheet_name in self.sheets.keys():
            return self.sheets[sheet_name]
        else:
            raise KeyError(f"The sheet '{sheet_name}' is not found in the excel object")

    def __getitem__(self, __name: str) -> pd.DataFrame:
        return self.view_sheet(__name)

    def __getattr__(self, __name: str) -> pd.DataFrame:
        if __name == "keys":
            return self.__dict__
        return self.view_sheet(__name)

    def __setitem__(self, __name: str, __value: pd.DataFrame) -> None:
        if isinstance(__value, pd.DataFrame):
            self.add_sheet(__name, __value)
        else:
            raise ValueError("The value to be set should be a pandas dataframe")

    def __delitem__(self, key):
        if key in self.sheets.keys():
            del self.sheets[key]
        else:
            raise KeyError("The sheet does not exist.")

    def save(self, filepath: Path | str) -> None:
        """Saves the excel to filesystem

        Args:
            filepath (Path | str): path or filename to save the excel file

        Raises:
            Exception: if filepath is not either path instance or str
        """

        if not (isinstance(filepath, Path) or isinstance(filepath, str)):
            raise Exception(
                "The save path should be a pathlib.Path or str formated path"
            )

        writer = pd.ExcelWriter(filepath, engine="openpyxl")

        if self.sheets:
            for sheet_name, dataframe in self.sheets.items():
                dataframe.to_excel(
                    writer, sheet_name=sheet_name, index=self.index_info[sheet_name]
                )
        writer.close()

    def __repr__(self) -> str:
        return f"""Excel file with {len(self.sheets)} sheets. 
    Sheets: {", ".join(self.sheets.keys())}"""

    def __iter__(self) -> Self:
        return self

    def __next__(self) -> pd.DataFrame:
        if self.index < len(self.sheets):
            sheet = self.sheets[list(self.sheets.keys())[self.index]]
            self.index += 1
            return sheet
        else:
            self.index = 0
            raise StopIteration

    def __dict__(self) -> dict[str, pd.DataFrame]:
        return self.sheets.copy()

    def __len__(self) -> int:
        return len(self.sheets)
