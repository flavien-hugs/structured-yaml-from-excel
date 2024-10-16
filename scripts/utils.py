from pathlib import Path
from typing import Union

import pandas as pd


def load_sheets(filepath: Union[Path, str]):
    """
    Load sheets from an Excel file specified by filepath

    :param filepath: Path to the Excel file to load sheets from (str or Path)
    :type filepath: Union[Path, str]
    :return: A dictionary of loaded sheets from the Excel file specified by filepath (dict)
    :rtype: dict
    """

    try:
        sheet = pd.read_excel(
            filepath,
            sheet_name=[
                "Délégations régionales", "Régions", "Départements",
                "Sous-préfectures", "Localités", "Quartiers"
            ],
            dtype=str
        )
        return sheet
    except FileNotFoundError as fe:
        typer.echo(f"Error : The file '{filepath}' cannot be found.", err=True)
        raise typer.Exit(code=1) from fe
    except ValueError as ve:
        typer.echo(f"File upload error : {ve}", err=True)
        raise typer.Exit(code=1) from ve
