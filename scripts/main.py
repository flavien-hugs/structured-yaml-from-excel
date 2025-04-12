from pathlib import Path
from typing import Union

import typer
import yaml

from .utils import load_sheets

app = typer.Typer(pretty_exceptions_show_locals=False)


def __delegations(sheet):
    delegations_data = [
        {"code": str(code), "name": str(name).strip()}
        for code, name in sheet["Délégations régionales"][["CODEDR", "NOMDR"]].values
    ]
    return [{"name": "Délégations régionales", "data": [{"code": "N/A", "data": delegations_data}]}]


def __regions(sheet):

    regions_data = []
    regions_grouped = sheet["Régions"].groupby("CODEDR")
    for codedr, _ in sheet["Délégations régionales"][["CODEDR", "NOMDR"]].values:
        try:
            regions_in_delegation = regions_grouped.get_group(codedr)[["CODEREG", "NOMREG"]].values
            delegation_regions = [
                {"code": str(code), "name": str(name).strip()}
                for code, name in regions_in_delegation
            ]
            if delegation_regions:
                regions_data.append({"code": str(codedr), "data": delegation_regions})
        except KeyError:
            continue

    unassigned_regions = sheet["Régions"][~sheet["Régions"]["CODEDR"].isin(sheet["Délégations régionales"]["CODEDR"])]
    if not unassigned_regions.empty:
        unassigned_data = [
            {"code": str(code), "name": str(name).strip()}
            for code, name in unassigned_regions[["CODEREG", "NOMREG"]].values
        ]

        regions_data.append({"code": "N/A", "data": unassigned_data})

    return [{"name": "Régions", "data": regions_data}]


def __departements(sheet):

    departments_data = []
    departements_grouped = sheet["Départements"].groupby("CODEREG")
    for codereg, nomreg, codedr in sheet["Régions"][["CODEREG", "NOMREG", "CODEDR"]].values:
        try:
            departments_in_region = departements_grouped.get_group(codereg)[["CODEDEP", "NOMDEP"]].values
            region_departments = [
                {"code": str(code), "name": str(name).strip()}
                for code, name in departments_in_region
            ]
            if region_departments:
                departments_data.append({"code": str(codereg), "data": region_departments})
        except KeyError:
            continue

    unassigned_departments = sheet["Départements"][
        ~sheet["Départements"]["CODEREG"].isin(sheet["Régions"]["CODEREG"])]
    if not unassigned_departments.empty:
        unassigned_data = [
            {"code": str(code), "name": str(name).strip()}
            for code, name in unassigned_departments[["CODEDEP", "NOMDEP"]].values
        ]

        regions_data.append({"code": "N/A", "data": unassigned_data})

    return [{"name": "Départements", "data": departments_data}]


def __sous_prefectures(sheet):

    sps_data = []
    sps_grouped = sheet["Sous-préfectures"].groupby("CODEDEP")
    for codedep, nomdep, codereg in sheet["Départements"][["CODEDEP", "NOMDEP", "CODEREG"]].values:
        try:
            sps_in_departement = sps_grouped.get_group(codedep)[["CODESP", "NOMSP"]].values
            department_sps = [
                {"code": str(code), "name": str(name).strip()}
                for code, name in sps_in_departement
            ]

            if department_sps:
                sps_data.append({"code": str(codedep), "data": department_sps})
        except KeyError:
            continue

    unassigned_sps = sheet["Sous-préfectures"][
        ~sheet["Sous-préfectures"]["CODEDEP"].isin(sheet["Départements"]["CODEDEP"])]
    if not unassigned_sps.empty:
        unassigned_data = [
                {"code": str(code), "name": str(name).strip()}
                for code, name in unassigned_sps[["CODESP", "NOMSP"]].values
        ]
        sps_data.append({"code": "N/A", "data": unassigned_data})

    structure = [
        {
            "name": "Sous-préfectures",
            "data": sps_data
        }
    ]

    return structure


def __localites(sheet):

    loc_data = []
    loc_grouped = sheet["Localités"].groupby("CODESP")

    for codesp, nomsp, codedep in sheet["Sous-préfectures"][["CODESP", "NOMSP", "CODEDEP"]].values:
        try:
            loc_in_sps = loc_grouped.get_group(codesp)[["CODELOC", "NOMLOC"]].values
            sps_loc = [
                {"code": str(code), "name": str(name).strip()}
                for code, name in loc_in_sps
            ]

            if sps_loc:
                loc_data.append({"code": str(codesp), "data": sps_loc})
        except KeyError:
            continue

    unassigned_loc = sheet["Localités"][~sheet["Localités"]["CODESP"].isin(sheet["Sous-préfectures"]["CODESP"])]
    if not unassigned_loc.empty:
        unassigned_data = [
            {"code": str(code), "name": str(name).strip()}
            for code, name in unassigned_loc[["CODELOC", "NOMLOC"]].values
        ]

        loc_data.append({"code": "N/A", "data": unassigned_data})

    structure = [{ "name": "Localités", "data": loc_data}]

    return structure


def __zone_denombrement(sheet):

    zone_data = []
    zone_grouped = sheet["ZD"].groupby("CODELOC")

    for codeloc, nameloc, codesp in sheet["Localités"][["CODELOC", "NOMLOC", "CODESP"]].values:
        try:
            zoned_in_loc = zone_grouped.get_group(codeloc)[["CODEZD", "NOMZD"]].values
            zoned_loc = [
                {"code": str(code), "name": str(name).strip()}
                for code, name in zoned_in_loc
            ]

            if zoned_loc:
                zone_data.append({"code": str(codeloc), "data": zoned_loc})
        except KeyError:
            continue

    structure = [{"name": "Zone de denombrement", "data": zone_data}]

    return structure


def __quartiers(sheet):

    quartiers_data = []
    quart_grouped = sheet["Quartiers"].groupby("CODELOC")

    for codeloc, nameloc, codesp in sheet["Localités"][["CODELOC", "NOMLOC", "CODESP"]].values:
        try:
            quart_in_loc = quart_grouped.get_group(codeloc)[["CODEQUART", "NOMQUART"]].values
            quart_loc = [
                {"code": str(code), "name": str(name).strip()}
                for code, name in quart_in_loc
            ]

            if quart_loc:
                quartiers_data.append({"code": str(codeloc), "data": quart_loc})
        except KeyError:
            continue

    unassigned_quart = sheet["Quartiers"][~sheet["Quartiers"]["CODELOC"].isin(sheet["Localités"]["CODELOC"])]
    if not unassigned_quart.empty:
        unassigned_data = [
            {"code": str(code), "name": str(name).strip()}
            for code, name in unassigned_quart[["CODEQUART", "NOMQUART"]].values
        ]

        quartiers_data.append({"code": "N/A", "data": unassigned_quart})

    structure = [{"name": "Quartiers", "data": quartiers_data}]

    return structure


@app.command(name="Generate data structure from excel file")
def main(
    filepath: Path = typer.Option(
        ...,
        "--filepath", "-f",
        help="Link or path to the Excel file."
    ),
    output_path: Path = typer.Option(
        "fixtures/output.yml",
        "--output-file", "-o",
        help="Path for the output YAML file."
    )
):
    """
    Generate a combined data structure from an Excel file and save it as a YAML file.

    This command processes the provided Excel file to extract structured data from different sheets
    (such as 'Délégations régionales', 'Régions', 'Départements', 'Sous-préfectures', 'Localités', 'Quartiers')
    and saves it into a single YAML file in the specified output path.

    Parameters:
    :param filepath: Link or path to the Excel file containing data.
    :type filepath: Union[Path, str]
    :param output_path: Path where the YAML output file will be saved (default is 'fixtures/areadata.yml').
    :type output_path: Union[Path, str]
    :return: None

    Example of use:
        To generate a YAML file from an Excel file located at "path/to/your/data.xlsx"
        and save the output to "path/to/output/areadata.yml":

        $ poetry run generate-data --filepath path/to/your/data.xlsx --output-file path/to/output/areadata.yml
    """

    class QuotedString(str):
        pass

    def quoted_presenter(dumper, data):
        return dumper.represent_scalar('tag:yaml.org,2002:str', data, style="'")

    yaml.add_representer(QuotedString, quoted_presenter)

    # Fonction pour convertir les chaînes en QuotedString
    def convert_to_quoted(obj):
        if isinstance(obj, dict):
            return {k: convert_to_quoted(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [convert_to_quoted(i) for i in obj]
        elif isinstance(obj, str):
            return QuotedString(obj)
        return obj

    sheet = load_sheets(filepath=filepath)

    delegation_structure = __delegations(sheet)
    regions_structure = __regions(sheet)
    departements_structure = __departements(sheet)
    sous_prectures_structure = __sous_prefectures(sheet)
    localites_structure = __localites(sheet)
    zone_structure = __zone_denombrement(sheet)
    # quartiers_structure = __quartiers(sheet)

    combined_structure = (
            delegation_structure + regions_structure +
            departements_structure + sous_prectures_structure +
            localites_structure + zone_structure
    )

    # Conversion des chaînes en QuotedString
    quoted_structure = convert_to_quoted(combined_structure)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as file:
        yaml.dump(quoted_structure, file, allow_unicode=True, sort_keys=False)

    typer.echo(f"Data structure generated successfully at '{output_path}'.")


if __name__ == '__main__':
    app()
