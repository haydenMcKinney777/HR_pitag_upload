import os
import pandas as pd
from collections import defaultdict

SKIP_TAG_PREFIXES = {
    "Forney_Atmos",
    "Kendall_Fuel",
    "Blackstone_BLA1_AGC_SP_CONTROL",
    "Forney_Kinder"
}

PLANTS = {
    "Baldwin",
    "Bellingham",
    "Blackstone",
    "Calumet",
    "Casco Bay",
    "Decordova",
    "Ennis",
    "Fayette",
    "Forney",
    "Graham",
    "Hanging Rock",
    "Hays",
    "Independence",
    "Kendall",
    "Kincaid",
    "Lake Hubbard",
    "Lake Road",
    "Lamar",
    "Liberty",
    "Martin Lake",
    "Masspower",
    "Miami Fort",
    "Midlothian",
    "Milford",
    "Morgan Creek",
    "Moss Landing",
    "Newton",
    "Oak Grove",
    "Odessa",
    "Ontelaunee",
    "Permian Basin",
    "Pleasants",
    "Stryker Creek",
    "Trinidad",
    "Washington",
    "Wise"
}

def extract_plant_name(tagname, plants):
    if pd.isna(tagname):
        return None

    tagname = str(tagname).strip()

    for skip_prefix in SKIP_TAG_PREFIXES:
        if tagname == skip_prefix or tagname.startswith(skip_prefix + " "):
            return None

    #sort plants so plants with longer names are first in the list. e.g. if our tag was Lake Road_Fuel 'startwith()' will pick up 'Lake Road' first before 'Lake '
    for plant in sorted(plants, key=len, reverse=True):
        if (tagname == plant or tagname.startswith(plant + " ") or tagname.startswith(plant + "_")):
            return plant

    return tagname.split(" ", 1)[0]

def main():
    print("Starting main program...\n")

    cwd = os.getcwd()
    file_path = f"{cwd}\\pitag_list.xlsx"

    df = pd.read_excel(file_path)
    plant_tag_mappings = defaultdict(list)

    for pitag in df["Create these POCPI tags"]:
        plant = extract_plant_name(pitag, PLANTS)
        if plant is None:
            continue
        plant_tag_mappings[plant].append(pitag)

    print(plant_tag_mappings.keys())
    print("Program complete.\n")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error running main - {e}")
