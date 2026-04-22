# Heat Rate PITAG Upload Tool

## Overview

This project automates the organization of PI tag data into plant specific Excel worksheets for use in fleet performance and PITAG creation workflows.

The script reads a source Excel file containing a master list of PI tags, analyzes each tag to determine which plant it belongs to, groups the tags by plant, and generates a new Excel workbook with separate worksheets for each plant.

The resulting workbook provides a cleaner and more organized structure for reviewing, importing, and processing tags across multiple Vistra plants, specifically to make creation of the pitags into PI AF easier.

---

## Purpose

Large lists of PI tags are often difficult to manage when stored in a single spreadsheet since the PI AF excel-addins expect each tag to be isolated to the plant it is being uploaded to in addition to needing the tag's settings to be included in the same sheet. This tool helps streamline that process by automatically:

* Identifying plant ownership from tag naming conventions
* Grouping tags by plant
* Creating a structured Excel workbook for downstream operations
* Eliminating manual sorting and spreadsheet manipulation

The tool is particularly useful when working with large-scale fleet performance data (which is what I wrote this code for in the first place) or preparing batches of tags for creation/upload workflows. Although this code was written originally for creating pitags used for heat rate calculations across the fleet, this can still be slightly altered and used for other general pitag upload tasks.

---

## How It Works

At a high level, the script performs the following operations:

1. Reads an input Excel file containing PI tags
2. Parses each tag name to determine the associated plant
3. Filters out unwanted or excluded tag prefixes
4. Groups tags into plant-specific collections
5. Generates a new Excel workbook
6. Creates a worksheet for each plant
7. Writes the corresponding tags into each worksheet

The generated workbook is automatically saved as:

```text
fleet_performance_tags_to_create.xlsx
```
Obviously, if you are using this code for a different project, a more appropriate name should be picked.

---

## Input

The program expects an Excel file named:

```text
pitag_list.xlsx
```

The input spreadsheet must contain a column named:

```text
Create these POCPI tags
```

This column should contain the PI tag names that need to be processed.

---

## Output

The script generates a new Excel workbook containing:

* One worksheet per plant
* Alphabetically ordered worksheets
* A list of associated tags for each plant

Example worksheet names:

```text
Baldwin_Create
Forney_Create
Odessa_Create
```

---

## Technologies Used

* Python
* pandas
* openpyxl
* collections.defaultdict

---

## Future Enhancements

This project could be expanded in several ways in the future, including:

* Automatic formatting/styling of Excel worksheets
* Duplicate tag detection
* Validation against PI systems or databases
* Command-line arguments for configurable input/output paths
* Logging support
* Exporting to CSV in addition to Excel
* Integration into automated fleet performance pipelines
* GUI or web interface for non-technical users
* Support for dynamically loading plant definitions from configuration files
* Removing inline 'PLANTS' set and externalizing the list into another file

---

## Usage

Run the script from the project directory:

```bash
python split_data.py
```

After execution completes, the generated workbook will appear in the project directory.

---

## Notes

This project is designed around existing PI tag naming conventions. Accurate grouping depends on consistent and standardized tag naming across plants.
