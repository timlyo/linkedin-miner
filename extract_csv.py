#!/usr/bin/env python3

"""
Produce a CSV from a given LinkedIn dump
"""

import argparse
import pandas
import csv


def get_arguments():
    """Setup arguments we can use to control behaviour"""
    parser = argparse.ArgumentParser(
        prog="extract_csv.py",
    )

    parser.add_argument("file", help="Linkedin data dump")
    parser.add_argument("--output", help="File to write to", default="output.csv")
    parser.add_argument("--invalid-output", help="File to write invalid lines to")

    return parser.parse_args()


def parse_spreadsheet(file):
    """Parse an Excel spreadsheet returning valid lines"""
    print(f"Reading {file}")

    spreadsheet = pandas.read_excel(file)

    valid_lines = []
    invalid_lines = []

    for index, line in spreadsheet.iterrows():
        try:
            is_valid_line, value = parse_line(line)

            if is_valid_line:
                valid_lines.append(value)
            else:
                invalid_lines.append(value)
        except csv.Error as e:
            print(f"Warning parsing line {index}: {e}")

    return valid_lines, invalid_lines


def parse_line(line):
    """Process a single line to see if it contains readable data in the correct format"""
    cell = line[0]
    parsed_cell = list(csv.reader([cell]))[0]

    if parsed_cell[0] == "" and parsed_cell[1] == "":
        # First name and Last Name are blank
        return False, parsed_cell
    elif len(parsed_cell) < 4:
        # Not enough data, excludes first few lines
        return False, parsed_cell
    else:
        # No issues found, we consider this valid
        return True, parsed_cell


def write_csv(name, lines):
    """Create our CSV files from the parsed lines"""
    print(f"Writing {name}")
    file = open(name, "w")
    writer = csv.writer(file)
    writer.writerows(lines)


if __name__ == "__main__":
    args = get_arguments()
    valid, invalid = parse_spreadsheet(args.file)

    print(f"Done. Found {len(valid)} valid lines and {len(invalid)} invalid lines")

    write_csv(args.output, valid)

    if args.invalid_output:
        write_csv(args.invalid_output, invalid)
