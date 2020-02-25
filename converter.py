import pandas as pd
import argparse
import datetime
import json
import os
import sys


def is_json(data):
    """
    Check if data is JSON
    :param data: string
    :return: bool
    """
    try:
        json.loads(data)
    except ValueError:
        return False
    return True


def excel_to_json(input_file, output_file):
    """
    Read excel and save it as json file.

    :param input_file: path string to excel file
    :param output_file: path string to json file
    :return: void
    """
    with pd.ExcelFile(input_file) as xls:
        dict_sheets = dict()
        for sheet_name in xls.sheet_names:
            sheet = xls.parse(sheet_name=sheet_name)
            rows = list()

            for _, row in sheet.iterrows():
                row_dict = dict()

                for index, value in row.iteritems():
                    if pd.isna(value):
                        row_dict[index] = None
                    elif isinstance(value, datetime.datetime):
                        row_dict[index] = value.strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'
                    elif isinstance(value, str) and is_json(value):
                        row_dict[index] = json.loads(value)
                    elif isinstance(value, datetime.time):
                        row_dict[index] = str(value)
                    else:
                        row_dict[index] = value
                rows.append(row_dict)
            dict_sheets[sheet_name] = rows

    with open(output_file, 'w', encoding='utf-8') as json_fd:
        json.dump(dict_sheets, json_fd, sort_keys=False, indent=4, ensure_ascii=False)


def validate_input_path(path):
    """
    Validate the input path

    :param path: Path string of the input file
    :return: string
    """
    if not os.path.isfile(path):
        msg = 'No such file'
        raise argparse.ArgumentTypeError(msg)
    return path


def validate_output_path(input_file, output_file):
    """
    Validate the output path

    :param output_file: Path string of the input file
    :param input_file: Path string of the output file
    :return: string
    """
    dir, file_name = os.path.split(output_file)

    if os.path.isdir(output_file):
        name = os.path.basename(input_file)
        name = os.path.splitext(name)[0]
        output_file = os.path.join(dir, name + '.json')

    if dir and not os.path.isdir(dir):
        msg = 'No such directory: ' + dir
        raise argparse.ArgumentTypeError(msg)
    elif os.path.isfile(output_file):
        msg = "Such file's already existed"
        raise argparse.ArgumentTypeError(msg)
    return output_file


def main():
    parser = argparse.ArgumentParser(description='Generate JSON files from the excel.')
    parser.add_argument('input_file',
                        help='Absolute path of the input file.',
                        type=validate_input_path)
    parser.add_argument('output_file',
                        help='Absolute path or dir where the output files will be written or name file.',
                        type=lambda x: validate_output_path(sys.argv[1], x))

    args = parser.parse_args()
    excel_to_json(args.input_file, args.output_file)


if __name__ == '__main__':
    main()
