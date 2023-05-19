import os
import re
import shutil
import subprocess
from argparse import ArgumentParser

import pandas as pd

from xlsx2csv import convert_recursive


def convert_and_upload(source, bq_location):
    bq_load_options = "--autodetect --allow_quoted_newlines=true --schema_update_option=ALLOW_FIELD_ADDITION " \
                      "--schema_update_option=ALLOW_FIELD_RELAXATION --source_format=CSV"
    pattern = r"(\d{4})(\d{2})(\d{2})"

    for i in range(1, 50):
        if os.path.exists(f'temp_{i}'):
            continue
        else:
            os.mkdir(f'temp_{i}')
            output_dir = f'temp_{i}'
            break
    else:
        raise OSError('All temp directories are in use')

    convert_recursive(source, 1, output_dir, {"escape_strings": True})

    with os.scandir(output_dir) as it:
        for file in it:
            match = re.search(pattern, file.path)
            year = match.group(1)
            month = match.group(2)
            day = match.group(3)
            date = f"{year}-{month}-{day}"

            df = pd.read_csv(file.path)
            df.columns = [re.sub(r'[\\/ #\-]', '_', col) for col in df.columns]
            df.columns = [re.sub(r'[?\n\r]', '', col) for col in df.columns]
            df['date_exported'] = date
            df.to_csv(file.path, index=False)

    os.system("gcloud storage rm gs://skynamo_history/temp/**")

    os.system(f"gcloud storage cp {output_dir}/*.csv gs://skynamo_history/temp")

    shutil.rmtree(output_dir)

    cmd = ['gcloud', 'storage', 'ls', 'gs://skynamo_history/temp/']
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode == 0:
        files = result.stdout

        files = files.splitlines()
        for i, file in enumerate(files):
            print(f"{i + 1}/{len(files)}")
            os.system(f"bq load {bq_load_options} {bq_location} {file}")
    else:
        print(result.stderr)

    os.system("gcloud storage rm -r gs://skynamo_history/temp/")


if __name__ == '__main__':
    parser = ArgumentParser(description='Upload a directory of xlsx files to a BigQuery table')
    parser.add_argument('-d', '--directory', help='The directory')
    parser.add_argument('-t', '--table', help='The output dataset and table name: "dataset.table"')
    args = parser.parse_args()
    convert_and_upload(args.directory, args.table)
