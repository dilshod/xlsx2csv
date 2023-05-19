import os
import re
import shutil
import subprocess
from argparse import ArgumentParser

import pandas as pd

from xlsx2csv import convert_recursive


def convert_and_upload(bq_dest: str, local_source: str | None, c_storage_source: str | None) -> None:
    """Converts all xlsx files in a local or gcloud directory to csv and uploads them to a BigQuery table

    Args:
        bq_dest: The BigQuery table to upload to
        local_source: The local directory to search for xlsx files
        c_storage_source: The Cloud Storage directory to search for xlsx files

    Returns:
        None

    Raises:
        OSError: Raised when no xlsx files are found in the source directory
        RuntimeError: Raised when no source is specified
    """
    bq_load_options = ("--autodetect --allow_quoted_newlines=true --schema_update_option=ALLOW_FIELD_ADDITION "
                       "--schema_update_option=ALLOW_FIELD_RELAXATION --source_format=CSV")
    pattern = r"(\d{4})(\d{2})(\d{2})"

    output_dir = find_temp_dir()

    if c_storage_source:
        input_source = find_temp_dir()
        if c_storage_source.split('.')[-1] != 'xlsx':
            ending = '*.xlsx'
            if c_storage_source[-1] != '/':
                ending = '/*.xlsx'
            c_storage_source = c_storage_source + ending
        os.system(f"gcloud storage cp {c_storage_source} {input_source}")
    elif local_source:
        input_source = local_source
    else:
        raise RuntimeError("No source specified")

    if len(os.listdir(input_source)) == 0:
        raise OSError(f"No xlsx files found at {c_storage_source}")

    convert_recursive(input_source, 1, output_dir, {"escape_strings": True})

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

            name_no_spaces = file.path.replace(' ', '_')
            os.rename(file.path, name_no_spaces)

    os.system("gcloud storage rm gs://skynamo_history/temp/**")

    os.system(f"gcloud storage cp {output_dir}/*.csv gs://skynamo_history/temp")

    shutil.rmtree(output_dir)
    if c_storage_source:
        shutil.rmtree(input_source)

    cmd = ['gcloud', 'storage', 'ls', 'gs://skynamo_history/temp/']
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode == 0:
        files = result.stdout

        files = files.splitlines()
        for i, file in enumerate(files):
            print(f"{i + 1}/{len(files)}")
            os.system(f"bq load {bq_load_options} {bq_dest} {file}")
    else:
        print(result.stderr)

    os.system("gcloud storage rm -r gs://skynamo_history/temp/")


def find_temp_dir():
    for i in range(1, 50):
        if os.path.exists(f'temp_{i}'):
            continue
        else:
            os.mkdir(f'temp_{i}')
            output_dir = f'temp_{i}'
            return output_dir
    else:
        raise OSError('All temp directories are in use')


if __name__ == '__main__':
    parser = ArgumentParser(description='Upload a directory of xlsx files to a BigQuery table')
    parser.add_argument('-t', '--table', help='The output dataset and table name: "dataset.table"')
    parser.add_argument('-l', '--local-source', dest='local_source', default=None,
                        help='Optional: The local directory source: /path/to/directory')
    parser.add_argument('-c', '--cloud-source', dest='cloud_source', default=None,
                        help='Optional: The cloud storage source: gs://bucket/path')
    args = parser.parse_args()
    convert_and_upload(args.table, args.local_source, args.cloud_source)
