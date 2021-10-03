
import os
import json
import logging

from win32com import client as wincom_client


def get_file_version(file_path):
    logging.info('Get file version of [%s]', file_path)
    if not os.path.isfile(file_path):
        raise FileNotFoundError("{!r} is not found.".format(file_path))

    wincom_obj = wincom_client.Dispatch('Scripting.FileSystemObject')
    version = wincom_obj.GetFileVersion(file_path)
    logging.info('The file version of [%s] is %s', file_path, version)
    return version.strip()


def write_json(file_path, data):
    with open(file_path, "w") as f:
        json.dump(data, f, indent=2)


def read_json(file_path):
    with open(file_path, "r") as f:
        data = json.load(f)
    return data
