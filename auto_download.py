import logging
import os
import zipfile

import requests
from dotenv import load_dotenv
from win32com import client as wincom_client

load_dotenv()
logging.basicConfig(
    format="%(asctime)s %(message)s",
    datefmt="%Y/%m/%d %I:%M:%S %p",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

if os.getenv("CHROME_PATH") is None:
    raise Exception("Please set CHROME_PATH in .env")

if os.getenv("PLATFORM") is None:
    raise Exception("Please set PLATFORM in .env")

CHROME_DRIVER_BASE_URL = "https://googlechromelabs.github.io/chrome-for-testing"
CHROME_DRIVER_DOWNLOAD_URL = "https://storage.googleapis.com/chrome-for-testing-public"
DOWNLOAD_FOLDER = (
    os.getenv("DOWNLOAD_FOLDER") if os.getenv("DOWNLOAD_FOLDER") else os.getcwd()
)
CHROME_DRIVER_FOLDER = f"{DOWNLOAD_FOLDER}\\chromedriver-{os.getenv('PLATFORM')}"
CHROME_DRIVER_ZIP = f"{CHROME_DRIVER_FOLDER}.zip"
CHROME_DRIVER_EXE = f"{CHROME_DRIVER_FOLDER}\\chromedriver.exe"


def get_file_version(file_path):
    logging.info("Get file version of [%s]", file_path)
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"{file_path} is not found.")

    wincom_obj = wincom_client.Dispatch("Scripting.FileSystemObject")
    version = wincom_obj.GetFileVersion(file_path)
    logging.info("The file version of [%s] is %s", file_path, version)
    return version.strip()


def get_chrome_driver_major_version():
    chrome_path = os.getenv("CHROME_PATH")
    chrome_ver = get_file_version(chrome_path)
    chrome_major_ver = chrome_ver.split(".")[0]
    logger.info("Chrome version: %s", chrome_ver)
    return chrome_major_ver


def get_latest_driver_version(browser_ver):
    latest_api = f"{CHROME_DRIVER_BASE_URL}/LATEST_RELEASE_{browser_ver}"
    resp = requests.get(latest_api, timeout=10)
    lastest_driver_version = resp.text.strip()
    logger.info("Latest driver version: %s", lastest_driver_version)
    return lastest_driver_version


def download_driver(driver_ver, dest_folder):
    download_api = f"{CHROME_DRIVER_DOWNLOAD_URL}/{driver_ver}/{os.getenv('PLATFORM')}/chromedriver-{os.getenv('PLATFORM')}.zip"
    dest_path = os.path.join(dest_folder, os.path.basename(download_api))
    resp = requests.get(download_api, stream=True, timeout=300)

    if resp.status_code == 200:
        if not os.path.isdir(dest_folder):
            os.makedirs(dest_folder)
        with open(dest_path, "wb") as f:
            f.write(resp.content)
        logger.info("Download driver completed")
    else:
        raise Exception("Download chrome driver failed")


def unzip_driver_to_target_path(src_file, dest_path):
    with zipfile.ZipFile(src_file, "r") as zip_ref:
        zip_ref.extractall(dest_path)
    logger.info("Unzip [%s] -> [%s]", src_file, dest_path)


def check_browser_driver_available():
    if os.path.isfile(CHROME_DRIVER_EXE):
        return

    chrome_major_ver = get_chrome_driver_major_version()
    driver_ver = get_latest_driver_version(chrome_major_ver)

    download_driver(driver_ver, DOWNLOAD_FOLDER)
    unzip_driver_to_target_path(CHROME_DRIVER_ZIP, DOWNLOAD_FOLDER)
    os.remove(CHROME_DRIVER_ZIP)


if __name__ == "__main__":
    try:
        check_browser_driver_available()
    except FileNotFoundError as e:
        logger.error(e)
    except Exception as e:
        logger.error("An error occurred: %s", e)
