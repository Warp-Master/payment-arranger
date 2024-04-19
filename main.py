import openpyxl
import glob
import os
import warnings
import logging
import re


# disable warnings from openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module=re.escape('openpyxl.styles.stylesheet'))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("payment.log"),
        logging.StreamHandler()
    ]
)

for filename in glob.glob('*.xlsx'):
    # read payment info
    wb = openpyxl.load_workbook(filename, data_only=True)
    ws = wb.active
    inn = ws["D21"].value
    reciever = ws["B22"].value
    purpose = ws["B27"].value
    amount = ws["T11"].value
    logging.info(f"[New payment] to={reciever} for={purpose} ({amount})")
    wb.close()

    # move file to new location
    os.makedirs(inn, exist_ok=True)
    os.rename(filename, f"{inn}/{filename}")


