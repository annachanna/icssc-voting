import os
from time import sleep
from dotenv import load_dotenv

import requests
from bs4 import BeautifulSoup
import gspread

# Variables form .env
load_dotenv()
ELECTION_FORM = os.getenv("ELECTION_FORM")
ELECTION_SHEET = os.getenv("ELECTION_SHEET")
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE")

# UCI WebReg variables
BASE_URL = "https://directory.uci.edu/people"
VALID_MAJORS = ["CmptSci", "CSGames", "SW Engr", "IN4MATX", "DataSci", "CSE", "BIM", "SW", "GameDes"]

# Google sheet variables
EMAIL_COL = 2
START_ROW = 2

INVALID_COLOR = {
    "red": 1.0,
    "green": 0.5,
    "blue": 0.5
}

# Delay to make sure registry can process next command
DELAY = 1.0

def student_is_part_of_school(uci_id):
    """
    Check if the student (based on UCI net ID) is part of the school of ICS.

    Args:
        uci_id: UCI net ID of student

    Return:
        valid: Whether student is an ICS student
    """
    valid = True
    url = f"{BASE_URL}/{uci_id}.txt"
    res = requests.get(url, allow_redirects=False)
    if res.status_code != 200:
        valid = False
        print(f"{uci_id} - Not in UCI directory")
    else:
        soup = BeautifulSoup(res.content, features="html.parser")
        body = soup.body.get_text().split('\n')
        idx = 0
        major = body[idx]
        while "Major" not in major and idx < len(body):
            idx += 1
            major = body[idx]
        print(f"{uci_id} - {major.split(' ')[1]}")
        if major.split(' ')[1] not in VALID_MAJORS:
            valid = False
    return valid
    
def collect_net_ids(sheet):
    """
    Collects all of the UCI net IDs from a response sheet.

    Starts at row 2 since first row 1 is all of the headers.
    UCI net IDs are usually the name before the @ in UCI email.

    Args:
        sheet: Google sheet with responses

    Returns:
        net_ids: List of UCI net IDs
    """
    net_ids = []
    row = START_ROW
    cell = sheet.cell(row, EMAIL_COL).value
    while cell is not None:
        uci_id = cell.rstrip('\n').split('@')[0]
        print(uci_id)
        net_ids.append(uci_id)
        row += 1
        cell = sheet.cell(row, EMAIL_COL).value
        sleep(DELAY)
    return net_ids
    
def update_voting_sheet(sheet, net_ids):
    """
    Checks each row of the sheet and if UCI student is not part
    of school of ICS, will highlight the row.

    Args:
        sheet: Google sheet with responses
        net_ids: UCI net IDs
    """
    for i in range(len(net_ids)):
        if not student_is_part_of_school(net_ids[i]):
            invalid_cell = f"A{START_ROW + i}:K{START_ROW + i}"
            sheet.format(invalid_cell, {"backgroundColor": INVALID_COLOR})
        sleep(DELAY)

if __name__ == "__main__":
    # Finds and gets the Google sheet with the election information
    google_service = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
    spreadsheet = google_service.open(ELECTION_FORM)
    sheet = spreadsheet.worksheet(ELECTION_SHEET)

    # Update 
    net_ids = collect_net_ids(sheet)
    update_voting_sheet(sheet, net_ids)
