# -*- coding: utf-8 -*-
import os
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SPREADSHEET_ID = "1h2gEuLAee9ubKC_gRmI3EcUx88Gl4kxTTKVunqbBfkI"
SHEET_NAME     = "\u042d\u0444\u0444\u0435\u043a\u0442\u0438\u0432\u043d\u043e\u0441\u0442\u044c"
CREDS_FILE     = os.path.join(os.path.dirname(__file__), "service_account.json")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def _get_service():
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    return build("sheets", "v4", credentials=creds)


def _find_existing_row(service, order_number, designation):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!C:D"
    ).execute()
    rows = result.get("values", [])
    for i, row in enumerate(rows):
        if len(row) >= 2 and row[0] == order_number and row[1] == designation:
            return i + 1
    return None


def _next_empty_row(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!C:C"
    ).execute()
    values = result.get("values", [])
    return len(values) + 1


def upsert_position(order_number, designation, name, total_qty_done):
    """
    Upsert: if row with order_number+designation exists -> update H.
    Otherwise -> create new row with C, D, E, H.
    total_qty_done is the full current total, not a delta.
    """
    try:
        service = _get_service()
        existing_row = _find_existing_row(service, order_number, designation)

        if existing_row:
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"{SHEET_NAME}!H{existing_row}",
                valueInputOption="USER_ENTERED",
                body={"values": [[total_qty_done]]}
            ).execute()
        else:
            next_row = _next_empty_row(service)
            service.spreadsheets().values().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body={
                    "valueInputOption": "USER_ENTERED",
                    "data": [
                        {"range": f"{SHEET_NAME}!C{next_row}", "values": [[order_number]]},
                        {"range": f"{SHEET_NAME}!D{next_row}", "values": [[designation]]},
                        {"range": f"{SHEET_NAME}!E{next_row}", "values": [[name]]},
                        {"range": f"{SHEET_NAME}!H{next_row}", "values": [[total_qty_done]]},
                    ]
                }
            ).execute()

        return True

    except HttpError as e:
        raise RuntimeError(f"Google Sheets API error: {e}")
    except Exception as e:
        raise RuntimeError(f"Error writing to sheet: {e}")


def is_configured():
    return (
        os.path.exists(CREDS_FILE) and
        SPREADSHEET_ID != "YOUR_SPREADSHEET_ID"
    )
