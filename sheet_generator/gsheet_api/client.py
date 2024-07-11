import logging
import os

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from sheet_generator import LOGGER


class GoogleSheetApi:
    def __init__(self):
        self.SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = None

        if os.path.exists("./token.json"):
            creds = Credentials.from_authorized_user_file("./token.json", self.SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                LOGGER.debug("Refreshing token.")
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    f"./credentials.json", self.SCOPES
                )
                creds = flow.run_local_server(port=4903)
                LOGGER.debug(creds)
            with open(f"./token.json", "w") as token:
                token.write(creds.to_json())

        try:
            self.service = build("sheets", "v4", credentials=creds)
            LOGGER.info("Built sheet_service")

        except HttpError as error:
            LOGGER.error("An error occurred: %s" % error)

    def create_sheet(self, title: str):
        try:
            body = {"properties": {"title": title}}
            request = self.service.spreadsheets().create(
                body=body, fields="spreadsheetId"
            )
            response = request.execute()
            LOGGER.debug(response)

            self.sheet_id = response.get("spreadsheetId")
            LOGGER.info(f"Created sheet. ID: {self.sheet_id}")
        except HttpError as error:
            LOGGER.error("An error occurred: %s" % error)

    def create_tab(self, title: str):
        try:
            body = {"requests": [{"addSheet": {"properties": {"title": title}}}]}
            request = self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.sheet_id, body=body
            )
            response = request.execute()
            LOGGER.debug(response)

            tab_props = response["replies"][0]["addSheet"]["properties"]
            tab_id = tab_props["sheetId"]
            tab_name = tab_props["title"]

            LOGGER.info(f"Created {title} tab. ID: {tab_id}")
            return tab_id, tab_name
        except HttpError as error:
            LOGGER.error("An error occurred creating tab: %s" % error)

    def delete_tab(self, tab_id: int):
        try:
            body = {"requests": [{"deleteSheet": {"sheetId": tab_id}}]}
            request = self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.sheet_id, body=body
            )
            response = request.execute()
            LOGGER.debug(response)
            LOGGER.info(f"Deleted tab. ID: {tab_id}")

        except HttpError as error:
            LOGGER.error("An error occurred creating tab: %s" % error)

    def send_batch_req(self, requests: list):
        body = {"requests": requests}

        try:
            request = self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.sheet_id, body=body
            )
            response = request.execute()
            LOGGER.debug(response)
            LOGGER.info(f"Completed Batch Request.")
        except HttpError as error:
            LOGGER.error("An error occurred in batch request: %s" % error)

    def write_cell(self, range: str, values: list):
        self.service.spreadsheets().values().update(
            spreadsheetId=self.sheet_id,
            range=range,
            valueInputOption="RAW",
            body={"values": values},
        ).execute()
