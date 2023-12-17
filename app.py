import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import tkinter as tk


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1w6zfS9AESuJ98fJYwc5Ypc68FjlMaElGo5x3tJOsGBU"
SAMPLE_RANGE_NAME = "Página1!A1:C10"


def main():
  
  print("Empresa: %s\nData: %s \nSituação: %s" % (e1.get(), e2.get(), e3.get()))
  """Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()


    adicionar = [
      [e1.get(), e2.get(), e3.get()],
    ]


    buscaTamanho = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="A:A")
        .execute()
    )

    tamanho = "A" + str(len(buscaTamanho['values']) + 1)

    print(tamanho)
    body = {"values": adicionar}
    result = (
        service.spreadsheets()
        .values()
        .update(
            spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range=tamanho,
            valueInputOption="RAW",
            body=body,
        )
        .execute()
    )
    

  except HttpError as err:
    print(err)

master = tk.Tk()
tk.Label(master, text="Empresa").grid(row=0)
tk.Label(master, text="Data").grid(row=1)
tk.Label(master, text="Situação").grid(row=2)

e1 = tk.Entry(master)
e2 = tk.Entry(master)
e3 = tk.Entry(master)

e1.grid(row=0, column=1)
e2.grid(row=1, column=1)
e3.grid(row=2, column=1)



tk.Button(master, 
          text='Sair', 
          command=master.quit).grid(row=3, 
                                    column=0, 
                                    sticky=tk.W, 
                                    pady=4)
tk.Button(master, 
          text='Cadastrar', command=main).grid(row=3, 
                                                       column=1, 
                                                       sticky=tk.W, 
                                                       pady=4)


master.mainloop()
