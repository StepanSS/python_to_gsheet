
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import gspread_dataframe as gd
import pprint

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://spreadsheets.google.com/feeds',
          'https://www.googleapis.com/auth/drive']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1UAG_6remfTM-muSHW5HBTpBHJYAo4IUpk1f-vq0qzzs'
RANGE_NAME = 'Sheet1!A1:E'


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    # """

    # get data frame from Excel
    excel_df = pd.read_excel('excel_templates/ConstantPositionSize.xlsx', index_col=0)
    print(excel_df)

    # connect to GSheet
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", SCOPES)
    client = gspread.authorize(creds)

    sheet = client.open_by_key(SPREADSHEET_ID)
    worksheet = sheet.get_worksheet(0)
    # val = worksheet.cell(1, D).value
    # val = worksheet.get_all_records()

    # Set excel dataframe to GSheet
    gd.set_with_dataframe(worksheet, excel_df, include_index=True, row=1, col=1)

    # get pandas dataframe from GSheet
    gsheet_df = gd.get_as_dataframe(worksheet, parse_dates=True, usecols=[0, 1, 2], skiprows=0, header=None)

    pp = pprint.PrettyPrinter()
    pp.pprint(gsheet_df)

    # Share spreadsheet
    # sheet.share('otto@example.com', perm_type='user', role='reader')
    domain = '"nXtIxdhlFDFIhsGsH38h1lfMTKg/3WhJSLpxx2jU91-T0zi6z1Qut0w"'
    # sheet.remove_permissions(domain)
    permission_list = sheet.list_permissions()

    pp.pprint(permission_list)


if __name__ == '__main__':
    main()
