import gspread
from oauth2client.service_account import ServiceAccountCredentials

class SheetService:
    def __init__(self,sheet_id,sheet_name):
        self.sheet_id = sheet_id
        self.sheet_name = sheet_name

    def _connect(self, sheet_id, sheet_name):
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('key/key.json', scope)
        client = gspread.authorize(credentials)
        sheet = client.open(f'{sheet_name}')
        self.worksheet = sheet.get_worksheet(0)
        return self.worksheet
    def get_order_status(self, tovaritelnica):
        for row in self.worksheet.get_all_records():
            if row.get(f'{tovaritelnica}'):
                match_row = list(row.values())
                return match_row[4]
        return None
    def update_order_status(self, tovaritelnica, status, time):
        
        for row in self.worksheet.get_all_records():
            # TODO
    

    def get_all_tovaritelnici(self, column_name):
        tovaritelnici = []
        for cell in self.worksheet.col_values(self.get_column_number(column_name)):
            tovaritelnici.append(cell)
        return tovaritelnici

    def get_column_number(self, column_name):
        for i, header in enumerate(self.worksheet.row_values(1), start=1):
            if header == column_name:
                return i
        return None

    