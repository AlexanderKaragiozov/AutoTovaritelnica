import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import CellFormat, Color, format_cell_ranges
CODE_DICT = {
    -14: "ПЛАТЕНА",
    115: "ПРЕНАСОЧЕНА",
    116: "ПРЕПРАТЕНА",
    111: "ВЪРНАТА",
    124: "ВЪРНАТА",
}

class SheetService:
    def __init__(self, sheet_id, sheet_name):
        self.sheet_id = sheet_id
        self.sheet_name = sheet_name
        self._headers = None
        self._column_cache = {}
        self.worksheet = None

    # ------------------ CONNECTION ------------------

    def _connect(self):
        if self.worksheet:
            return self.worksheet
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        credentials = ServiceAccountCredentials.from_json_keyfile_name('key/key.json', scope)
        client = gspread.authorize(credentials)
        sheet = client.open(self.sheet_name)
        self.worksheet = sheet.get_worksheet(0)
        return self.worksheet

    # ------------------ FILE I/O ------------------

    def get_last_treated_row(self):
        with open('last_treated_row.txt', 'r') as f:
            return int(f.read())

    def set_last_treated_row(self, last_row):
        with open('last_treated_row.txt', 'w') as f:
            f.write(str(last_row))

    # ------------------ HEADERS & CACHING ------------------

    def _get_headers(self):
        if self._headers is None:
            self._headers = self.worksheet.row_values(1)
        return self._headers

    def _get_column_number(self, column_name):
        headers = self._get_headers()
        try:
            return headers.index(column_name) + 1
        except ValueError:
            raise Exception(f"Column '{column_name}' not found")

    def _get_column_values(self, column_name):
        """Cache column values to avoid quota hits."""
        if column_name not in self._column_cache:
            col_num = self._get_column_number(column_name)
            self._column_cache[column_name] = self.worksheet.col_values(col_num)
        return self._column_cache[column_name]

    def refresh_cache(self):
        """Manually clear cached data when sheet changes."""
        self._headers = None
        self._column_cache.clear()

    # ------------------ LOOKUPS ------------------

    def get_row_number(self, tovaritelnica):
        """Get row by value in 'Товарителница' column (cached)."""
        col_values = self._get_column_values('Товарителница')
        for i, cell in enumerate(col_values, start=1):
            if cell == str(tovaritelnica):
                return i
        return None

    def get_whole_row(self, tovaritelnica):
        row_number = self.get_row_number(tovaritelnica)
        return self.worksheet.row_values(row_number) if row_number else None

    # ------------------ MAIN LOGIC ------------------

    def get_all_tovaritelnici(self, column_name, start_row=None):
        if start_row is None:
            start_row = self.get_last_treated_row()

        col_index = self._get_column_number(column_name)
        col_values = self.worksheet.col_values(col_index)[start_row - 1:]

        if col_values:
            last_row = col_values[-1]
            if last_row:
                self.set_last_treated_row(self.get_row_number(last_row) + 1)

        return [c for c in col_values if c.isdigit()]

    def get_only_needed_orders(self):
        last_row = self.get_last_treated_row()
        all_data = self.worksheet.get_all_values()

        headers = all_data[0]
        data_rows = all_data[1:]
        try:
            status_index = headers.index('Статус')
        except ValueError:
            print("Error: Could not find 'Статус' column.")
            return []

        start_index = last_row - 2
        needed_orders = []

        for i, row_values in enumerate(data_rows[start_index:], start=last_row):
            if not any(row_values):
                continue
                    #'', 'Препратена', 'Изпратена',
            status = row_values[status_index].strip()
            if status.lower() not in ['платена', 'върната', 'отказана']:
                order = dict(zip(headers, row_values))
                order['row_number'] = i
                if order.get('Товарителница', '').isdigit():
                    needed_orders.append(order['Товарителница'])

        return needed_orders

    # ------------------ UPDATES ------------------

    def update_order_status(self, tovaritelnica, status):
        """Single update (use sparingly)"""
        row = self.get_row_number(tovaritelnica)
        if row:
            self.worksheet.update_cell(row, 10, status)

    def bulk_update_cells(self, column_name, updates):
        """
        updates: list of tuples (row_number, value)
        Example: [(2, 'ABC'), (5, 'XYZ')]
        """
        from gspread.cell import Cell

        col_number = self._get_column_number(column_name)
        cells = [Cell(row=row, col=col_number, value=value) for row, value in updates]
        self.worksheet.update_cells(cells)

    def col_num_to_letter(self, col_num):
        """ Преобразува номер на колона (1-базиран) в нейната буква (напр. 1 -> A, 27 -> AA). """
        letter = ''
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            letter = chr(65 + remainder) + letter
        return letter

    def color_cell_statuses(self):
        """
        Оцветява клетките в колона 'Статус', като първо намира колоната
        автоматично по нейното име. Използва една batch заявка.
        """
        # --- Стъпка 1: Автоматично намиране на колоната 'Статус' ---
        status_column_name = 'Статус'
        try:
            headers = self.worksheet.row_values(1)
            column_index = headers.index(status_column_name)
            column_number = column_index + 1
            status_column_letter = self.col_num_to_letter(column_number)
            print(f"Колоната '{status_column_name}' беше намерена: колона {status_column_letter}")
        except ValueError:
            print(f"ГРЕШКА: Не мога да намеря колона с име '{status_column_name}' в документа.")
            return

        # --- Стъпка 2: Дефинираме карта на статусите и цветовете ---
        # Използваме речник за бърз и лесен достъп
        COLOR_MAP = {
            'ВЪРНАТА': Color(255/255, 0/255, 0/255),           # #ff0000 (Red)
            'ПРЕПРАТЕНА': Color(255/255, 112/255, 0/255),      # #ff7000 (Orange)
            'ИЗПРАТЕНО ИЗВЕСТИЕ': Color(204/255, 255/255, 0/255), # #ccff00 (Lime Green)
            'ПЛАТЕНА': Color(0/255, 255/255, 0/255)            # #00ff00 (Bright Green)
        }
        
        # Дефинираме цвета по подразбиране за всички останали статуси
        DEFAULT_COLOR = Color(240/255, 243/255, 220/255) # #f0f3dc

        # --- Стъпка 3: Извличаме всички стойности от намерената колона ---
        try:
            # Взимаме стойностите от 2-ри ред до края
            all_status_values = [item[0] if item else '' for item in self.worksheet.get(f'{status_column_letter}2:{status_column_letter}')]
        except IndexError:
            print(f"Колоната '{status_column_name}' е празна или не може да бъде прочетена.")
            return

        # --- Стъпка 4: Подготвяме списък с форматирания за batch update ---
        formats_to_apply = []
        
        for i, status_str in enumerate(all_status_values):
            row_number = i + 2  # Цикълът започва от i=0, което е ред 2 в документа
            cell_range = f'{status_column_letter}{row_number}'
            
            # Почистваме текста от излишни интервали и го правим с главни букви
            # за по-надеждно сравнение
            cleaned_status = status_str.strip().upper()
            
            # .get() е идеален тук: взима цвета от картата, ако статусът съществува,
            # в противен случай използва DEFAULT_COLOR.
            color_to_set = COLOR_MAP.get(cleaned_status, DEFAULT_COLOR)
            
            # Създаваме обект за форматиране и го добавяме към списъка
            cell_format = CellFormat(backgroundColor=color_to_set)
            formats_to_apply.append((cell_range, cell_format))

        # --- Стъпка 5: Изпращаме една-единствена batch заявка за оцветяване ---
        if formats_to_apply:
            print(f"Прилагане на {len(formats_to_apply)} форматирания за статуси...")
            format_cell_ranges(self.worksheet, formats_to_apply)
            print("Оцветяването на статусите завърши.")
        else:
            print("Няма статуси за оцветяване.")

   

    def color_cell_days(self):
        """
        Оцветява клетките в колона 'Дни', като първо намира колоната
        автоматично по нейното име. Използва една batch заявка за всички промени.
        """
        days_column_name = 'Дни'
        try:
            headers = self.worksheet.row_values(1)
            column_index = headers.index(days_column_name)
            column_number = column_index + 1
            days_column_letter = self.col_num_to_letter(column_number)
            print(f"Колоната '{days_column_name}' беше намерена: колона {days_column_letter}")
        except ValueError:
            print(f"ГРЕШКА: Не мога да намеря колона с име '{days_column_name}' в документа.")
            return

        # --- Стъпка 1: Дефинираме цветовете и условията ---
        # Редът в този списък е важен! Проверките ще се правят отгоре надолу.
        COLOR_RULES = [
            {'name': 'white', 'color': Color(1, 1, 1), 'condition': lambda d: d == 0},  # Бяло за 0 дни
            {'name': 'yellow', 'color': Color(255/255, 243/255, 59/255), 'condition': lambda d: d < 2},   # Жълто за 1 ден
            {'name': 'orange_2_4', 'color': Color(253/255, 199/255, 12/255), 'condition': lambda d: d < 4},   # Оранжево за 2-3 дни
            {'name': 'orange_4_6', 'color': Color(243/255, 144/255, 63/255), 'condition': lambda d: d <= 6},  # Тъмно оранжево за 4-6 дни
            {'name': 'red_7', 'color': Color(237/255, 104/255, 60/255), 'condition': lambda d: d == 7},  # Червено за 7-мия ден
            {'name': 'red_8_plus', 'color': Color(233/255, 62/255, 58/255), 'condition': lambda d: d >= 8}   # Тъмно червено за 8+ дни
        ]
        
        # Дефинираме цвета по подразбиране (бял) за празни или невалидни клетки
        DEFAULT_COLOR = Color(1, 1, 1) # 1, 1, 1 е бяло

        # --- Стъпка 2: Извличаме всички стойности от намерената колона ---
        try:
            all_day_values = [item[0] if item else '' for item in self.worksheet.get(f'{days_column_letter}2:{days_column_letter}')]
        except IndexError:
            print(f"Колоната '{days_column_name}' е празна.")
            return

        # --- Стъпка 3: Подготвяме списък с форматирания ---
        formats_to_apply = []
        
        for i, day_str in enumerate(all_day_values):
            row_number = i + 2
            cell_range = f'{days_column_letter}{row_number}'
            final_color = DEFAULT_COLOR # По подразбиране цветът е бял

            try:
                days = int(day_str)
                # Обхождаме правилата и прилагаме първото, което отговаря на условието
                for rule in COLOR_RULES:
                    if rule['condition'](days):
                        final_color = rule['color']
                        break # Важно: спираме след първото съвпадение
            except (ValueError, TypeError):
                # Ако стойността е празен стринг или текст, final_color остава DEFAULT_COLOR
                pass

            # Добавяме форматиране за всяка клетка, дори и да е с бял цвят
            cell_format = CellFormat(backgroundColor=final_color)
            formats_to_apply.append((cell_range, cell_format))

        # --- Стъпка 4: Изпращаме една-единствена batch заявка за оцветяване ---
        if formats_to_apply:
            print(f"Прилагане на {len(formats_to_apply)} форматирания за цвят...")
            format_cell_ranges(self.worksheet, formats_to_apply)
            print("Оцветяването завърши.")
        else:
            print("Няма клетки за оцветяване.")

    def get_days_remaining(self, arrival_date):
        from datetime import datetime, timedelta

        returns_at = datetime.strptime(arrival_date, "%d/%m/%Y")
        days_to_add = 8
        added = 0

        while added < days_to_add:
            returns_at += timedelta(days=1)
            # weekday(): Monday=0 ... Sunday=6
            if returns_at.weekday() != 6:
                added += 1

        return returns_at.strftime("%d/%m/%Y")
