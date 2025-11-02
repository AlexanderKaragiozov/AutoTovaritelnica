class Processor:
    def __init__(self, speedy_service, sheet_service):
        self.speedy_service = speedy_service
        self.sheet_service = sheet_service

    def process_all_tovaritelnici(self):
        for tovaritelnica in self.tovaritelnici:
            if self.get_order_status(tovaritelnica) != 'УСПЕШНО':
                status , time = self.speedy_service.get_order_status(tovaritelnica)
                self.sheet_service.update_order_status(tovaritelnica, status,time)

    def run(self):
        self.worksheet = self.sheet_service._connect(self.sheet_service.sheet_id, self.sheet_service.sheet_name)
        self.tovaritelnici = self.sheet_service.get_all_tovaritelnici('Товарителница')

    

