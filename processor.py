import datetime
from speedy_service import SpeedyService
from sheet_service import SheetService
class Processor:
    def __init__(self, speedy_service, sheet_service):
        self.speedy_service : SpeedyService = speedy_service 
        self.sheet_service: SheetService = sheet_service

    def process_all_tovaritelnici(self, data):
        STATUS_MAP = {
            -14: 'ПЛАТЕНА',
            115: 'ПРЕНАСОЧЕНА',
            11: 'ПРИЕТА В ОФИС',
            116: 'ПРЕПРАТЕНА',
            111: 'ВЪРНАТА',
            1134: 'ИЗПРАТЕНО ИЗВЕСТИЕ'
        }
        
        status_updates = []
        day_updates = []
        todays_date = datetime.datetime.now().strftime("%d/%m/%Y")
        for waybill in data:
            data = self.speedy_service.track(waybill)
            status = ''
            day_duration = 0
            print("waybill: ", waybill)
            for parcel in data['parcels']:
                client_ready = False
                for operation in parcel['operations']:
                    
                        
                    if operation['operationCode'] == 134 and client_ready == False:
                        arrival_date = self.speedy_service.format_speedy_date(operation['dateTime'])
                        client_ready = True
                        
                    else:
                        arrival_date =''
                    if arrival_date:
                        day_duration = (datetime.datetime.strptime(todays_date, "%d/%m/%Y") - datetime.datetime.strptime(arrival_date, "%d/%m/%Y")).days
                        
                    
                    code = operation['operationCode']
                    description = operation['description']
                    try:
                        comment = operation['comment']
                    except KeyError:
                        comment = 'Няма коментар'
                    status = description.upper()
                    if code == -14:
                        arrival_date = ''
                        status = STATUS_MAP[code]
                        day_duration = 0
                        break
                    elif code == 111:
                        arrival_date = ''
                        status = STATUS_MAP[code]
                        day_duration = 0
                        break
                    else:
                        
                        try:
                            status = STATUS_MAP[code]
                        except KeyError:
                            status = description.upper()
                    
                    # elif operation['operationCode'] == 115:
                    #     status = 'ПРЕНАСОЧЕНА'
                    # elif operation['operationCode'] == 11:
                    #     status = 'ПРИЕТА В ОФИС'
                    # elif operation['operationCode'] == 116:
                    #     status = 'ПРЕПРАТЕНА'
                    # elif operation['operationCode'] == 111:
                    #     status = 'ВЪРНАТА'
                    # elif operation['operationCode'] == 1134:
                    #     status = 'ИЗПРАТЕНО ИЗВЕСТИЕ'
                    
                    
            if day_duration != 0:
                day_updates.append((self.sheet_service.get_row_number(waybill), day_duration))
                    
            status_updates.append((self.sheet_service.get_row_number(waybill), status))

        self.sheet_service.bulk_update_cells('Статус', status_updates)
        self.sheet_service.bulk_update_cells('Дни', day_updates)

   
            
            
          

    

    def run(self):
        self.worksheet = self.sheet_service._connect()
        print
        # self.tovaritelnici = self.sheet_service.get_all_tovaritelnici('Товарителница')
        data = self.sheet_service.get_only_needed_orders()
        self.process_all_tovaritelnici(data)
        self.sheet_service.color_cell_days()
        self.sheet_service.color_cell_statuses()
        
        # asd = self.speedy_service.track('63321900494')
        
        
       
    

