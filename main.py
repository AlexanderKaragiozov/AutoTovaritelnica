from variables import USERNAME,PASSWORD,SHEET_ID,SHEET_NAME
from sheet_service import SheetService
from speedy_service import SpeedyService
from processor import Processor
import datetime 



if __name__ == '__main__':
    sheet_service = SheetService(SHEET_ID,SHEET_NAME)
    speedy_service = SpeedyService(USERNAME,PASSWORD)
    
    
        
    
    processor = Processor(speedy_service, sheet_service)
    
    processor.run()
    # print(sheet_service.get_whole_row('63308045061')[11])
    # sheet_service.update_order_status('63308045061','')
    
            
                