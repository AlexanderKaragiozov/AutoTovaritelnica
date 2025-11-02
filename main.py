from variables import API_KEY,APY_SECRET,SHEET_ID,SHEET_NAME
from sheet_service import SheetService
from speedy_service import SpeedyService
from processor import Processor


if __name__ == '__main__':
    sheet_service = SheetService(SHEET_ID,SHEET_NAME)
    speedy_service = SpeedyService(API_KEY,APY_SECRET)
    processor = Processor(sheet_service, speedy_service)