import requests
import json
from datetime import datetime, timedelta, timezone

class SpeedyService:
    def __init__(self, username, password, base_url="https://api.speedy.bg/v1", language="BG"):
        self.username = username
        self.password = password
        self.BASE_URL = base_url
        
        self.language = language

    def get_shipment_info(self, shipment_id):
        url = f"{self.BASE_URL}/shipment/info"
        paylaod = {
            "userName": self.username,
            "password": self.password,
            'language': self.language,
            'shipmentIds': [shipment_id]
        }
        
        response = requests.post(url, json=paylaod, headers={"Content-Type": "application/json"})
        return response.json()
    
    def track(self, shipment_id):
        url = f"{self.BASE_URL}/track"
        paylaod = {
            "userName": self.username,
            "password": self.password,
            'language': self.language,
            'parcels': [
                {
                    'id': shipment_id
                }
            ]
        }
        
        response = requests.post(url, json=paylaod, headers={"Content-Type": "application/json"})
        return response.json()


    def is_paid(self, shipment_id):
        shipment_info = self.get_shipment_info(shipment_id)


    def calc_time(self, days_ago):
        time_point = datetime.now() - timedelta(days=days_ago)
        return time_point.strftime('%Y-%m-%d %H:%M')
    
    def format_speedy_date(self,date_string: str) -> str:
        """
        Converts a date string from Speedy's API format to 'dd/mm/yyyy'.

        Args:
            date_string: A string in the format 'YYYY-MM-DDTHH:MM:SS+ZZZZ',
                        for example, '2025-11-05T08:25:44+0200'.

        Returns:
            A formatted date string as 'dd/mm/yyyy', or an empty string
            if the input is invalid.
        """
        if not date_string:
            return ""
            
        try:
            
            datetime_object = datetime.strptime(date_string, "%Y-%m-%dT%H:%M:%S%z")
            return datetime_object.strftime("%d/%m/%Y")
            
        except (ValueError, TypeError):
            print(f"Warning: Could not parse the date string: '{date_string}'")
            return ""