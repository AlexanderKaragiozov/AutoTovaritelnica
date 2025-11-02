
import requests,datetime
class SpeedyService:
    def __init__(self,api,secret,):
        self.api = api
        self.secret = secret

    def get_order_status(self, tovaritelnica):
        request = requests.get(f'https://api.speedy.bg/api/v2/orders/{tovaritelnica}?api_key={self.api}&api_secret={self.secret}')
        
        return request.json()['status'] , self.calc_time(request.json()['created_at'])

    def calc_time(self, date):
        today = datetime.today()
        time_delta = today - datetime.timedelta(days=date)
        return time_delta.strftime('%Y-%m-%d %H:%M')


