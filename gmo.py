

import requests
import json

endPoint = 'https://api.coin.z.com/public'
path     = '/v1/ticker?symbol=BTC'

response = requests.get(endPoint + path)
print(json.dumps(response.json(), indent=2))





