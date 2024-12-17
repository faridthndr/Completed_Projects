import requests  

client_id = 'YOUR_CLIENT_ID'  
client_secret = 'YOUR_CLIENT_SECRET'  
refresh_token = 'YOUR_REFRESH_TOKEN'  

url = 'https://slack.com/api/oauth.v2.access'  

data = {  
    'client_id': client_id,  
    'client_secret': client_secret,  
    'refresh_token': refresh_token,  
    'grant_type': 'refresh_token'  
}  

response = requests.post(url, data=data)  

if response.status_code == 200:  
    new_tokens = response.json()  
    if 'access_token' in new_tokens:  
        print("ok", new_tokens['access_token'])  
    else:  
        print("ok", new_tokens.get('error'))  
else:  
    print("error", response.status_code)

