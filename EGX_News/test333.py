import requests

def get_user_info(email):
    url = f'https://usertool-internal.thndr.app/users?q={email}'
    response = requests.get(url)

    if response.status_code == 200:
        print(response.text)  
        try:
            data = response.json()  
            return data
        except requests.exceptions.JSONDecodeError:
            return "Received response is not in JSON format"
    else:
        return f'Failed to get user info. Status code: {response.status_code}'

email = 'farid.shawky@thndr.app'
user_info = get_user_info(email)
print(user_info)
