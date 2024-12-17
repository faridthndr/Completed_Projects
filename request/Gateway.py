import requests

def send_web_request(gateway_number):
    # Define the URL you want to access
    url = "https://digital.gov.eg/"

    # Set the proxy settings based on the gateway number
    if gateway_number == 1:
        default_gateway = "10.24.40.200"
    elif gateway_number == 2:
        default_gateway = "10.10.172.1"
    else:
        raise ValueError("Invalid gateway number. Must be 1 or 2.")

    # Create a new session object
    session = requests.Session()

    # Set the proxy settings to use the specified default gateway
    session.proxies = {
        "http": f"http://{default_gateway}:80",
        "https": f"https://{default_gateway}:443"
    }

    # Send a request to the URL using the session object
    response = session.get(url)

    # Check the status code of the response
    if response.status_code == 200:
        print("Request successful!")
    else:
        print(f"Request failed with status code: {response.status_code}")

# Example usage
send_web_request(1)  # Use the first gateway (10.24.40.200)
# send_web_request(2)  # Use the second gateway (10.10.172.1)