import requests

file_id = '17VTB2IBtGwDFss-fRvmNQckWkWrS1SRQ'
url = f'https://drive.google.com/uc?export=download&id={file_id}'

output = 'downloaded_file.pdf'

response = requests.get(url, stream=True)

if response.status_code == 200:
    with open(output, 'wb') as file:
        for chunk in response.iter_content(chunk_size=8192):
            file.write(chunk)
    print(f'File downloaded successfully as {output}')
else:
    print(f'Failed to download file. Status code: {response.status_code}')
