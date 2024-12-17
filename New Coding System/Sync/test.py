import boto3
import json
import os
from botocore.exceptions import ClientError



def get_credentials_from_file(profile_name):
    script_dir = os.path.dirname(os.path.abspath(__file__))
   
    config_path = os.path.join(script_dir, ".aws", "config")
    cli_directory = os.path.join(script_dir, ".aws", "cli", "cache")
    sso_cache_directory = os.path.join(script_dir, ".aws", "sso", "cache")

    # قراءة بيانات الاعتماد من ملفات التعريف
    with open(config_path, 'r') as config_file:
        config_lines = config_file.readlines()
        sso_session = None
        found_profile = False

        for line in config_lines:
            if line.strip() == f"[profile {profile_name}]":
                found_profile = True
            elif found_profile and line.startswith("sso_session"):
                sso_session = line.split('=')[1].strip()
                break

        if sso_session is None:
            raise ValueError(f"Profile {profile_name} not found or sso_session not defined.")

    # العثور على ملف CLI JSON في الدليل المحدد
    cli_files = [f for f in os.listdir(cli_directory) if f.endswith('.json')]
    cli_file_path = None
    for cli_file in cli_files:
        with open(os.path.join(cli_directory, cli_file), 'r') as file:
            data = json.load(file)
            if data.get("ProviderType") == "sso":
                cli_file_path = os.path.join(cli_directory, cli_file)
                break
    
    if cli_file_path is None:
        raise FileNotFoundError("CLI JSON file with 'ProviderType': 'sso' not found in the specified directory.")

    with open(cli_file_path, 'r') as credentials_file:
        credentials_data = json.load(credentials_file)
        access_key_id = credentials_data['Credentials']['AccessKeyId']
        secret_access_key = credentials_data['Credentials']['SecretAccessKey']
        session_token = credentials_data['Credentials']['SessionToken']
        # print(credentials_data)
        # print(access_key_id)
        # print(secret_access_key)
        print(session_token)

    
    # العثور على ملفات SSO JSON في الدليل المحدد
    sso_files = os.listdir(sso_cache_directory)
    sso_data_file = None
    client_info_file = None

    for sso_file in sso_files:
        with open(os.path.join(sso_cache_directory, sso_file), 'r') as file:
            data = json.load(file)
            if "startUrl" in data:
                sso_data_file = sso_file
            elif "clientId" in data:
                client_info_file = sso_file

    if sso_data_file is None:
        raise FileNotFoundError("SSO data file with 'startUrl' not found.")
    
    if client_info_file is None:
        raise FileNotFoundError("Client info file with 'clientId' not found.")
    
    with open(os.path.join(sso_cache_directory, sso_data_file), 'r') as sso_file:
        sso_data = json.load(sso_file)
        region = sso_data['region']

    with open(os.path.join(sso_cache_directory, client_info_file), 'r') as client_info_file:
        client_info = json.load(client_info_file)
        client_id = client_info['clientId']
        client_secret = client_info['clientSecret']

    # تكوين جلسة Boto3
    session = boto3.Session(
        aws_access_key_id=access_key_id,
        aws_secret_access_key=secret_access_key,
        aws_session_token=session_token,
        region_name=region
    )

    return session

def download_files_from_s3(folder_name, target_folder):
    try:
        session = get_credentials_from_file('DocumentService-598269941583')
        s3 = session.resource('s3')
        bucket_name = 'thndr-coding-prod'
        bucket = s3.Bucket(bucket_name)
        
        for obj in bucket.objects.filter(Prefix=folder_name):
            target_path = os.path.join(target_folder, obj.key.split('/')[-1])
            bucket.download_file(obj.key, target_path)
            print(f"Downloaded {obj.key} to {target_path}")
    except ClientError as e:
        if e.response['Error']['Code'] == 'ExpiredToken':
            print("The provided token has expired. Please renew the token and try again.")
        else:
            raise

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    folder_name = '2024-11-22_at_13_b8962f18-03cc-4438-99d4-eba42a3a519b'
    target_folder = './New folder'
    os.makedirs(target_folder, exist_ok=True)
    download_files_from_s3(folder_name, target_folder)
