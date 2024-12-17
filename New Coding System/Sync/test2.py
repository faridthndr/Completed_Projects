import boto3
import json
import os
from datetime import datetime, timezone
from botocore.exceptions import ClientError

def get_credentials_from_file(profile_name):
    config_path = "Data/config"
    credentials_path = "Data/cli.json"
    sso_cache_path = "Data/sso.json"
    client_info_path = "Data/client_info.json"

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

    with open(credentials_path, 'r') as credentials_file:
        credentials_data = json.load(credentials_file)
        access_key_id = credentials_data['Credentials']['AccessKeyId']
        secret_access_key = credentials_data['Credentials']['SecretAccessKey']
        session_token = credentials_data['Credentials']['SessionToken']
    
    with open(sso_cache_path, 'r') as sso_file:
        sso_data = json.load(sso_file)
        region = sso_data['region']
        expires_at = sso_data['expiresAt']

    with open(client_info_path, 'r') as client_info_file:
        client_info = json.load(client_info_file)
        client_id = client_info['clientId']
        client_secret = client_info['clientSecret']


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
    folder_name = '2024-11-22_at_13_b8962f18-03cc-4438-99d4-eba42a3a519b'  
    target_folder = './New folder'  
    os.makedirs(target_folder, exist_ok=True)
    download_files_from_s3(folder_name, target_folder)
