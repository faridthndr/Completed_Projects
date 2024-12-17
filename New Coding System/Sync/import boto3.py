import boto3  
import os

def download_folder_from_s3(folder_name, target_folder):  
    # إنشاء جلسة باستخدام ملف التعريف المحدد  
    session = boto3.Session(profile_name='DocumentService-598269941583')  
    s3 = session.client('s3')  # الآن نقوم بإنشاء عميل S3  

    bucket_name = 'thndr-coding-prod'  
    
    # تحميل جميع الكائنات من المجلد (prefix)  
    response = s3.list_objects_v2(Bucket=bucket_name, Prefix=f'{folder_name}/')  
    
    if 'Contents' in response:  
        for obj in response['Contents']:  
            key = obj['Key']  
            # استخدم المسار الكامل للملف لتنزيله  
            local_file_path = f"./{target_folder}/{key.split('/')[-1]}"  
            # تأكد من وجود المجلد الهدف  
            os.makedirs(os.path.dirname(local_file_path), exist_ok=True)  
            s3.download_file(bucket_name, key, local_file_path)  
            print(f'Downloaded {key} to {local_file_path}')  
    else:  
        print('No files found in this folder.')  

# استخدام الدالة مع اسم المجلد المستخرج من الـ API  
folder_name = 'اسم_المجلد_الذي_تلقيته'  # تأكد من وضع اسم المجلد الصحيح الذي تلقيته من الـ API  
target_folder = 'المجلد_الذي_تريد_تنزيل_الملفات_إليه'  
download_folder_from_s3('2024-11-20_at_10_6151156d-afad-46ac-887d-0ba3af89f317', './New folder')