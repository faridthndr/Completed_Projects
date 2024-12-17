import winsound  

# تحديد مسار ملف الصوت  
sound_file = "Alarm07.wav"  # تأكد من وضع المسار الصحيح للملف  

# تشغيل الملف الصوتي  
# winsound.PlaySound(sound_file, winsound.SND_FILENAME)

Start_line = 4
if isinstance(Start_line, str) or Start_line == "" :  
        Start_line = 2

print(Start_line)

