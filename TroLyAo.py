# -*- coding: utf-8 -*-
"""
Created on Thu Nov 19 08:32:44 2020

@author: MINH THU
"""
import os
import playsound
import speech_recognition as sr
import time
import ctypes
import wikipedia
import datetime
import json
import re
import ctypes
import random
import urllib.request
import sys
import webbrowser
import smtplib
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from time import strftime
from gtts import gTTS
from youtube_search import YoutubeSearch
from email.mime.multipart import MIMEMultipart
from newspaper import Article

#%% ngon ngu
wikipedia.set_lang('vi')
language = 'vi'
path = ChromeDriverManager().install()

#%%
def speak(text):
    print("Trợ lý ảo: {}".format(text))
    tts = gTTS(text=text, lang=language, slow=False,)
    tts.save("sound.mp3")
    playsound.playsound("sound.mp3", False)
    os.remove("sound.mp3")

#%%
def get_audio():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Tôi: ", end='')
        audio = r.record(source, duration=5)
        try:
            text = r.recognize_google(audio, language="vi-VN")
            print(text)
            return text
        except:
            print("...")
            return 0

#%%
def stop():
    speak("Hẹn gặp lại bạn sau!")
    
#%%
def get_text():
    for i in range(3):
        text = get_audio()
        if text:
            return text.lower()
        elif i < 2:
            speak("Tôi không nghe rõ. Bạn nói lại được không!")
    time.sleep(2)
    stop()
    return 0
#%%
def sing():  
    speak("bạn muốn nghe thể loại nào?")
    #text = get_text()
    #if 'rap' in text:
    time.sleep(3)
    speak("Tôi chỉ mới biết hát chúc mừng sinh nhật thôi bạn nghe đỡ nhe hihi")
    time.sleep(6)
    text = """Mừng ngày sinh nhật đóa hoa mừng ngày sinh một khúc ca
          Mừng ngày đã sinh cho cuộc đời một bông hoa xinh rực rỡơ
          Cuộc đời em là đóa hoa cuộc đời em là khúc ca
          Cuộc đời sẽ thêm tươi đẹp vì những khúc ca và đóa hoa."""   
    time.sleep(2)
    playsound.playsound("HPBD.mp3", False)
    tts = gTTS(text=text, lang=language, slow=True,)
    tts.save("sound.mp3")
    playsound.playsound("sound.mp3", True)
    os.remove("sound.mp3")


#%%
def hello(name):
    day_time = int(strftime('%H'))
    if day_time < 12:
        speak("Chào buổi sáng bạn {}. Chúc bạn một ngày tốt lành.".format(name))
    elif 12 <= day_time < 18:
        speak("Chào buổi chiều bạn {}. Bạn đã dự định gì cho chiều nay chưa.".format(name))
    else:
        speak("Chào buổi tối bạn {}. Bạn đã ăn tối chưa nhỉ.".format(name))
        
#%%
def get_time(text):
    now = datetime.datetime.now()
    if "giờ" in text:
        speak('Bây giờ là %d giờ %d phút' % (now.hour, now.minute))
    elif "ngày" in text:
        speak("Hôm nay là ngày %d tháng %d năm %d" %
              (now.day, now.month, now.year))
    else:
        speak("Tôi chưa hiểu ý của bạn. Bạn nói lại được không?")
        
#%%
def open_application(text):
    if "google" in text:
        speak("Mở Google Chrome")
        os.startfile(
            r'C:\Program Files\Google\Chrome\Application\chrome.exe')
    elif "word" in text:
        speak("Mở Microsoft Word")
        os.startfile(
            r'C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE')
    elif "excel" in text:
        speak("Mở Microsoft Excel")
        os.startfile(
            r'C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE')
    elif "powerpoint" in text or "thuyết trình" in text:
        speak("Mở PowerPoint")
        os.startfile(
            r'C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE')     
    else:
        speak("Ứng dụng chưa được cài đặt. Bạn hãy thử lại!")
        
#%%
def open_website(text):
    reg_ex = re.search('mở (.+)', text)
    if reg_ex:
        domain = reg_ex.group(1)
        url = 'https://www.' + domain
        webbrowser.open(url)
        speak("Trang web bạn yêu cầu đã được mở.")
        return True
    else:
        return False
    
#%%
def open_google_and_search(text):
    search_for = text.split("kiếm", 1)[1]
    speak('Okay!')
    driver = webdriver.Chrome(path)
    driver.get("https://www.google.com")
    que = driver.find_element_by_xpath("//input[@name='q']")
    que.send_keys(str(search_for))
    que.send_keys(Keys.RETURN)
    
#%%
def send_email(text):
    speak('Bạn gửi email cho ai nhỉ')
    recipient = get_text()
    if 'minh' in recipient:
        speak('Nội dung bạn muốn gửi là gì')
        time.sleep(3)
        content = get_text()
        receivers = ['ngchai2410@gmail.com', '18128062@student.hcmute.edu.vn', 'tanleit151@gmail.com','18110361@student.hcmute.edu.vn']
        for i in receivers:
            msg = MIMEMultipart()
            msg['From'] = 'minhthuthum@gmail.com'      
            mail = smtplib.SMTP('smtp.gmail.com',587)
            mail.ehlo()
            mail.starttls()
            mail.login('minhthuthum@gmail.com', '0975226327')
            msg['To'] = i
            mail.sendmail(msg['From'], msg['To'], content.encode('utf-8'))  
        mail.close()
        speak('Email của bạn vùa được gửi. Bạn check lại email nhé hihi.')
    else:
        speak('Tôi không hiểu bạn muốn gửi email cho ai. Bạn nói lại được không?')
#%%
def current_weather():
    speak("Bạn muốn xem thời tiết ở đâu ạ.")
    base_url = "http://api.openweathermap.org/data/2.5/weather?"
    time.sleep(2)
    city = get_text()
    if not city:
        pass
    api_key = "5f9242d740d1e6278e390d5fd21a3881"
    call_url = base_url + "appid=" + api_key + "&q=" + city 
    response = requests.get(call_url)
    data = response.json()
    if data["cod"] != "404":
        city_res = data["main"]
        current_temperature = city_res["temp"]
        current_pressure = city_res["pressure"]
        current_humidity = city_res["humidity"]
        
        wthr = data["weather"]
        weather_description = wthr[0]["description"]
        now = datetime.datetime.now()
        content = """
        Hôm nay là ngày {day} tháng {month} năm {year}
        Nhiệt độ trung bình là {temp} độ Kelvin
        Áp suất không khí là {pressure} héc tơ Pascal
        Độ ẩm là {humidity}%
        Trời hôm nay quang mây. Dự báo mưa rải rác ở một số nơi.""".format(day = now.day,month = now.month, year= now.year, 
                                                                           temp = current_temperature, pressure = current_pressure, humidity = current_humidity)
        speak(content)
        time.sleep(20)
    else:
        speak("Tôi không tìm thấy địa chỉ nơi bạn muốn biết thời tiết")

#%%
def play_song():
    speak('Xin mời bạn chọn tên bài hát')
    time.sleep(2)
    mysong = get_text()
    while True:
        result = YoutubeSearch(mysong, max_results=10).to_dict()
        if result:
            break
    url = 'https://www.youtube.com' + result[0]['url_suffix']
    webbrowser.open(url)
    speak("Bài hát bạn yêu cầu đã được mở.")
    time.sleep(2)
    speak("Chúc bạn nghe nhạc vui vẻ!")
    
#%%
def change_wallpaper():
    
    speak('Oke bạn iu tui sẽ đổi ngay nè')    
    img1 = f"D:/HocTap/TT/DoAn/img1.jpg"
    img2 = f"D:/HocTap/TT/DoAn/img2.jpg"
    img3 = f"D:/HocTap/TT/DoAn/img3.jpg"
    img4 = f"D:/HocTap/TT/DoAn/img4.jpg"
    img5 = f"D:/HocTap/TT/DoAn/img5.jpg"
    img6 = f"D:/HocTap/TT/DoAn/img6.jpg"
    img7 = f"D:/HocTap/TT/DoAn/img7.jpg"
    img8 = f"D:/HocTap/TT/DoAn/img8.jpg"
    img9 = f"D:/HocTap/TT/DoAn/img9.jpg"
    img10 = f"D:/HocTap/TT/DoAn/img10.jpg"
    img11 = f"D:/HocTap/TT/DoAn/img11.jpg"
    while True:
        
        img = [img1, img2, img3, img4, img5, img6, img7, img8, img9, img10, img11]
        num = random.randrange(0, 10, 1)
        print(num)
        ctypes.windll.user32.SystemParametersInfoW(20, 0, img[num], 0)
        
        speak("Đã đổi rồi nè bạn có thích không?")
        text = get_text()
        if 'có' in text or 'được' in text or 'ok' in text:
            break
        
#%%
def read_news():
    
    speak("Bạn muốn đọc báo về gì")
    
    queue = get_text()
    #queue = "Trump"
    params = {
        'apiKey': '9007bba8aca54b379ccba2964abac887',
        'q': queue,
    }
    api_result = requests.get('http://newsapi.org/v2/top-headlines?', params)
    api_response = api_result.json()
    print("Tin tức")

    for number, result in enumerate(api_response['articles'], start=1):
        print(f"""Tin {number}:\nTiêu đề: {result['title']}\nTrích dẫn: {result['description']}\nLink: {result['url']}\nNội dung: {result['content']}
    """)
        if number <= 3:
            webbrowser.open(result['url'])
        #text = (result['content'])
        #speak(text)
        #break
        #article = Article(result['url'])
        #article.download()
        #article.parse()
        

#%%
def tell_me_about():
    try:
        speak("Bạn muốn nghe về gì ạ")
        time.sleep(2)
        text = get_text()
        contents = wikipedia.summary(text).split('\n')
        speak(contents[0])
        time.sleep(30)
        for content in contents[1:]:
            speak("Bạn muốn nghe thêm không")
            ans = get_text()
            if "có" not in ans:
                break    
            speak(content)
            time.sleep(30)
        time.sleep(3)
        speak('Cảm ơn bạn đã lắng nghe!!!')
    except:
        speak("Tôi không định nghĩa được thuật ngữ của bạn. Xin mời bạn nói lại")
        
#%%
def help_me():
    speak("""Tôi có thể giúp bạn thực hiện các câu lệnh sau đây:
    1. Chào hỏi
    2. Hiển thị giờ
    3. Mở website, application
    4. Tìm kiếm trên Google
    5. Gửi email
    6. Dự báo thời tiết
    7. Mở video nhạc
    8. Thay đổi hình nền máy tính
    9. Đọc báo hôm nay
    10. Kể bạn biết về thế giới 
    11. Hát khá hay """)
    time.sleep(22)

#%%
def assistant():
    speak("Xin chào, bạn iu tên là gì nhỉ?")
    time.sleep(2)
    name = get_text()
    if name:
        speak("Chào bạn {}".format(name))
        time.sleep(2)
        speak("Bạn iu cần tôi giúp gì ạ?")
        while True:
            time.sleep(3)
            text = get_text()
            if not text:
                break
            elif "dừng" in text or "tạm biệt" in text or "tạm biệt robot" in text or "ngủ thôi" in text:
                stop()
                break
            elif "có thể làm gì" in text or 'làm được gì' in text:
                help_me()
            elif "chào trợ lý ảo" in text or "Xin chào" in text:
                hello(name)
            elif "hiện tại" in text or "mấy giờ" in text:
                get_time(text)
            elif "mở" in text:
                if 'mở google và tìm kiếm' in text:
                    open_google_and_search(text)
                elif "." in text:
                    open_website(text)
                else:
                    open_application(text)
            elif "email" in text or "mail" in text or "gmail" in text:
                send_email(text)
            elif "thời tiết" in text:
                current_weather()
            elif "chơi nhạc" in text:
                play_song()
                break
            elif "hình nền" in text:
                change_wallpaper()
            elif "đọc báo" in text:
                read_news()
            elif "định nghĩa" in text or "cho tôi biết" in text or "tôi muốn biết" in text:
                tell_me_about()
            elif "hát" in text:
                sing()
            else:
                speak("Bạn cần tôi giúp gì ạ?")
#%%
assistant()