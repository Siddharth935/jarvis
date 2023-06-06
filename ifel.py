import pyttsx3
import pywhatkit
import speech_recognition as sr
import os
from requests import get
from pytube import YouTube
import instaloader
import webbrowser
import time
from pywikihow import search_wikihow

# opencv coding imports
import numpy as np
import HandTrackingModule as htm
import math
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import cv2
import speedtest


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)
        print("Say that again please...")
        return "None"
    return query

def openGoogle():
    speak("Bro, what should i search on google")
    cm = takeCommand().lower()
    webbrowser.open(f"{cm}")


def openyoutube():
    speak("Bro, what should i search on youtube")
    cm = takeCommand().lower()
    pywhatkit.playonyt(f"{cm}")

def openDownloads():
    dow = "C:\\Users\\Lenovo"
    os.startfile(dow)

def openWord():
    word = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Word 2016"
    os.startfile(word)


def openVsCode():
    codePath = "C:\\Users\\Lenovo\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
    os.startfile(codePath)

def openAndroidStudio():
    android = "C:\\Program Files\\Android\Android Studio\\bin\\studio64.exe"
    os.startfile(android)

def openVsCode():
    codePath = "C:\\Users\\Lenovo\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
    os.startfile(codePath)


def openExcel():
    excel = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Excel 2016"
    os.startfile(excel)

def closeAndroidStudio():
    speak("Okay sir closing android studio")
    os.system("taskkill /f /im studio64.exe")


def closeVsCode():
    speak("okay sir closing vs code")
    os.system("taskkill /f /im Code.exe")



def openPowerPoint():
    point = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\PowerPoint 2016"
    os.startfile(point)


def openWord():
    word = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Word 2016"
    os.startfile(word)

def openFiles():
    files = "C:\\"
    os.startfile(files)


def closeExecl():
    speak("okay sir closing excel")
    os.system("taskkill /f /im Excel 2016")

def closeVsCode():
    speak("okay sir closing vs code")
    os.system("taskkill /f /im Code.exe")

def closeWord():
    speak("okay sir closing word")
    os.system("taskkill /f /im Word 2016")

def closePowerPoint():
    speak("okay sir closing powerpoint")
    os.system("taskkill /f /im PowerPoint 2016")

def closePycharm():
    speak("Okay sir closing pycharm")
    os.system("taskkill /f /im pycharm64.exe")

def closeWord():
    speak("okay sir closing word")
    os.system("taskkill /f /im Word 2016")

def ipAddress():
    ip = get('https://api.ipify.org').text
    speak(f"Your ip address is {ip}")
    print("your IP address is",ip)


def youtube_downloder():
    speak("bro please entre link")
    url = input('Enter Link Below\n')
    my_video = YouTube(url)
    speak("I get that video would you like download")
    condition = takeCommand().lower()
    if "yes" in condition:
        my_video = my_video.streams.get_highest_resolution()
        my_video.download()
    else:
        pass

def insta():
    speak("sir please entre user name correctly")
    name = input("Enter user name\n")
    webbrowser.open(f"www.instagram.com/{name}")
    speak(f"sir here is a profile of the user{name}")
    time.sleep(4)
    speak("sir would you like to download profile picture of this account.")
    condition = takeCommand().lower()
    if "yes" in condition:
        mod = instaloader.Instaloader()
        mod.download_profile(name, profile_pic=True)
        speak("i am done sir, profile pictuire is saved in main folder.")
    else:
        pass

def volumeUpDown():
    Wcam, Hcam = 640, 480
    cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
    cap.set(3, Wcam)
    cap.set(4, Hcam)

    detector = htm.handDetector(detectionCon=0.7)

    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volRange = volume.GetVolumeRange()
    print(volume.GetVolumeRange())
    minVol = volRange[0]
    maxVol = volRange[1]
    val = 0
    volBar = 400
    volPer = 0

    while True:
        success, img = cap.read()
        img = detector.findHands(img)
        lmList = detector.findPosition(img, draw=False)
        if len(lmList) != 0:

            # x1, y1 = lmList[4][First Elemant] and, lmList[4][Second element]
            x1, y1 = lmList[4][1], lmList[4][2]
            x2, y2 = lmList[8][1], lmList[8][2]
            cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

            cv2.circle(img, (x1, y1), 10, (0, 0, 255), cv2.FILLED)
            cv2.circle(img, (x2, y2), 10, (0, 0, 255), cv2.FILLED)
            cv2.circle(img, (cx, cy), 10, (0, 0, 255), cv2.FILLED)
            cv2.line(img, (x1, y1), (x2, y2), (0, 0, 255), 3)

            length = math.hypot(x2 - x1, y2 - y1)

            # Hand range 50 - 300
            # Volume range -65 - 0

            vol = np.interp(length, [50, 300], [minVol, maxVol])
            volBar = np.interp(length, [50, 300], [400, 150])
            volPer = np.interp(length, [50, 300], [0, 100])
            volume.SetMasterVolumeLevel(vol, None)

            if length < 50 or length > 280:
                cv2.circle(img, (cx, cy), 10, (0, 255, 0), cv2.FILLED)

        cv2.rectangle(img, (50, 150), (85, 400), (0, 255, 0), 3)
        cv2.rectangle(img, (50, int(volBar)), (85, 400), (0, 255, 0), cv2.FILLED)
        cv2.putText(img, f'{int(volPer)} %', (40, 450), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
        cv2.imshow("Video", img)
        k = cv2.waitKey(1)
        if k == ord('q'):
            break

    cv2.destroyAllWindows()

def mediaPlayer():
    import cv2
    import pyautogui as p
    import HandTrackingModule as htm
    import time

    wSrn, hSrn = 640, 480
    cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
    cap.set(3, wSrn)
    cap.set(4, hSrn)
    pTime = 0
    tipIds = [4, 8, 12, 16, 20]

    detecor = htm.handDetector(maxHands=2, detectionCon=0.5)

    while True:
        success, img = cap.read()
        img = detecor.findHands(img)
        lmList = detecor.findPosition(img, draw=False)

        if len(lmList) != 0:
            x0, y0 = lmList[4][1], lmList[4][2]
            x1, y1 = lmList[8][1], lmList[8][2]

            fingers = []
            # Thumb
            if lmList[tipIds[0]][1] > lmList[tipIds[0] - 1][1]:
                fingers.append(1)
            else:
                fingers.append(0)
            # four finger
            for id in range(1, 5):
                if lmList[tipIds[id]][2] < lmList[tipIds[id] - 2][2]:
                    fingers.append(1)
                else:
                    fingers.append(0)

            # print(fingers)
            totalFinger = fingers.count(1)
            # print(totalFinger)

            if totalFinger == 5:
                cv2.circle(img, (x1, y1), 15, (0, 0, 255), cv2.FILLED)
                p.press("space")
            elif totalFinger == 4:
                p.press("right")
            elif totalFinger == 3:
                p.press("left")
            elif totalFinger == 2:
                p.press("volumeup")
            elif totalFinger == 1:
                p.press("volumedown")

        cTime = time.time()
        fps = 1 / (cTime - pTime)
        pTime = cTime

        cv2.putText(img, str(int(fps)), (30, 70), cv2.FONT_HERSHEY_SIMPLEX, 2, (0, 0, 255), 3)

        cv2.imshow("Video", img)
        k = cv2.waitKey(1)
        if k == ord('q'):
            break

    cv2.destroyAllWindows()

def internet():
    import speedtest
    st = speedtest.Speedtest()
    dl = st.download()
    up = st.upload()

    speak(f"bro we have {dl} bit per second downloading speed and {up} bit per second uploading speed")

def mod():
    speak("mod is activated bro. tell me what you want to know")
    how = takeCommand()
    max_result = 1
    how_to = search_wikihow(how, max_result)
    assert len(how_to) == 1
    how_to[0].print()
    speak(how_to[0].summary)


# def exitJarvis():
#     speak("Thank sir. have a good day.")
#     sys.exit()



