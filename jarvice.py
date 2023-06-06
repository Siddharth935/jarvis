import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import cv2
import pyautogui as p

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
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

def taskexectution():
    speak("verification successful")
    speak("welcome back siddharth brother")
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")
    speak("I am Jarvis. Please tell me how may I help you")

    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open google' in query:
            from ifel import openGoogle
            openGoogle()


        elif 'open youtube' in query:
            from ifel import openyoutube
            openyoutube()

        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")

        elif 'what the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        # all open comands
        elif 'open vs code' in query:
            from ifel import openVsCode
            openVsCode()

        elif 'open excel' in query:
            from ifel import openExcel
            openExcel()

        elif 'open word' in query:
            from ifel import openWord
            openWord()

        elif 'open powerpoint' in query:
            from ifel import openPowerPoint
            openPowerPoint()

        elif 'open files' in query:
            from ifel import openFiles
            openFiles()

        # elif 'open android studio' or 'open androidstudio'in query:
        #     from ifel import openAndroidStudio
        #     openAndroidStudio()
        #
        # elif 'close android studio' in query:
        #     from ifel import closeAndroidStudio
        #     closeAndroidStudio()

        #all close camands
        elif 'close android studio' in query:
            from ifel import closeAndroidStudio
            closeAndroidStudio()

        elif 'close vs code' in query:
            from ifel import closeVsCode
            closeVsCode()

        elif 'close excel' in query:
            from  ifel import  closeExecl
            closeExecl()

        elif 'close word' in query:
            from ifel import  closeWord
            closeWord()

        elif 'close powerpoint' in query:
            from ifel import closePowerPoint
            closePowerPoint()


        elif 'what is my ip address' in query:
            from ifel import ipAddress
            ipAddress()

        elif 'open downloads' in query:
            from ifel import openDownloads
            openDownloads()

        elif 'close pycharm' in query:
            from ifel import  closePycharm
            closePycharm()

        elif 'volume' in query:
            from ifel import volumeUpDown
            volumeUpDown()

        elif 'media player' in query:
            from ifel import mediaPlayer
            mediaPlayer()

        elif 'instagram profile' in query:
            from ifel import insta
            insta()

        # elif 'activate mod' or 'mod' in query:
        #     from ifel import mod
        #     mod()

        elif 'download video' in query:
            from ifel import youtube_downloder
            youtube_downloder()

        # elif 'no thanks' or 'no thank you' or 'no' in query:
        #     from ifel import exitJarvis
        #     exitJarvis()

        elif 'internet speed' or 'jarvis what is my internet speed' or 'jarvis whats my internet speed' or 'whats my internet speed' or 'my internet speed jarvis' or 'whats my net speed jarvis' in query:
            from ifel import  internet
            internet()

        speak("Brother, i am ready for other work")

if __name__ == "__main__":

    recognizer = cv2.face.LBPHFaceRecognizer_create()  # Local Binary Patterns Histograms
    recognizer.read('trainer/trainer.yml')  # load trained model
    cascadePath = "haarcascade_frontalface_default.xml"
    faceCascade = cv2.CascadeClassifier(cascadePath)  # initializing haar cascade for object detection approach

    font = cv2.FONT_HERSHEY_SIMPLEX  # denotes the font type

    id = 2  # number of persons you want to Recognize

    names = ['', 'siddharth']  # names, leave first empty bcz counter starts from 0

    cam = cv2.VideoCapture(0, cv2.CAP_DSHOW)  # cv2.CAP_DSHOW to remove warning
    cam.set(3, 640)  # set video FrameWidht
    cam.set(4, 480)  # set video FrameHeight

    # Define min window size to be recognized as a face
    minW = 0.1 * cam.get(3)
    minH = 0.1 * cam.get(4)

    # flag = True

    while True:

        ret, img = cam.read()  # read the frames using the above created object

        converted_image = cv2.cvtColor(img,
                                       cv2.COLOR_BGR2GRAY)  # The function converts an input image from one color space to another

        faces = faceCascade.detectMultiScale(
            converted_image,
            scaleFactor=1.2,
            minNeighbors=5,
            minSize=(int(minW), int(minH)),
        )

        for (x, y, w, h) in faces:

            cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)  # used to draw a rectangle on any image

            id, accuracy = recognizer.predict(converted_image[y:y + h, x:x + w])  # to predict on every single image

            # Check if accuracy is less them 100 ==> "0" is perfect match
            if (accuracy < 100):
                id = names[id]
                accuracy = "  {0}%".format(round(100 - accuracy))
                taskexectution()
            else:
                id = "unknown"
                accuracy = "  {0}%".format(round(100 - accuracy))
                speak("user authentication is failed")
                break

            cv2.putText(img, str(id), (x + 5, y - 5), font, 1, (255, 255, 255), 2)
            cv2.putText(img, str(accuracy), (x + 5, y + h - 5), font, 1, (255, 255, 0), 1)

        cv2.imshow('camera', img)

        k = cv2.waitKey(10) & 0xff  # Press 'ESC' for exiting video
        if k == 27:
            break

    # Do a bit of cleanup
    print("Thanks for using this program, have a good day.")
    cam.release()
    cv2.destroyAllWindows()