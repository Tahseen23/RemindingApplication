import pandas as pd
import time
from datetime import datetime
from openpyxl import load_workbook
import pyttsx3
from winotify import Notification

engine=pyttsx3.init("sapi5")
voices=engine.getProperty('voices')
engine.setProperty('voice',voices[0].id)

toast=Notification(app_id="Python script",
                   title="Medicine Time",
                   msg='Hey Tahseen its medicine time',
                   duration='long')

def speak(audio):
    """
    This functon will be used to give output in form of voice.
    """
    engine.say(audio)
    engine.runAndWait()

# speak("Hi Tahseen how are you")

def formatTime(t,m):
    """
    This function will clean the time and convert it to 24 hours time.
    t:Time
    m:AM/PM
    """
    temp=int(t.split(":")[0])
    if m=="PM" and temp!=12:
        temp+=12
        return str(temp)+":"+str(t.split(":")[1])+":00"
    else:
        return t+":00"

def validTime(time):
    """
    This function is used to check whether the entered time is valid or not.
    """
    if int(time.split(":")[0])>=24:
        return False
    else:
        return True


if __name__=="__main__":
    diary=load_workbook("Remainder.xlsx")['Sheet1'].values
    next(diary)
    df=pd.DataFrame(data=diary,columns=('time',"am/pm","task"))
    del diary
    # print(df)
    while True:
        for index,row in df.iterrows():
            if row['time']==None:
                break
            try:
                a_time=formatTime(row['time'].strftime("%H:%M:%S"),row["am/pm"])
                if not(validTime(a_time)):
                    continue
                currentTime=datetime.now().strftime('%H:%M:%S')
                delta = datetime.strptime(a_time, '%H:%M:%S') - datetime.strptime(currentTime, '%H:%M:%S')
                del currentTime
                delta = str(delta).split(':')
                if delta[0]=='-1 day, 23':
                    delta[0]=delta[0].split(',')[1]
                # print(tdelta[0])
                secs = int(int(delta[0])*3600+int(delta[1])*60+int(delta[2]))    # Converting the remaining time into seconds.
                del delta,a_time
                time.sleep(secs)
                toast.show()
                time.sleep(1)
                speak("Hey Tahseen its medicine time")
            except Exception as e:
                print("There is some error in your sheduling")
                print(e)