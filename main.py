import os
import webbrowser

import speech_recognition as sr

import win32com.client
import openai
import datetime

speaker = win32com.client.Dispatch("SAPI.SpVoice")

chatstr = ""
def chat(query):
    global chatstr
    openai.api_key = ("your  open ai api key")
    chatstr += f"deep:{query}\njarvis:"
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=chatstr,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    chatstr += f"{(response['choices'][0]['text'])}\n"
    print(chatstr)
    speaker.speak(chatstr)
    response (response["choices"][0]["text"])


def ai(prompt):
    openai.api_key = ("your  open ai api key")
    text = f"Openai response for prompt:{prompt} \n ******************************* \n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalzty=0,
        presence_penalty=0
    )
    try:
        text += (response["choices"][0]["text"])
    except Exception as e:
        text += "sorry sir i am not able to understand"

    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    with open(f"Openai/{''.join(prompt.split('openai')[1:])}.txt","w") as f:
        f.write(text)
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"user said: {query}")
            return query
        except Exception as e:
            return "some error occurred.sorry from jarvis"


while 1:
    speaker.speak("hello deep")
    while True:
        print("listening...")
        query = takecommand()
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"],
                 ["google", "https://www.google.com"], ["facebook", "https://www.facebook.com"],
                 ["whatsapp", "https://www.whatsapp.com"], ["instagram", "https://www.instagram.com"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():

                speaker.speak(f"opening {site[0]} sir...")
                webbrowser.open(site[1])
        if "open music" in query:
             musicpath = "F:\\desktop\\stvan\\02 Ratnakar Pachchisi\\Ratnakar Pachchisi - Apoorv Avsar.mp3"
             os.startfile(musicpath)

        if "the time" in query:
             strfTime = datetime.datetime.now().strftime("%H:%M:%S")
             speaker.speak(f"sir the time is {strfTime}")

        if "open blender" .lower() in query.lower():
            speaker.speak("opening blender sir...")
            os.startfile("C:\\Program Files\\Blender Foundation\\Blender 3.0\\blender-launcher.exe")

        if "open xampp" .lower() in query.lower():
            speaker.speak("opening xampp sir...")
            os.startfile("C:\\xampp\\xampp-control.exe")

        if "open android studio" .lower() in query.lower():
            speaker.speak("opening android studio sir...")
            os.startfile("E:\\android studio\\bin\\studio64.exe")

        if "open microsoft edge" .lower() in query.lower():
            speaker.speak("opening microsoft edge sir...")
            os.startfile("C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe")

        if "jarvis" .lower() in query.lower():
            ai(prompt = query)

        if "shutdown".lower() in query.lower():
            speaker.speak("shutting down sir...")
            os.system("shutdown.exe -s -t 00")

        if "restart".lower() in query.lower():
            speaker.speak("restarting sir...")
            os.system("shutdown.exe -r -t 00")

        else:
             chat(query)

        #speaker.speak(query)
