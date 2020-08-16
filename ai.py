import pyttsx3
import webbrowser
import smtplib
import random
import speech_recognition as sr
import wikipedia
import datetime
import wolframalpha
import gmplot
from win32com.client import Dispatch
import requests
import json
import os
import sys

engine = pyttsx3.init('sapi5')

client = wolframalpha.Client('Your_App_ID')

voices = engine.getProperty('voices')
engine.setProperty('voice', voices[len(voices)-1].id)

def speak(audio):
    print('Computer: ' + audio)
    engine.say(audio)
    engine.runAndWait()

def greetMe():
    currentH = int(datetime.datetime.now().hour)
    if currentH >= 0 and currentH < 12:
        speak('Good Morning!')

    if currentH >= 12 and currentH < 18:
        speak('Good Afternoon!')

    if currentH >= 18 and currentH !=0:
        speak('Good Evening!')

greetMe()

speak('Hello Lance, what\'s up?!')
speak('How may I help you?')


def myCommand():
   
    r = sr.Recognizer()                                                                                   
    with sr.Microphone() as source:                                                                       
        print("Listening...")
        r.pause_threshold =  1
        audio = r.listen(source)
    try:
        query = r.recognize_google(audio, language='en-in')
        print('User: ' + query + '\n')
        
    except sr.UnknownValueError:
        speak('Sorry! I didn\'t get that! Try typing the command!')
        query = str(input('Command: '))

    return query
        

if __name__ == '__main__':

    while True:
    
        query = myCommand();
        query = query.lower()
        
        if 'open youtube' in query:
            speak('okay')
            webbrowser.open('www.youtube.com')

        elif 'open google' in query:
            speak('okay')
            webbrowser.open('www.google.co.in')

        elif 'open gmail' in query:
            speak('okay')
            webbrowser.open('www.gmail.com')

        elif 'open netflix' in query:
            speak('okay')
            webbrowser.open('www.netflix.com')
        
        elif 'open messenger' in query:
            speak('okay')
            webbrowser.open('www.messenger.com')

        elif 'open tumblr' in query:
            speak('okay')
            webbrowser.open('www.nonstopjaguar.tumblr.com')

        elif 'open twitter' in query:
            speak('okay sure')
            webbrowser.open('www.twitter.com')

        elif 'open amazon' in query:
            speak('okay sure')
            webbrowser.open('www.primevideo.com')

        elif 'open facebook' in query:
            speak('okay sure')
            webbrowser.open('www.facebook.com')

        elif 'open instagram' in query:
            speak('okay sure')
            webbrowser.open('www.instagram.com')

        elif 'open my location' in query:
            speak('sure, noted')
            webbrowser.open('https://ipinfo.io/')

        elif 'check internet speed' in query:
            speak('wait up')
            webbrowser.open('https://fast.com/')    

        elif "what\'s up" in query or 'how are you' in query:
            stMsgs = ['Just doing my thing!', 'I am fine!', 'Nice!', 'I am nice and full of energy']
            speak(random.choice(stMsgs))

        elif "i missed you" in query: 
            stMsgs = ['i missed you too!',]
            speak(random.choice(stMsgs))

        elif "i love you" in query: 
            stMsgs = ['i love you too!',]
            speak(random.choice(stMsgs))

        elif "do you know myrtle" in query: 
            stMsgs = ['yes ofcourse, she\'s the only girl that you need, the smartest, prettiest, and the most perfect girl you love the most',]
            speak(random.choice(stMsgs))

        elif "i love her" in query: 
            stMsgs = ['you should! because she loves you too!',]
            speak(random.choice(stMsgs))

        elif "is she pretty" in query: 
            stMsgs = ['yes she is, in fact she\'s stunning that\'s why you love her']
            speak(random.choice(stMsgs))

        elif "what do you think of her" in query: 
            stMsgs = ['she\'s perfect in all aspects, especially for you']
            speak(random.choice(stMsgs))


        elif 'email' in query:
            speak('Who is the recipient? ')
            recipient = myCommand()

            if 'me' in recipient:
                try:
                    speak('What should I say? ')
                    content = myCommand()
        
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.ehlo()
                    server.starttls()
                    server.login("mrlancesalen@gmail.com", 'april102001')
                    server.sendmail('mrlancesalen@gmail.com', "lance.salen@icloud.com", content)
                    server.close()
                    speak('Email sent!')

                except:
                    speak('Sorry! I am unable to send your message at this moment!')


        elif 'nothing' in query or 'abort' in query or 'stop' in query:
            speak('okay')
            speak('Bye Lance, have a good day.')
            sys.exit()
           
        elif 'hello' in query:
            speak('Hello Lance')

        elif 'bye' in query:
            speak('Bye Lance, have a good day.')
            sys.exit()
            

        else:
            query = query
            speak('Searching...')
            try:
                try:
                    res = client.query(query)
                    results = next(res.results).text
                    speak('WOLFRAM-ALPHA says - ')
                    speak('Got it.')
                    speak(results)
                    
                except:
                    results = wikipedia.summary(query, sentences=2)
                    speak('Got it.')
                    speak('WIKIPEDIA says - ')
                    speak(results)
        
            except:
                webbrowser.open('www.google.com')
        
        speak('Anything else Lance?')
