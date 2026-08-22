import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")
while True:
    text=input("you :       ")
    if text.lower()=='exit':
        break;
    speaker.Speak(text)