import requests
import json
import win32com.client

while True:
    city = input("Enter your city to know the current temperature in celsius or type stop to exit:")

    if city == 'stop':
        break
    # get new key from WhetherAPI.com if current key not working
    url = f"http://api.weatherapi.com/v1/current.json?key=727fff96e6ed4bddbc5200704242104&q={city}"
    response = requests.get(url)

    weather_dic = json.loads(response.text)
    temp = weather_dic['current']['temp_c']

    speaker = win32com.client.Dispatch("SAPI.SpVoice")

    s = f"The Temperature of {city} is {temp}"
    speaker.Speak(s)