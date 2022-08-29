import requests
import pprint
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import pandas as pd
import datetime
import smtplib
import ssl
import os
import yaml
from email.message import EmailMessage


def read_settings_file():
    with open('settings_yaml.yaml', 'r') as f:
        settings = yaml.load(f, Loader=yaml.FullLoader)
    return settings


settings = read_settings_file()

API_KEY = settings['api_key']
BASE_URL = "http://api.openweathermap.org/data/2.5/weather"

file_name = f'weather_report_{datetime.datetime.now().strftime("%d_%m_%Y")}.xlsx'

country_cities = {
    'India': ['Bangalore', 'Mysore', 'Delhi', 'Mumbai', 'Chennai', 'Kolkata', 'Hyderabad', 'Pune', 'Surat', 'Kanpur', 'Jaipur', 'Lucknow'],
    'Netherlands': ['Rotterdam', 'Amsterdam', 'Utrecht', 'Gouda', 'Maastricht', 'Delft', ' Den Haag', 'Haarlem', 'Lisse', 'Arnhem', 'Breda', 'Eindhoven', 'Groningen', 'Leiden', 'Nijmegen']
}

dfs = {}
for country in country_cities.keys():
    weather_data = {}
    for city in country_cities[country]:
        # city = input("Enter the city name: \n")
        weather_data[city] = {}
        request_url = f"{BASE_URL}?appid={API_KEY}&q={city}"
        response = requests.get(request_url)
        if response.status_code == 200:
            data = response.json()
            print(data)

            date_time = datetime.datetime.fromtimestamp(data["dt"])
            weather_data[city]['datetime'] = date_time
            print("datetime:", date_time)

            weather = data['weather'][0]['description']
            weather_data[city]['weather'] = weather

            temperature = round(data["main"]["temp"] - 273.15, 2)
            weather_data[city]['temperature'] = temperature
            print("weather:", weather)
            print("temperature:", temperature, "celsius")

            temp_min = round(data["main"]["temp_min"] - 273.15, 2)
            weather_data[city]['temp_min'] = temp_min
            print("temp_min:", temp_min, "celsius")

            temp_max = round(data["main"]["temp_max"] - 273.15, 2)
            weather_data[city]['temp_max'] = temp_max
            print("temp_max:", temp_max, "celsius")

            feels_like = round(data["main"]["feels_like"] - 273.15, 2)
            weather_data[city]['feels_like'] = feels_like
            print("feels_like:", feels_like, "celsius")

            humidity = data["main"]["humidity"]
            weather_data[city]['humidity'] = humidity
            print("humidity:", humidity)

            wind_speed = data["wind"]["speed"]
            weather_data[city]['wind_speed'] = wind_speed
            print("wind_speed:", wind_speed)

            wind_degree = data["wind"]["deg"]
            weather_data[city]['wind_degree'] = wind_degree
            print("wind_degree:", wind_degree)

        else:
            print("An error occurred")

    pp = pprint.PrettyPrinter(depth=4)
    pp.pprint(weather_data)

    df = pd.DataFrame(weather_data).T
    df.index.name = 'cities'
    dfs[country] = df

# Write report to Excel
with pd.ExcelWriter(file_name) as writer:
    workbook = writer.book
    for country, df in dfs.items():
        df.to_excel(writer, sheet_name=country)

# Email of weather report
for subscriber in settings['email_subscribers']:
    subject = "Daily Weather Report from Pooja"
    body = "Weather report of the day! Please check attachment."
    sender_email = settings['email_address']
    password = settings['email_password']
    message = EmailMessage()
    message["From"] = sender_email
    message["To"] = subscriber
    message["Subject"] = subject
    message.set_content(body)

    with open(file_name, 'rb') as f:
        file_data = f.read()
    message.add_attachment(file_data, maintype="application", subtype="xlsx", filename=file_name)
    context = ssl.create_default_context()

    print(f'Sending email to: {subscriber}')
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, subscriber, message.as_string())
    print(f'Email sent to: {subscriber}')
