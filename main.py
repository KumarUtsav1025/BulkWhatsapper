import pywhatkit as kit
import pyautogui as pg
import pandas as pd
import numpy as np
import time


def get_profiles():
    df = pd.read_excel('profileList.xlsx')
    df_array = np.array(df)
    return df_array


def message_body(name, salutation):
    body = f"Good Afternoon {name} {salutation}, This is python automated message."
    return body


def send_whatsapp_message(phone, name, salutation):
    # Specify the phone number (with country code) and the message
    phone_number = phone
    message = message_body(name, salutation)

    # Send the message instantly
    kit.sendwhatmsg_instantly(phone_number, message)
    time.sleep(0.5)
    pg.click(x=1000, y=960)  # Adjust the coordinates based on your screen
    time.sleep(1)
    pg.press("Enter")
    time.sleep(1)
    pg.hotkey('ctrl', 'w')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print("Process Initiated....")
    profiles = get_profiles()
    for profile in profiles:
        send_whatsapp_message(phone=f"+91{profile[2]}", name=f"{profile[0]}", salutation=f"{profile[1]}")
        time.sleep(1.5)
    print("Process completed....")