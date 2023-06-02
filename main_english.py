import sys
import time
import pyttsx3
import speech_recognition as sr
import win32com.client
from openrgb import OpenRGBClient
from openrgb.utils import RGBColor

# Initialize the voice assistant object
engine = pyttsx3.init()

# Set the voice to be used
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

# Initialize the speech recognition library
r = sr.Recognizer()

# Initialize the ASUS Aura SDK object
try:
    auraSdk = win32com.client.Dispatch("aura.sdk.1")
    auraSdk.SwitchMode()
    devices = auraSdk.Enumerate(0)  # 0 means ALL
except:
    print("ATTENTION: Install the Armoury Crate program.")
    sys.exit(0)

# Initialize OpenRGB
client_openrgb = None


def try_connect_to_client_openrgb():
    global client_openrgb
    try:
        client_openrgb = OpenRGBClient()
    except:
        client_openrgb = None


# Method to listen for voice commands
def listen():
    with sr.Microphone() as source:
        print("Say 'computer' followed by the color command.")
        r.adjust_for_ambient_noise(source)

        while True:
            audio = r.listen(source)

            try:
                # Speech recognition
                command = r.recognize_google(audio, language="en-US")
                print("You said:", command)

                # Check if the word "computer" is present in the command
                if "computer" in command.lower():
                    # Process the command and check if the loop should be terminated
                    if not process_command(command):
                        break  # Terminate the listen loop

            except sr.UnknownValueError:
                print("Unable to recognize the command.")

            except sr.RequestError as e:
                print("Error requesting speech recognition service; {0}".format(e))

        # Farewell message
        engine.say("Closing Smart Aura Sync. Goodbye!")
        engine.runAndWait()


def process_command(command):
    # Convert the command and color options to lowercase
    command = command.lower()

    # Check the command and execute the motherboard color change logic
    if "red" in command:
        change_color("red")
        engine.say("Changed the color to red.")
    elif "green" in command:
        change_color("green")
        engine.say("Changed the color to green.")
    elif "blue" in command:
        change_color("blue")
        engine.say("Changed the color to blue.")
    elif "yellow" in command:
        change_color("yellow")
        engine.say("Changed the color to yellow.")
    elif "purple" in command:
        change_color("purple")
        engine.say("Changed the color to purple.")
    elif "cyan" in command:
        change_color("cyan")
        engine.say("Changed the color to cyan.")
    elif "orange" in command:
        change_color("orange")
        engine.say("Changed the color to orange.")
    elif "pink" in command:
        change_color("pink")
        engine.say("Changed the color to pink.")
    elif "white" in command:
        change_color("white")
        engine.say("Changed the color to white.")
    elif "terminate" in command:
        # Terminate the inner loop of process_command
        return False
    else:
        engine.say("Invalid command.")

    engine.runAndWait()

    # Continue the inner loop of process_command
    return True


def change_color(color_command):
    print("-- changing to: " + color_command)
    try_connect_to_client_openrgb()
    for dev in devices:  # Use enumeration
        for i in range(dev.Lights.Count):  # Use index
            if color_command == "red":
                color = RGBColor(255, 0, 0)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "green":
                color = RGBColor(0, 255, 0)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "blue":
                color = RGBColor(0, 0, 255)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "yellow":
                color = RGBColor(255, 255, 0)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "purple":
                color = RGBColor(128, 0, 128)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "cyan":
                color = RGBColor(0, 255, 255)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "orange":
                color = RGBColor(255, 165, 0)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "pink":
                color = RGBColor(255, 192, 203)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "white":
                color = RGBColor(255, 255, 255)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)

        dev.Apply()

    if client_openrgb is not None:
        client_openrgb.disconnect()


def apply_color_aura_sync(device, i, color):
    device.Lights(i).red = color.red
    device.Lights(i).green = color.green
    device.Lights(i).blue = color.blue


def apply_color_open_rgb(color):
    try:
        if client_openrgb is not None:
            for device in client_openrgb.devices:
                # print(device.name)
                device.set_mode(0)
                device.set_color(color)
    except:
        # Error changing the color via OpenRGB client
        return False


# Execute the voice assistant
listen()
