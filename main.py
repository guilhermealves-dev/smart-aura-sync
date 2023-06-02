import sys
import time
import pyttsx3
import speech_recognition as sr
import win32com.client
from openrgb import OpenRGBClient
from openrgb.utils import RGBColor

# Inicializar o objeto do assistente de voz
engine = pyttsx3.init()

# Definir a voz a ser utilizada
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

# Inicialização da biblioteca de reconhecimento de fala
r = sr.Recognizer()

# Inicializar o objeto do ASUS Aura SDK
try:
    auraSdk = win32com.client.Dispatch("aura.sdk.1")
    auraSdk.SwitchMode()
    devices = auraSdk.Enumerate(0)  # 0 means ALL
except:
    print("ATENCAO: Instale o programa Armoury Crate.")
    sys.exit(0)

# Inicializar OpenRGB
client_openrgb = None


def try_connect_to_client_openrgb():
    global client_openrgb
    try:
        client_openrgb = OpenRGBClient()
    except:
        client_openrgb = None


# Método para ouvir o comando de voz
def listen():
    with sr.Microphone() as source:
        print("Diga 'computador' seguido do comando de cor.")
        r.adjust_for_ambient_noise(source)

        while True:
            audio = r.listen(source)

            try:
                # Reconhecimento de fala
                command = r.recognize_google(audio, language="pt-BR")
                print("Você disse:", command)

                # Verificar se a palavra "computador" está presente no comando
                if "computador" in command.lower():
                    # Processar o comando e verificar se deve encerrar o loop
                    if not process_command(command):
                        break  # Encerra o loop do listen

            except sr.UnknownValueError:
                print("Não foi possível reconhecer o comando.")

            except sr.RequestError as e:
                print("Erro ao solicitar o serviço de reconhecimento de fala; {0}".format(e))

        # Mensagem de encerramento
        engine.say("Encerrando Smart Aura Sync. Adeus!")
        engine.runAndWait()


def process_command(command):
    # Converter o comando e as opções de cores para letras minúsculas
    command = command.lower()

    # Verificar o comando e executar a lógica de alteração de cor da placa-mãe
    if "vermelho" in command:
        change_color("vermelho")
        engine.say("Alterei a cor para vermelho.")
    elif "verde" in command:
        change_color("verde")
        engine.say("Alterei a cor para verde.")
    elif "azul" in command:
        change_color("azul")
        engine.say("Alterei a cor para azul.")
    elif "amarelo" in command:
        change_color("amarelo")
        engine.say("Alterei a cor para amarelo.")
    elif "roxo" in command:
        change_color("roxo")
        engine.say("Alterei a cor para roxo.")
    elif "ciano" in command:
        change_color("ciano")
        engine.say("Alterei a cor para ciano.")
    elif "laranja" in command:
        change_color("laranja")
        engine.say("Alterei a cor para laranja.")
    elif "rosa" in command:
        change_color("rosa")
        engine.say("Alterei a cor para rosa.")
    elif "branco" in command:
        change_color("branco")
        engine.say("Alterei a cor para branco.")
    elif "encerrar" in command:
        # Encerra o loop interno do process_command
        return False
    else:
        engine.say("Comando inválido.")

    engine.runAndWait()

    # Continua o loop interno do process_command
    return True


def change_color(color_command):
    print("-- mudando para: " + color_command)
    try_connect_to_client_openrgb()
    for dev in devices:  # Use enumeration
        for i in range(dev.Lights.Count):  # Use index
            if color_command == "vermelho":
                color = RGBColor(255, 0, 0)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "verde":
                color = RGBColor(0, 255, 0)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "azul":
                color = RGBColor(0, 0, 255)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "amarelo":
                color = RGBColor(255, 255, 0)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "roxo":
                color = RGBColor(128, 0, 128)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "ciano":
                color = RGBColor(0, 255, 255)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "laranja":
                color = RGBColor(255, 165, 0)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "rosa":
                color = RGBColor(255, 192, 203)
                apply_color_aura_sync(dev, i, color)
                apply_color_open_rgb(color)
            elif color_command == "branco":
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
        # erro ao alterar a cor via cliente open rgb
        return False


# Executar o assistente de voz
listen()
