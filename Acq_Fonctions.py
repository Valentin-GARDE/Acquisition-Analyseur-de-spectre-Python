import socket
import pyvisa # Importation de la bibliothèque PyVISA pour la communication avec les instruments
import time
import re

# Module Acq_Fonctions.py
# Fonctions de connexion et configuration d'un analyseur de spectre

def connecter_appareil_usb():
    """Cherche et ouvre un appareil USB via VISA."""
    try:
        rm = pyvisa.ResourceManager()
        ressources = rm.list_resources()
        usb_devices = [r for r in ressources if "USB" in r]

        if not usb_devices:
            print("Aucun appareil USB compatible trouvé.")
            return None

        ressource = usb_devices[0]
        print(f"Appareil USB trouvé : {ressource}")
        instrument = rm.open_resource(ressource)
        instrument.timeout = 5000  # 5 secondes
        idn = instrument.query("*IDN?")  # Test de communication
        print(f"Réponse *IDN? : {idn.strip()}")
        return instrument

    except Exception as e:
        print(f"Erreur de connexion USB : {e}")
        return None


def send_usb_command(appareil, command):
    """Envoie une commande SCPI via USB et retourne la réponse éventuelle."""
    try:
        if '?' in command:
            return appareil.query(command).strip()
        else:
            appareil.write(command)
    except Exception as e:
        print(f"Erreur USB '{command}' : {e}")
        return None


def connect_to_device(ip, port):
    """Connecte l'appareil via LAN (socket TCP)."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.connect((ip, port))
        s.settimeout(5)
        print(f"Connecté à {ip}:{port}")
        return s
    except Exception as e:
        print(f"Erreur connexion LAN : {e}")
        return None


def send_scpi_command(sock, command):
    """Envoie une commande SCPI via socket LAN."""
    try:
        sock.sendall((command + '\n').encode())
    except Exception as e:
        print(f"Erreur SCPI '{command}' : {e}")


def receive_data(sock):
    """Reçoit la réponse d'une commande et renvoie la chaîne."""
    try:
        return sock.recv(8192).decode()
    except Exception as e:
        print(f"Erreur réception : {e}")
        return None


def save_data_to_file(data, filename):
    """Sauvegarde la chaîne data dans un fichier texte."""
    try:
        with open(filename, 'w') as f:
            f.write(data)
    except Exception as e:
        print(f"Erreur sauvegarde fichier : {e}")


def configure_spectrum(inst, *,
                       start_freq_hz=None, span_hz=None,
                       rbw_hz=None, vbw_hz=None,
                       input_atten_db=None,
                       continuous=True):
    """
    Configure l'analyseur de spectre avant sweep :
     - start_freq_hz : fréquence de début (Hz)
     - span_hz       : largeur de bande (Hz)
     - rbw_hz        : bande passante IF (Hz)
     - vbw_hz        : bande passante vidéo (Hz)
     - input_atten_db: atténuation d'entrée (dB)
     - continuous    : balayage continu si True, unique si False
    """
    # Gestion des fréquences via centre/span pour stabilité
    if start_freq_hz is not None and span_hz is not None:
        center = start_freq_hz + span_hz/2
        inst.write(f":SENSe:FREQuency:CENTer {center}\n")
        inst.write(f":SENSe:FREQuency:SPAN {span_hz}\n")
    elif start_freq_hz is not None:
        inst.write(f":SENSe:FREQuency:STARt {start_freq_hz}\n")
    elif span_hz is not None:
        inst.write(f":SENSe:FREQuency:SPAN {span_hz}\n")

    # RBW / VBW
    if rbw_hz is not None:
        inst.write(f":SENSe:BWIDth:RESolution {rbw_hz}\n")
    if vbw_hz is not None:
        inst.write(f":SENSe:BWIDth:VIDEO {vbw_hz}\n")

    # Atténuation d'entrée
    if input_atten_db is not None:
        inst.write(f":INPut:ATTenuation {input_atten_db}\n")

    # Mode de balayage
    mode = "CONTinuous" if continuous else "SINGle"
    inst.write(f":INITiate:{mode} ON\n")
