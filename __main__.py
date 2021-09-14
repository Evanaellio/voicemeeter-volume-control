import voicemeeter
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import time
from PIL import Image
import pystray
import threading
import pkgutil

kind = 'banana'
voicemeeter.launch(kind)

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

VOLUME_DB_SHIFT = -15


def control_voicemeeter_volume():
    with voicemeeter.remote(kind) as vmr:
        while True:
            time.sleep(0.1)  # in seconds
            new_volume = volume.GetMasterVolumeLevel() + VOLUME_DB_SHIFT
            vmr.outputs[0].gain = new_volume  # Output A1
            vmr.outputs[2].gain = new_volume  # Output A3


def exit_app():
    icon.stop()


TRAY_TOOLTIP = 'Voicemeeter Volume Control'
TRAY_ICON = 'tray_icon.png'
icon = pystray.Icon(TRAY_TOOLTIP, Image.open(TRAY_ICON), menu=pystray.Menu(
    pystray.MenuItem('Exit ' + TRAY_TOOLTIP, exit_app)
))

control_thread = threading.Thread(target=control_voicemeeter_volume, daemon=True)

if __name__ == '__main__':
    control_thread.start()
    icon.run()
