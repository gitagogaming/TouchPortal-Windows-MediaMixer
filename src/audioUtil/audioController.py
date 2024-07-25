import pythoncom
from pycaw.pycaw import AudioUtilities
from logging import getLogger

g_log = getLogger(__name__)

class AudioController:
    def __init__(self, process_name):
        pythoncom.CoInitialize()
        self.process_name = process_name
        self.volume = self.get_volume()
         
    def __del__(self):
        pythoncom.CoUninitialize()
        
    def get_volume(self):
        """Getting the current focused application's volume"""
        session = self.get_session()
        if session:
            interface = session.SimpleAudioVolume
            return interface.GetMasterVolume()
        return None

    def get_session(self):
        """Helper function to get the audio session for the given process name"""
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process and session.Process.name() == self.process_name:
                return session
        return None

    def set_volume(self, volume):
        """Setting the application's volume"""
        session = self.get_session()
        if session:
            interface = session.SimpleAudioVolume
            # only set volume in the range 0.0 to 1.0
            volume = min(1.0, max(0.0, volume))
            interface.SetMasterVolume(volume, None)
            self.volume = volume

    def decrease_volume(self, decibels):
        """Decrease the application's volume"""
        session = self.get_session()
        if session:
            interface = session.SimpleAudioVolume
            # 0.0 is the min value, reduce by decibels
            new_volume = max(0.0, self.volume - decibels)
            interface.SetMasterVolume(new_volume, None)
            self.volume = new_volume

    def increase_volume(self, decibels):
        """Increase the application's volume"""
        session = self.get_session()
        if session:
            interface = session.SimpleAudioVolume
            # 1.0 is the max value, raise by decibels
            new_volume = min(1.0, self.volume + decibels)
            interface.SetMasterVolume(new_volume, None)
            self.volume = new_volume


def get_process_id(name):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        if session.Process and session.Process.name() == name:
            return session.Process.pid
    return None

def muteAndUnMute(process, value):
    sessions = AudioUtilities.GetAllSessions() ## why do we get all sessions again if done previously? why not keep track of sessions
    for session in sessions:
        volume = session.SimpleAudioVolume
        if session.Process and session.Process.name() == process:
            if value == "Toggle":
                value = 0 if volume.GetMute() == 1 else 1
            elif value == "Mute":
                value = 1
            elif value == "Unmute":
                value = 0
            volume.SetMute(value, None)


def volumeChanger(process, action, value):
    if action == "Set":
        AudioController(str(process)).set_volume((int(value)*0.01))
    elif action == "Increase":
        AudioController(str(process)).increase_volume((int(value)*0.01))

    elif action == "Decrease":
        AudioController(str(process)).decrease_volume((int(value)*0.01))


def setDeviceVolume(device, deviceid, value,  action = "Set"):
    if device is None:
        g_log.info(f"Device {deviceid} not found in audio_manager.devices.")
        return
    
    volume_scalar = value / 100.0
    if action == "Set":
        new_volume = volume_scalar
    elif action == "Increase":
        current_volume = device.GetMasterVolumeLevelScalar()
        new_volume = min(current_volume + volume_scalar, 1.0)
    elif action == "Decrease":
        current_volume = device.GetMasterVolumeLevelScalar()
        new_volume = max(current_volume - volume_scalar, 0.0)
    else:
        g_log.info(f"Unknown action {action}")
        return

    device.SetMasterVolumeLevelScalar(new_volume, None)
    g_log.info(f"Device {deviceid} volume {action.lower()}d to {new_volume * 100:.2f}%")


def setDeviceMute(device, deviceid, mute_choice)  :
            
    # Check if the volume object exists and set the master volume level
    if device:
        if mute_choice == "Toggle":
            current_mute_state = device.GetMute()
            new_mute_state = not current_mute_state
        elif mute_choice == "Mute":
            new_mute_state = 1
        elif mute_choice == "Un-Mute":
            new_mute_state = 0
            
        device.SetMute(new_mute_state, None)
    else:
        g_log.info(f"Device {device} not found in audio_manager.devices. ({deviceid})")












# def volumeChanger(process, action, value):
#     # pythoncom.CoInitialize() // this is done when creating the audicontroller object
#     if process == "Master Volume":
#         if action == "Set":
#             setMasterVolume(value)
#         else:
#             value = +value if action == "Increase" else -value
#             setMasterVolume(100 if (master_vol := getMasterVolume() + value) and master_vol > 100 else master_vol)
#     if action == "Set":
#         AudioController(str(process)).set_volume((int(value)*0.01))
#     elif action == "Increase":
#         AudioController(str(process)).increase_volume((int(value)*0.01))

#     elif action == "Decrease":
#         AudioController(str(process)).decrease_volume((int(value)*0.01))


# def setMasterVolume(Vol):
#     pythoncom.CoInitialize()
#     devices = AudioUtilities.GetSpeakers()
#     interface = devices.Activate(
#         IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
#     volume = cast(interface, POINTER(IAudioEndpointVolume))
#     scalarVolume = int(Vol) / 100
#     volume.SetMasterVolumeLevelScalar(scalarVolume, None)

# def getMasterVolume() -> int:
#     devices = AudioUtilities.GetSpeakers()
#     interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)

#     # Get the volume range and current volume level
#     volume = interface.QueryInterface(IAudioEndpointVolume)
#     volume_percent = int(round(volume.GetMasterVolumeLevelScalar() * 100))

#     devices.Release()
#     # Release the interface COM object
#     comtypes.CoUninitialize()
#     return volume_percent

# def getDeviceObject(device_id, direction="Output"):
#     deviceEnumerator = comtypes.CoCreateInstance(
#             CLSID_MMDeviceEnumerator,
#             IMMDeviceEnumerator,
#             comtypes.CLSCTX_INPROC_SERVER)
    
#     if deviceEnumerator is None: return None

#     flow = EDataFlow.eCapture.value
#     if direction.lower() == "output":
#         flow = EDataFlow.eRender.value

#     devices = deviceEnumerator.EnumAudioEndpoints(flow, 1)

#     for dev in devices:
#         if dev.GetId() == device_id:
#             return dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    
#     return None
    
# def setDeviceVolume(device_id, direction, volume_level):
#     pythoncom.CoInitialize()
#     if device_id == "default":
#         if direction.lower() == "output":
#             device = AudioUtilities.GetSpeakers()
#         else:
#             device = AudioUtilities.GetMicrophone()
#         device_object = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
#     else:
#         device_object = getDeviceObject(device_id, direction)

#     if device_object:
#         volume = cast(device_object, POINTER(IAudioEndpointVolume))
#         scalar_volume = float(volume_level) / 100
#         volume.SetMasterVolumeLevelScalar(scalar_volume, None)

#     pythoncom.CoUninitialize()

