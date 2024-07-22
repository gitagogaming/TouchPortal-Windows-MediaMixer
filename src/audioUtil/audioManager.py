import ctypes

import comtypes
from pycaw.constants import CLSID_MMDeviceEnumerator
from pycaw.pycaw import (DEVICE_STATE, AudioUtilities, EDataFlow,
                         IMMDeviceEnumerator)

import threading
import time
import pythoncom
from pycaw.constants import EDataFlow, DEVICE_STATE
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from pycaw.callbacks import MMNotificationClient
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL, COMObject
from pycaw.api.endpointvolume import IAudioEndpointVolumeCallback
from pycaw.api.endpointvolume.depend import AUDIO_VOLUME_NOTIFICATION_DATA

from TPClient import TPClient, g_log
from tppEntry import *
from tppEntry import __version__


class AudioEndpointVolumeCallback(COMObject):
    """
    A class that implements the IAudioEndpointVolumeCallback interface for handling
    audio endpoint volume notifications.

    This class is used to receive volume change notifications and mute state changes 
    for audio endpoints. It is initialized with the type, name, and ID of the device 
    it is monitoring.

    Attributes:
        device_type (str): The type of the audio device (e.g., 'input' or 'output').
        device_name (str): The friendly name of the audio device.
        device_id (str): The unique ID of the audio device.

    Methods:
        OnNotify(pNotify):
            Called when the volume or mute state of the audio device changes. 
            Prints the device type, mute status, volume level, and device name.
    """
    
    _com_interfaces_ = [IAudioEndpointVolumeCallback]

    def __init__(self, device_type, device_name, device_id):
        """
        Initializes the AudioEndpointVolumeCallback instance.

        Args:
            device_type (str): The type of the audio device (e.g., 'input' or 'output').
            device_name (str): The friendly name of the audio device.
            device_id (str): The unique ID of the audio device.
        """
        super().__init__()
        self.device_type = device_type
        self.device_name = device_name
        self.device_id = device_id

    def OnNotify(self, pNotify):
        """
        Handles volume and mute state change notifications.

        This method is called by the system when there are changes to the volume
        or mute state of the associated audio device. It prints the device type,
        current mute state, volume level, and the device name.

        Args:
            pNotify (POINTER(AUDIO_VOLUME_NOTIFICATION_DATA)): A pointer to an 
                AUDIO_VOLUME_NOTIFICATION_DATA structure that contains information 
                about the volume and mute state of the device.
        """
        notify_data = cast(pNotify, POINTER(AUDIO_VOLUME_NOTIFICATION_DATA)).contents
        # Print volume and mute status along with device name
        print(f"({self.device_type}){self.device_name}: Muted:{notify_data.bMuted} Volume: {notify_data.fMasterVolume} ")
        
        # print("Default Input:", audio_manager.device_change_client.defaultInputDevice)
        if self.device_id == audio_manager.device_change_client.defaultInputDevice:
            print("adjusting input...")
            pass
        elif self.device_id == audio_manager.device_change_client.defaultOutputDevice:
            print("adjusting output...")
            ## we need to find out what device is triggering the volume changes... hmm
            master_volume = round(notify_data.fMasterVolume * 100)
            master_volume_connector_id = f"pc_{TP_PLUGIN_INFO['id']}_{TP_PLUGIN_CONNECTORS['APP control']['id']}|{TP_PLUGIN_CONNECTORS['APP control']['data']['appchoice']['id']}=Master Volume"
            if master_volume_connector_id in TPClient.shortIdTracker:
                TPClient.shortIdUpdate(
                        TPClient.shortIdTracker[master_volume_connector_id],
                        master_volume)

            TPClient.stateUpdate(TP_PLUGIN_STATES["master volume"]["id"], str(master_volume))
            TPClient.stateUpdate(TP_PLUGIN_STATES["master volume mute"]["id"], "Muted" if notify_data.bMuted == 1 else "Un-muted")
            #         # print(f"Event Context: {notify_data.guidEventContext}")
            #         # print(f"Muted: {notify_data.bMuted}")
            #         # print(f"Master Volume: {notify_data.fMasterVolume}")
            #         # print(f"Channel Count: {notify_data.nChannels}")
            #         # print(f"Channel Volumes: {list(notify_data.afChannelVolumes[:notify_data.nChannels])}")
        else:
            print("its other devices...")

class Client(MMNotificationClient):
    """
    Handles audio device notifications such as default device changes,
    device state changes, and property value changes.
    """
    def __init__(self, setup_default_device_callback):
        """
        Initializes the Client with a callback for default device changes.
        """
        super().__init__()
        self.setup_default_device = setup_default_device_callback
        self.defaultInputDevice = None
        self.defaultOutputDevice = None

    def OnDefaultDeviceChanged(self, flow, role, pwstrDeviceId):
        print(f"Default device changed: flow={flow}, role={role}, device_id={pwstrDeviceId}")    
            
        if flow == EDataFlow.eRender.value:
            self.defaultOutputDevice = pwstrDeviceId
        elif flow == EDataFlow.eCapture.value:
            self.defaultInputDevice = pwstrDeviceId
            

        if flow == EDataFlow.eRender.value or flow == EDataFlow.eCapture.value:
            threading.Thread(target=self.setup_default_device).start()

    def OnDeviceStateChanged(self, pwstrDeviceId, dwNewState):
        print(f"Device state changed: {pwstrDeviceId} {dwNewState}")

    def OnPropertyValueChanged(self, pwstrDeviceId, key):
        print(f"Property value changed: device={pwstrDeviceId} key={key}")

class AudioManager:
    """
    Manages audio devices and their volume callbacks. Handles device registration, 
    unregistration, and notifications for changes in audio devices.

    Attributes:
        devices (dict): Stores devices by ID with their volume and callback.
        device_change_client (Client): Client for handling device change notifications.
        inputDevices (dict): Stores input devices with device names and IDs.
        outputDevices (dict): Stores output devices with device names and IDs.

    Methods:
        create_callback_for_device(device, device_type, device_name, device_id):
            Creates and registers a volume callback for a specified device.

        register_single_device(device_id):
            Registers a single device for notifications based on its ID.

        setup_devices(data_flow):
            Sets up devices (input or output) and registers callbacks for them.

        register_all_devices():
            Registers all active input and output devices.

        unregister_all_devices():
            Unregisters all devices and their callbacks.

        unregister_device(device_id):
            Unregisters a specific device by its ID.

        start_listening():
            Starts listening for device changes and sets up notifications.

        stop_listening():
            Stops listening for device changes and unregisters all devices.

        getAllDevices(direction, State):
            Retrieves all devices (input or output) based on the specified direction and state.
    """
    
    def __init__(self):
        self.devices = {}  # Store devices by ID (the callbacks)
        self.device_change_client:Client = None
        
        self.inputDevices = {}
        self.outputDevices = {}

    def create_callback_for_device(self, device, device_type, device_name, device_id):
        try:
            interface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = interface.QueryInterface(IAudioEndpointVolume)
            callback = AudioEndpointVolumeCallback(device_type, device_name, device_id)
            volume.RegisterControlChangeNotify(callback)
            return volume, callback
        except Exception as e:
            print(f"Error creating callback for device {device.GetId()}: {e}")
            return None, None

    def register_single_device(self, device_id):
        pythoncom.CoInitialize()
        enumerator = AudioUtilities.GetDeviceEnumerator()
        
        for data_flow in [EDataFlow.eRender, EDataFlow.eCapture]:
            devices = enumerator.EnumAudioEndpoints(data_flow.value, DEVICE_STATE.ACTIVE.value)
            for i in range(devices.GetCount()):
                device = devices.Item(i)
                if device.GetId() == device_id:
                    volume, callback = self.create_callback_for_device(device, "output" if data_flow == EDataFlow.eRender else "input")
                    if volume and callback:
                        self.devices[device_id] = (volume, callback)
                        print(f"Callback registered for {'output' if data_flow == EDataFlow.eRender else 'input'} device: {device_id}")
                    return
        print(f"Device with ID {device_id} not found.")
        pythoncom.CoUninitialize()
        
        
    def setup_devices(self, data_flow):
        pythoncom.CoInitialize()
        enumerator = AudioUtilities.GetDeviceEnumerator()
        
        try:
            devices = enumerator.EnumAudioEndpoints(data_flow.value, DEVICE_STATE.ACTIVE.value)
        except Exception as e:
            print(f"Error enumerating audio endpoints: {e}")
            return []

        device_list = []

        for i in range(devices.GetCount()):
            try:
                device = devices.Item(i)
                device_id = device.GetId()
                state = device.GetState()
                
                print(f"Device ID: {device_id}, State: {state}")


                if state == DEVICE_STATE.ACTIVE.value:
                    if data_flow == EDataFlow.eRender:
                        device_name = self.outputDevices.get(device_id, "None")
                        volume, callback = self.create_callback_for_device(device, "output", device_name, device_id)
                        if volume and callback:
                            device_list.append((device_id, volume, callback))
                            print(f"Callback registered for 'output' device: {device_name} {device_id}")
                    elif data_flow == EDataFlow.eCapture:
                        device_name = self.inputDevices.get(device_id, "None")
                        volume, callback = self.create_callback_for_device(device, "input", device_name, device_id)
                        if volume and callback:
                            device_list.append((device_id, volume, callback))
                            print(f"Callback registered for 'input' device: {device_name} {device_id}")
                        

            except Exception as e:
                print(f"Error processing device: {e}")

        pythoncom.CoUninitialize()
        return device_list

    def register_all_devices(self):
        self.unregister_all_devices()
        self.devices = {}
        self.devices.update({device_id: (volume, callback) for device_id, volume, callback in self.setup_devices(EDataFlow.eRender)})
        self.devices.update({device_id: (volume, callback) for device_id, volume, callback in self.setup_devices(EDataFlow.eCapture)})

    def unregister_all_devices(self):
        for volume, callback in self.devices.values():
            volume.UnregisterControlChangeNotify(callback)
        self.devices = {}

    def unregister_device(self, device_id):
        if device_id in self.devices:
            volume, callback = self.devices.pop(device_id)
            volume.UnregisterControlChangeNotify(callback)
            print(f"Callback unregistered for device: {device_id}")
        else:
            print(f"Device with ID {device_id} not registered.")

    def start_listening(self):
        self.register_all_devices()

        # Set up device change notification
        self.device_change_client = Client(self.register_all_devices)
        enumerator = AudioUtilities.GetDeviceEnumerator()
        enumerator.RegisterEndpointNotificationCallback(self.device_change_client)

    def stop_listening(self):
        enumerator = AudioUtilities.GetDeviceEnumerator()
        if self.device_change_client:
            enumerator.UnregisterEndpointNotificationCallback(self.device_change_client)
        self.unregister_all_devices()
        
    @staticmethod
    def getAllDevices(direction, State = DEVICE_STATE.ACTIVE.value):
        devices = {}
        # for all use EDataFlow.eAll.value
        if direction.lower() == "input":
            Flow = EDataFlow.eCapture.value     # 1
        else:
            # Output
            Flow = EDataFlow.eRender.value      # 0
        comtypes.CoInitialize()
        deviceEnumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER)
        if deviceEnumerator is None:
            return devices
        

        collection = deviceEnumerator.EnumAudioEndpoints(Flow, State)
        if collection is None:
            return devices

        count = collection.GetCount()
        for i in range(count):
            dev = collection.Item(i)
            if dev is not None:
                createDev = AudioUtilities.CreateDevice(dev)
                if not ": None" in str(createDev):
                    devices[createDev.FriendlyName] = createDev.id
                createDev._dev.Release()
                
        comtypes.CoUninitialize()
        return devices
    
    
audio_manager = AudioManager()

# def main():
#     # global audio_manager
#     audio_manager = AudioManager()
#     # Obtain a list of all devices
#     audio_manager.outputDevices= audio_manager.getAllDevices(direction="output")
#     audio_manager.inputDevices= audio_manager.getAllDevices(direction="input")
    
#     outputDevices = {v: k for k, v in audio_manager.outputDevices.items()} 
#     inputDevices = {v: k for k, v in audio_manager.inputDevices.items()} 
    
#     audio_manager.outputDevices = outputDevices
#     audio_manager.inputDevices = inputDevices
    
#     # audio_manager.outputDevices = all_devices
    
#     # # Print the devices to choose one
#     # print("Available output devices:")
#     # for friendly_name, device_id in all_devices.items():
#     #     print(f"Device: {friendly_name}, ID: {device_id}")

#     # # Example: Register a specific device by its ID
#     # selected_device_id = input("Enter the ID of the device you want to register: ")
#     # if selected_device_id in all_devices.values():
#     #     audio_manager.register_single_device(selected_device_id)
#     # else:
#     #     print("Invalid device ID.")
        
#     # # Example: Register a specific device by its ID
#     # selected_device_id = input("Enter the ID of the device you want to register: ")
#     # if selected_device_id in all_devices.values():
#     #     audio_manager.register_single_device(selected_device_id)
#     # else:
#     #     print("Invalid device ID.")

#     try:
#         audio_manager.start_listening()
#         print("Listening for volume and device changes... Press Ctrl+C to exit.")
#         while True:
#             time.sleep(1)
#     except KeyboardInterrupt:
#         print("Exiting...")
#     finally:
#         audio_manager.stop_listening()

# if __name__ == "__main__":
#     main()

