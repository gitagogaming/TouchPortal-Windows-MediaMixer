import os
import sys
from argparse import ArgumentParser
from ctypes import windll
from logging import (DEBUG, INFO, WARNING, FileHandler, Formatter, NullHandler,
                     StreamHandler, getLogger)
from time import sleep

import psutil
import pythoncom
import win32process
from comtypes import COMError
from pycaw.constants import AudioSessionState
from pycaw.magic import MagicManager, MagicSession
from pycaw.pycaw import EDataFlow, ERole
import comtypes


from audioUtil import audioSwitch
from audioUtil.audioController import (get_process_id, muteAndUnMute, volumeChanger)
                                       #setDeviceVolume, setMasterVolume,)
from audioUtil.audioManager import AudioManager
from windowFocusListener import WindowFocusListener

import TouchPortalAPI as TP
from tppEntry import PLUGIN_ID, TP_PLUGIN_ACTIONS, TP_PLUGIN_CONNECTORS, TP_PLUGIN_INFO, TP_PLUGIN_SETTINGS, __version__


sys.coinit_flags = 0
g_log = getLogger(__name__)

try:
    TPClient = TP.Client(
        pluginId = PLUGIN_ID,  # required ID of this plugin
        sleepPeriod = 0.05,    # allow more time than default for other processes
        autoClose = True,      # automatically disconnect when TP sends "closePlugin" message
        checkPluginId = True,  # validate destination of messages sent to this plugin
        maxWorkers = 4,        # run up to 4 event handler threads
        updateStatesOnBroadcast = False,  # do not spam TP with state updates on every page change
    )
except Exception as e:
    sys.exit(f"Could not create TP Client, exiting. Error was:\n{repr(e)}")
    
### Any action which interacts with anything NOT app related can now be updated using the audio_controller.inputDevicesReversed etc..
### See slider example...

dataMapper = {
            "Output": EDataFlow.eRender.value,
            "Input": EDataFlow.eCapture.value,
            "Default": ERole.eMultimedia.value,
            "Communications": ERole.eCommunications.value
        }


class volumeProcess:
    def __init__(self):
        self.volumeprocess = ["Master Volume", "Current app"]
        self.audio_ignore_list = []
        
    def updateVolumeMixerChoicelist(self):
        TPClient.choiceUpdate(TP_PLUGIN_ACTIONS["Inc/DecrVol"]['data']['AppChoice']['id'], self.volumeprocess[1:])
        TPClient.choiceUpdate(TP_PLUGIN_ACTIONS["AppMute"]['data']['appChoice']['id'], self.volumeprocess[1:])
        TPClient.choiceUpdate(TP_PLUGIN_CONNECTORS["APP control"]["data"]["appchoice"]['id'], self.volumeprocess)
        TPClient.choiceUpdate(TP_PLUGIN_ACTIONS["AppAudioSwitch"]["data"]["AppChoice"]["id"], self.volumeprocess[1:])
    
    def removeAudioState(self, app_name):
        if TPClient.currentStates.get(PLUGIN_ID + f".createState.{app_name}.muteState"):
            TPClient.removeStateMany([
                        PLUGIN_ID + f".createState.{app_name}.muteState",
                        PLUGIN_ID + f".createState.{app_name}.volume",
                        PLUGIN_ID + f".createState.{app_name}.active"
                        ])
            self.volumeprocess.remove(app_name)
            self.updateVolumeMixerChoicelist() # Update with new changes
        else:
            g_log.info("AudioState not present")
    
    def audioStateManager(self, app_name):
        g_log.debug(f"AUDIO EXEMPT LIST {self.audio_ignore_list}")
    
        if app_name not in self.volumeprocess:
            g_log.info("Creating states")
            TPClient.createStateMany([
                    {   
                        "id": PLUGIN_ID + f".createState.{app_name}.muteState",
                        "desc": f"{app_name} Mute State",
                        "parentGroup": "Audio process state",
                        "value": ""
                    },
                    {
                        "id": PLUGIN_ID + f".createState.{app_name}.volume",
                        "desc": f"{app_name} Volume",
                        "parentGroup": "Audio process state",
                        "value": ""
                    },
                    {
                        "id": PLUGIN_ID + f".createState.{app_name}.active",
                        "desc": f"is {app_name} Active",
                        "parentGroup": "Audio process state",
                        "value": ""
                    },
                    ])
            self.volumeprocess.append(app_name)
    
            """UPDATING CHOICES"""
            self.updateVolumeMixerChoicelist()
            g_log.debug(f"{app_name} state added")
    
        """ Checking for Exempt Audio"""
        if app_name in self.audio_ignore_list:
            self.removeAudioState(app_name)
            return True
        return False
        

class AppAudioCallBack(MagicSession):
    def __init__(self):
        super().__init__(volume_callback=self.update_volume,
                         mute_callback=self.update_mute,
                         state_callback=self.update_state)

        # ______________ DISPLAY NAME ______________
        self.app_name = self.magic_root_session.app_exec
        #g_log.info(f":: new session: {self.app_name}")
        
        if self.app_name not in volumeprocess.audio_ignore_list:
            # set initial:
            self.update_mute(self.mute)
            self.update_state(self.state)
            self.update_volume(self.volume)
        

    def update_state(self, new_state):
        """
        when status changed
        (see callback -> AudioSessionEvents -> OnStateChanged)
        """
        
        if self.app_name not in volumeprocess.audio_ignore_list:
            if new_state == AudioSessionState.Inactive:
                # AudioSessionStateInactive
                """Sesssion is Inactive"""
                g_log.debug(f"{self.app_name} not active")
                TPClient.stateUpdate(PLUGIN_ID + f".createState.{self.app_name}.active","False")
    
            elif new_state == AudioSessionState.Active:
                """Session Active"""
                g_log.debug(f"{self.app_name} is an Active Session")
                TPClient.stateUpdate(PLUGIN_ID + f".createState.{self.app_name}.active","True")
    
        if new_state == AudioSessionState.Expired:
            """Removing Expired States"""
            volumeprocess.removeAudioState(self.app_name)

    
    def update_volume(self, new_volume):
        """
        when volume is changed externally - Updating Sliders and Volume States
        (see callback -> AudioSessionEvents -> OnSimpleVolumeChanged )
        """
        
        if self.app_name not in volumeprocess.audio_ignore_list:
            TPClient.stateUpdate(PLUGIN_ID + f".createState.{self.app_name}.volume", str(round(new_volume*100)))
            #g_log.info(f"{self.app_name} NEW VOLUME", str(round(new_volume*100)))
            app_connector_id =f"pc_{TP_PLUGIN_INFO['id']}_{TP_PLUGIN_CONNECTORS['APP control']['id']}|{TP_PLUGIN_CONNECTORS['APP control']['data']['appchoice']['id']}={self.app_name}"
            if app_connector_id in TPClient.shortIdTracker :
                TPClient.shortIdUpdate(
                    TPClient.shortIdTracker[app_connector_id],
                    round(new_volume*100))
            """Checking for Current App If Its Active, Adjust it also"""
            if (activeWindow := getActiveExecutablePath()) != "":
                current_app_connector_id = f"pc_{TP_PLUGIN_INFO['id']}_{TP_PLUGIN_CONNECTORS['APP control']['id']}|{TP_PLUGIN_CONNECTORS['APP control']['data']['appchoice']['id']}=Current app"

                if current_app_connector_id in TPClient.shortIdTracker :
                    TPClient.shortIdUpdate(
                        TPClient.shortIdTracker[current_app_connector_id],
                        int(new_volume*100) if os.path.basename(activeWindow) == self.app_name else 0)
            g_log.debug(f"Volume: {self.app_name} - {new_volume}")
            
    def update_mute(self, muted):
        """ when mute state is changed by user or through other app """
        if self.app_name not in volumeprocess.audio_ignore_list:
            isDeleted = volumeprocess.audioStateManager(self.app_name)
            if not isDeleted:
                TPClient.stateUpdate(PLUGIN_ID + f".createState.{self.app_name}.muteState", "Muted" if muted else "Un-muted")
                g_log.debug(f"Mute State: {self.app_name} - {'Muted' if muted else 'Un-muted'}")



def updateDevice(options, choiceId, instanceId=None):
    deviceList = list(audioSwitch.MyAudioUtilities.getAllDevices(options).keys())
    if (choiceId == TP_PLUGIN_ACTIONS["AppAudioSwitch"]["data"]["devicelist"]["id"] or \
            choiceId == TP_PLUGIN_ACTIONS["setDeviceVolume"]["data"]["deviceOption"]["id"] or \
                choiceId == TP_PLUGIN_CONNECTORS['Windows Audio']['data']['deviceOption']['id']):

        deviceList.insert(0, "Default")
    if instanceId:
        TPClient.choiceUpdateSpecific(choiceId, deviceList, instanceId)
    else:
        TPClient.choiceUpdate(choiceId, deviceList)
    g_log.debug(f'updating {options} {deviceList}')


def getActiveExecutablePath():
    hWnd = windll.user32.GetForegroundWindow()
    if hWnd == 0:
        return None # Note that this function doesn't use GetLastError().
    else:
        _, pid = win32process.GetWindowThreadProcessId(hWnd)
        return psutil.Process(pid).exe()



STGM_READ = 0x00000000
def getDevicebydata(edata, erole):
    DEVPKEY_Device_FriendlyName = "{a45c254e-df1c-4efd-8020-67d146a850e0} 14".upper()

    device = ""
    audioDevice = None

    comtypes.CoInitialize()
    try:
        audioDevice = audioSwitch.MyAudioUtilities.GetDeviceState(edata, erole)
        if audioDevice:
            properties = {}
            store = audioDevice.OpenPropertyStore(STGM_READ)
            try:
                propCount = store.GetCount()
                for j in range(propCount):
                    pk = store.GetAt(j)
                    value = store.GetValue(pk)
                    v = value.GetValue()
                    value.clear()

                    name = str(pk)
                    properties[name] = v
                device = properties.get(DEVPKEY_Device_FriendlyName, "")

            finally:
                value.clear()
                store.Release()
                del store
                del properties

    except COMError as exc:
        pass

    finally:
        if audioDevice:
            audioDevice.Release()
            del audioDevice
        comtypes.CoUninitialize()
    return str(device)
        



def handleSettings(settings, on_connect=False):
    settings = { list(settings[i])[0] : list(settings[i].values())[0] for i in range(len(settings)) }
    if (value := settings.get(TP_PLUGIN_SETTINGS['ignore list']['name'])) is not None:
        volumeprocess.audio_ignore_list = value if value != TP_PLUGIN_SETTINGS['ignore list']['default'] else []
        
        try:   ## removing audio states assuming the user edited settings while plugin is running.. this avoids needing for user to reboot plugin for changes to take place
            split_ignorelist = [app_name.replace(",", "").strip() + ".exe" for app_name in volumeprocess.audio_ignore_list.split(".exe") if app_name.strip()]
            for app_name in split_ignorelist:
                if app_name in volumeprocess.audio_ignore_list:
                    volumeprocess.removeAudioState(app_name)
        except:
            pass


@TPClient.on(TP.TYPES.onConnect)
def onConnect(data):
    g_log.info(f"Connected to TP v{data.get('tpVersionString', '?')}, plugin v{data.get('pluginVersion', '?')}.")
    g_log.debug(f"Connection: {data}")
    if settings := data.get('settings'):
        handleSettings(settings, True)

    run_callbacks()    
    

def run_callbacks():
    pythoncom.CoInitialize()
    try:
        MagicManager.magic_session(AppAudioCallBack)
    except Exception as e:
        g_log.info(e, exc_info=True)
    
    try:  
        listener.start()
    except Exception as e:
        g_log.info(e, exc_info=True)

    try:
        audio_manager.fetch_devices()
        audio_manager.start_listening()
    except Exception as e:
        g_log.info(e, exc_info=True)



# Settings handler
@TPClient.on(TP.TYPES.onSettingUpdate)
def onSettingUpdate(data):
    g_log.debug(f"Settings: {data}")
    if (settings := data.get('values')):
        handleSettings(settings, False)

# Action handler
@TPClient.on(TP.TYPES.onAction)
def onAction(data):
    g_log.debug(f"Action: {data}")
    # check that `data` and `actionId` members exist and save them for later use
    if not (action_data := data.get('data')) or not (actionid := data.get('actionId')):
        return
    
    ## For Apps            
    elif actionid == TP_PLUGIN_ACTIONS['AppMute']['id']:
        if action_data[0]['value'] != '':
            if action_data[0]['value'] == "Current app":
                activeWindow = getActiveExecutablePath()
                if activeWindow != "":
                    muteAndUnMute(os.path.basename(activeWindow), action_data[1]['value'])
            else:
                muteAndUnMute(action_data[0]['value'], action_data[1]['value'])
    
    ## For Apps            
    elif actionid == TP_PLUGIN_ACTIONS['Inc/DecrVol']['id']:
        volume_value = int(action_data[2]['value'])
        volume_value = max(0, min(volume_value, 100))

        if action_data[0]['value'] == "Current app":
            activeWindow = getActiveExecutablePath()
            if activeWindow != "":
                volumeChanger(os.path.basename(activeWindow), action_data[1]['value'], volume_value)
        else:
            volumeChanger(action_data[0]['value'], action_data[1]['value'], volume_value)
    
    # For Devices
    elif actionid == TP_PLUGIN_ACTIONS["ChangeOut/Input"]["id"] and action_data[0]['value'] != "Pick One": 
        name = action_data[1]['value']
        device_type = action_data[0]['value']
        device, deviceId = audio_manager.getDeviceByName(name, device_type)
        if deviceId:
            audioSwitch.switchOutput(deviceId, dataMapper[action_data[2]['value']])

    # For Devices
    elif actionid == TP_PLUGIN_ACTIONS["ToggleOut/Input"]["id"] and action_data[0]['value'] != "Pick One":
        deviceId = audioSwitch.MyAudioUtilities.getAllDevices(action_data[0]['value'])
        currentDeviceId = deviceId.get(getDevicebydata(dataMapper[action_data[0]['value']], dataMapper[action_data[3]['value']]))
        choiceDeviceId1 = deviceId.get(action_data[1]['value'])
        choiceDeviceId2 = deviceId.get(action_data[2]['value'])
        if (choiceDeviceId1 and choiceDeviceId2):
            if choiceDeviceId1 == currentDeviceId:
                audioSwitch.switchOutput(choiceDeviceId2, dataMapper[action_data[3]['value']])
            else:
                audioSwitch.switchOutput(choiceDeviceId1, dataMapper[action_data[3]['value']])

    # For Apps
    elif actionid == TP_PLUGIN_ACTIONS["AppAudioSwitch"]["id"] and action_data[2]["value"] != "Pick One":
        deviceId = ""
        if (action_data[1]["value"] != "Default"):
            deviceId = audioSwitch.MyAudioUtilities.getAllDevices(action_data[2]["value"])
            deviceId = deviceId.get(action_data[1]["value"])

        if ((processid := get_process_id(action_data[0]['value'])) != None):
            g_log.info(f"args devId: {deviceId}, processId: {processid}")
            if (deviceId == "" and action_data[1]["value"] == "Default") or deviceId:
                audioSwitch.SetApplicationEndpoint(deviceId, 1 if action_data[2]["value"] == "Input" else 0, processid)

    # For Devices
    elif actionid == TP_PLUGIN_ACTIONS["setDeviceVolume"]["id"] and action_data[0]["value"] != "Pick One":
        device_type = action_data[0]['value']
        name = action_data[1]['value']
        volume_value = float(action_data[2]['value'])
        setDeviceVolume(device_type, name, volume_value)
    
    # For Devices        
    elif actionid == TP_PLUGIN_ACTIONS["setDeviceMute"]["id"] and action_data[0]["value"] != "Pick One":
        device_type = action_data[0]['value']
        name = action_data[1]['value']
        mute_choice = action_data[2]['value']
        
        device, deviceid = audio_manager.getDeviceByName(name, device_type)
            
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

    else:
        g_log.warning("Got unknown action ID: " + actionid)

### Needs fixed with new stuff...
@TPClient.on(TP.TYPES.onHold_down)
def heldingButton(data):
    g_log.debug(f"heldingButton: {data}")
    while True:
        if TPClient.isActionBeingHeld(TP_PLUGIN_ACTIONS['Inc/DecrVol']['id']):
            volume_value = int(data['data'][2]['value'])
            volume_value = max(0, min(volume_value, 100))
            volumeChanger(data['data'][0]['value'], data['data'][1]['value'], volume_value)
            sleep(0.05)
        else:
            break
    g_log.debug(f"Not helding button {data}")
    

            
@TPClient.on(TP.TYPES.onConnectorChange)
def connectors(data):
    g_log.debug(f"connector Change: {data}")
    if data['connectorId'] == TP_PLUGIN_CONNECTORS["APP control"]['id']:
        if data['data'][0]['value'] == "Master Volume":
            # maintaining backwards compatible with old plugin actions...
            # this was originally in the 'app' connector.. ask KB why
            device_type = "Output"
            name = "Default"

            slider_value = float(data['value'])
            volume_scalar = slider_value / 100.0

            device, deviceid = audio_manager.getDeviceByName(name, device_type)

            # Check if the volume object exists and set the master volume level
            if device:
                device.SetMasterVolumeLevelScalar(volume_scalar, None)
            else:
                g_log.info(f"Device {device} not found in audio_manager.devices. ({deviceid})")
                
        elif data['data'][0]['value'] == "Current app":
            activeWindow = getActiveExecutablePath()
            
            if activeWindow != "":
                volumeChanger(os.path.basename(activeWindow), "Set", data['value'])
        else:
            try:
                volumeChanger(data['data'][0]['value'], "Set", data['value'])
            except Exception as e:
                g_log.debug(f"Exception in other app volume change Error: " + str(e))



    elif data["connectorId"] == TP_PLUGIN_CONNECTORS["Windows Audio"]["id"]:
        device_type = data['data'][0]['value']
        name = data['data'][1]['value']
        slider_value = float(data['value'])
        setDeviceVolume(device_type, name, slider_value)

            
def setDeviceVolume(device_type, name, value):
    volume_scalar = value / 100.0
    
    device, deviceid = audio_manager.getDeviceByName(name, device_type)
        
    # Check if the volume object exists and set the master volume level
    if device:
        device.SetMasterVolumeLevelScalar(volume_scalar, None)
    else:
        g_log.info(f"Device {device} not found in audio_manager.devices. ({deviceid})")

    

@TPClient.on(TP.TYPES.onListChange)
def onListChange(data):
    g_log.info(f"onlistChange: {data}")
    if data['actionId'] == TP_PLUGIN_ACTIONS["ChangeOut/Input"]['id'] and \
        data['listId'] == TP_PLUGIN_ACTIONS["ChangeOut/Input"]["data"]["optionSel"]["id"]:
        try:
            updateDevice(data['value'], TP_PLUGIN_ACTIONS["ChangeOut/Input"]['data']['deviceOption']['id'], data['instanceId'])
        except Exception as e:
            g_log.info("Update device input/output KeyError: " + str(e))
    elif data['actionId'] == TP_PLUGIN_ACTIONS["ToggleOut/Input"]['id'] and \
        data['listId'] == TP_PLUGIN_ACTIONS["ToggleOut/Input"]["data"]["optionSel"]["id"]:
        try:
            updateDevice(data['value'], TP_PLUGIN_ACTIONS["ToggleOut/Input"]['data']['deviceOption1']['id'], data['instanceId'])
            updateDevice(data['value'], TP_PLUGIN_ACTIONS["ToggleOut/Input"]['data']['deviceOption2']['id'], data['instanceId'])
        except Exception as e:
            g_log.info("Update device input/output KeyError: " + str(e))
    elif data['actionId'] == TP_PLUGIN_ACTIONS["AppAudioSwitch"]["id"] and \
        data["listId"] == TP_PLUGIN_ACTIONS["AppAudioSwitch"]["data"]["deviceType"]["id"]:
        try:
            updateDevice(data['value'], TP_PLUGIN_ACTIONS["AppAudioSwitch"]["data"]["devicelist"]["id"], data['instanceId'])
        except Exception as e:
            g_log.info("Update device input/output KeyError: " + str(e))
    
    elif data['actionId'] == TP_PLUGIN_ACTIONS["setDeviceVolume"]["id"] and \
        data["listId"] == TP_PLUGIN_ACTIONS["setDeviceVolume"]["data"]["deviceType"]["id"]:
        try:
            updateDevice(data['value'], TP_PLUGIN_ACTIONS["setDeviceVolume"]["data"]["deviceOption"]["id"], data['instanceId'])
        except Exception as e:
            g_log.info("Update device setDeviceVolume error " + str(e))
    
    elif data['actionId'] == TP_PLUGIN_ACTIONS["setDeviceMute"]["id"] and \
        data["listId"] == TP_PLUGIN_ACTIONS["setDeviceMute"]["data"]["deviceType"]["id"]:
        try:
            updateDevice(data['value'], TP_PLUGIN_ACTIONS["setDeviceMute"]["data"]["deviceOption"]["id"], data['instanceId'])
        except Exception as e:
            g_log.info("Update device setDeviceVolume error " + str(e))
    
    ## left this the same, but it could be modified a bit.. atleast the UpdateDevice func so it retrieves from audio_controller
    elif data['actionId'] == TP_PLUGIN_CONNECTORS["Windows Audio"]["id"] and \
        data["listId"] == TP_PLUGIN_CONNECTORS["Windows Audio"]["data"]["deviceType"]["id"]:
        try:
            updateDevice(data['value'], TP_PLUGIN_CONNECTORS["Windows Audio"]["data"]["deviceOption"]["id"], data['instanceId'])
        except Exception as e:
            g_log.warning("Update device Windows Audio error " + str(e))


# Shutdown handler
@TPClient.on(TP.TYPES.onShutdown)
def onShutdown(data):
    audio_manager.unregister_all_devices() # Shuts down cleanly without these
    audio_manager.stop_listening()         # but dont know the potential issues that may arise..
    listener.stop()
    g_log.info('Received shutdown event from TP Client.')

# Error handler
# @TPClient.on(TP.TYPES.onError)
# def onError(exc):
#     g_log.error(f'Error in TP Client event handler: {repr(exc)}')

def main():
    global TPClient, g_log
    if not g_log.hasHandlers():
        # Handle CLI arguments
        parser = ArgumentParser()
        parser.add_argument("-d", action='store_true',
                            help="Use debug logging.")
        parser.add_argument("-w", action='store_true',
                            help="Only log warnings and errors.")
        parser.add_argument("-q", action='store_true',
                            help="Disable all logging (quiet).")
        parser.add_argument("-l", metavar="<logfile>",
                            help="Log to this file (default is stdout).")
        parser.add_argument("-s", action='store_true',
                            help="If logging to file, also output to stdout.")

        opts = parser.parse_args()
        del parser

        # set up logging
        if opts.q:
            # no logging at all
            g_log.addHandler(NullHandler())
        else:
            # set up pretty log formatting (similar to TP format)
            fmt = Formatter(
                fmt="{asctime:s}.{msecs:03.0f} [{levelname:.1s}] [{filename:s}:{lineno:d}] {message:s}",
                datefmt="%H:%M:%S", style="{"
            )
            # set the logging level
            if   opts.d: g_log.setLevel(DEBUG)
            elif opts.w: g_log.setLevel(WARNING)
            else:        g_log.setLevel(INFO)


            # set up log destination (file/stdout)
            if opts.l:
                try:
                    # note that this will keep appending to any existing log file
                    fh = FileHandler(str("log.txt"))
                    fh.setFormatter(fmt)
                    g_log.addHandler(fh)
                except Exception as e:
                    opts.s = True
                    g_log.info(f"Error while creating file logger, falling back to stdout. {repr(e)}")
            if not opts.l or opts.s:
                sh = StreamHandler(sys.stdout)
                sh.setFormatter(fmt)
                g_log.addHandler(sh)

    g_log.info(f"Starting {TP_PLUGIN_INFO['name']} v{__version__} on {sys.platform}.")
    ret = 1
    try:
        # Connect to Touch Portal desktop application.
        # If connection succeeds, this method will not return (blocks) until the client is disconnected.
        TPClient.connect()
        g_log.info('TP Client closed.')
    except KeyboardInterrupt:
        g_log.warning("Caught keyboard interrupt, exiting.")
    except Exception:
        # This will catch and report any critical exceptions in the base TPClient code,
        # _not_ exceptions in this plugin's event handlers (use onError(), above, for that).
        from traceback import format_exc
        g_log.error(f"Exception in TP Client:\n{format_exc()}")
        ret = -1
    finally:
        # Make sure TP Client is stopped, this will do nothing if it is already disconnected.
        TPClient.disconnect()

    # TP disconnected, clean up.
    del TPClient

    g_log.info(f"{TP_PLUGIN_INFO['name']} stopped.")
    return ret

if __name__ == "__main__":
    volumeprocess = volumeProcess()
    audio_manager = AudioManager(TPClient)
    listener = WindowFocusListener(TPClient)
    main()
