# -*- coding: utf-8 -*-
import pythoncom
from pyWinCoreAudio import AudioDevices
from pynput import keyboard
from win32api import keybd_event, GetKeyState


class Kc:
    def __init__(self):
        self.kpress = {keyboard.Key.space: 0xb3, keyboard.Key.left: 0xb1, keyboard.Key.up: 0xaf,
                       keyboard.Key.right: 0xb0, keyboard.Key.down: 0xae}  # keys to media keys
        self.kpress = {x.value.vk: self.kpress[x] for x in self.kpress.keys()}  # extract the values of the special keys
        self.kpress[ord('M')] = 0xAD

        self.audio_devices = {}  # setup a dict of all available audio devices
        for dev in AudioDevices:
            ends = dev.render_endpoints
            if ends:
                try:
                    self.audio_devices[ends[0].name] = ends[0]
                except Exception as e:
                    print(f'An exception has occurred: {e}({type(e)})')

        self.Primary = None
        self.Secondary = None

        self.toggler = {keyboard.Key.up.value.vk: None,
                        keyboard.Key.down.value.vk: None}

        self.ctrl = False
        self.shift = False
        self.listener = keyboard.Listener(on_press=self.on_press, on_release=self.on_release,
                                          win32_event_filter=self.fil)
        self.listener.setDaemon(True)  # allows for smooth termination process

    def set_primary(self, new_primary):
        """
        sets a new primary device
        :param new_primary: audio device name
        :return:
        """
        self.toggler[keyboard.Key.up.value.vk] = self.audio_devices[new_primary]

    def set_secondary(self, new_secondary):
        """
        sets a new secondary device
        :param new_secondary: audio device name
        :return:
        """
        self.toggler[keyboard.Key.down.value.vk] = self.audio_devices[new_secondary]

    def run(self):
        """
        starts the listener thread
        :return:
        """
        self.listener.start()

    def on_press(self, key):
        """
        checks whether control and alt are pressed
        :param key: key being pressed
        :return:
        """
        if self.is_control(key):
            self.ctrl = True

        elif self.is_shift(key):
            self.shift = True

        elif 'value' in dir(key):
            if key.value.vk not in list(self.kpress.values()):  # disable override if hotkey isn't pressed
                self.ctrl = False

    def on_release(self, key):
        """
        checks if the control key was released
        :param key: key being released
        :return:
        """
        if self.is_control(key):
            self.ctrl = False

        elif self.is_shift(key):
            self.shift = False

    def fil(self, msg, data):
        """
        activate media keys based on received keys
        :param msg: msg received
        :param data: data received
        :return:
        """
        if not self.is_caps_on() and self.ctrl and data.vkCode in self.kpress.keys() and data.flags < 2:
            if self.shift:
                if data.vkCode in list(self.toggler):
                    pythoncom.CoInitialize()
                    self.toggler[data.vkCode].set_default()

            else:
                self.key_event(self.kpress[data.vkCode])  # simulate media key press

            self.listener.suppress_event()  # stop the keys from getting sent to the rest of the system

    @staticmethod
    def is_caps_on():
        """
        checks if caps lock is enabled
        :return: bool
        """
        return GetKeyState(keyboard.Key.caps_lock.value.vk)

    @staticmethod
    def key_event(e):
        """
        press key
        :param e: key to press
        :return:
        """
        keybd_event(e, 0)

    @staticmethod
    def is_control(k):
        """
        checks if control was pressed
        :param k: key pressed
        :return: bool
        """
        return k == keyboard.Key.ctrl_l or k == keyboard.Key.ctrl_r

    @staticmethod
    def is_shift(k):
        """
        checks if shift was pressed
        :param k: key pressed
        :return: bool
        """
        return k == keyboard.Key.shift_l or k == keyboard.Key.shift_r
