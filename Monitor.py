#!/usr/bin/env python3
import configparser
import os
from sys import argv
import re

from time import sleep, strftime, localtime
from datetime import datetime, date, timedelta
import ctypes
from collections import deque, UserString
from threading import Lock, Thread
from io import BytesIO
import win32clipboard
import win32file

from watchdog import events, observers
from watchdog.observers.api import DEFAULT_OBSERVER_TIMEOUT, BaseObserver
from colorama import Fore, Style, init as colorama_init
from pandas.io import clipboard
from PIL import ImageGrab
import PIL  # required by openpyxl to allow handling of xlsx files with images in them

import Export

__version__ = '14.08.2019'


class Counter(object):
    # A counter for when to receive notifications. There is one counter for each machine.
    def __init__(self, machine=''):
        self._count = 0
        self._show = True  # whether to show notifications or not
        self._machine = machine

    @property
    def count(self):
        return self._count

    @count.setter
    def count(self, value):
        # If we set count to 0, we toggle notifications or or off. see notify_setting()
        if value == 0:
            self._count = 0 if self._count > 0 else 1337
        else:
            self._count = value

    @property
    def show(self):
        return Message(''.ljust(25, ' ') + self._machine + (' Displaying' if self._show else ' Hiding') + ' events')

    @show.setter
    def show(self, value):
        # If we directly assign a value, use that, else toggle
        self._show = value if value is True or value is False else not self._show

    @property
    def notify_setting(self):
        """return the current Notify setting"""
        if 0 < self.count < 10:
            return ''.ljust(25, ' ') + self._machine + ' Notifying in ' + Message(str(self.count)).white() \
                   + ' runs time'
        if self.count > 9:
            return ''.ljust(25, ' ') + self._machine + ' Notifying for ' + Message('ALL').white() + ' runs'
        elif self.count < 1:
            return ''.ljust(25, ' ') + self._machine + ' Notify OFF'


class Message(UserString):

    def __init__(self, seq):
        self.data = ''
        super().__init__(seq)

    def __repr__(self):
        # for debugging
        return f'{type(self).__name__}({super().__repr__()})'

    def __radd__(self, other):
        """
        Defining a reverse add method so that "string + Message instance" returns a Message instance
        :param other: str
        :return: Message()
        """
        if isinstance(other, str):
            return self.__class__(other + self.data)
        return self.__class__(str(other) + self.data)

    def __str__(self):
        """
        String representation of Message(). Here we can add colour highlighting to specific words
        """
        new = re.sub(r'Qiaxcel', Fore.MAGENTA + 'Qiaxcel' + Style.RESET_ALL, self.data)
        new = re.sub(r'\bQ\b', Fore.MAGENTA + 'Q' + Style.RESET_ALL, new)
        new = re.sub(r'Viia7', Fore.CYAN + 'Viia7' + Style.RESET_ALL, new)
        new = re.sub(r'\bV\b', Fore.CYAN + 'V' + Style.RESET_ALL, new)
        new = re.sub(r'\(Toggle\)', Fore.LIGHTBLACK_EX + '(Toggle)' + Style.RESET_ALL, new)
        new = re.sub(r'\bON\b', Fore.GREEN + 'ON' + Style.RESET_ALL, new)
        new = re.sub(r'\bOFF\b', Fore.RED + 'OFF' + Style.RESET_ALL, new)
        return new

    def reset(self):
        return Message(self.data + Style.RESET_ALL)

    def pre_reset(self):
        return Message(Style.RESET_ALL + self.data)

    def normal(self):
        return Message(self.data).pre_reset().reset()

    def white(self):
        return Message(Style.BRIGHT + self.data + Style.RESET_ALL)

    def white2(self):
        return Message(Fore.LIGHTWHITE_EX + self.data + Style.RESET_ALL)

    def grey(self):
        return Message(Fore.LIGHTBLACK_EX + self.data + Style.RESET_ALL)

    def green(self):
        return Message(Fore.GREEN + self.data + Style.RESET_ALL)

    def cyan(self):  # Viia7 colour
        return Message(Fore.CYAN + self.data + Style.RESET_ALL)

    def magenta(self):  # Qiaxcel colour
        return Message(Fore.MAGENTA + self.data + Style.RESET_ALL)

    def red(self):
        return Message(Fore.RED + self.data + Style.RESET_ALL)

    def yellow(self):
        return Message(Fore.YELLOW + self.data + Style.RESET_ALL)

    def timestamp(self, machine=None, distinguish=False):
        """
        Adds a timestamp to messages
        :param machine: Str : None, Viia7 or Qiaxcel
        :param distinguish: bool : Green >>> if true
        :return: Message() : Highlighted message
        """
        # TODO could this make use of the __str__ method?
        pad = 12  # The .ljust pad value- because colour is added as 0-width characters, this value changes.
        if machine:  # None, Viia7, Qiaxcel or Export
            if machine == 'Viia7' or machine == 'Qiaxcel':
                pad += 9
                machine = Message(machine)

        if distinguish:
            pad += 9
            pref = Message('>>> ').green()
        else:
            pref = ' -  '
        if machine:
            pref = pref + '{}:'.format(machine)
        ret = Message(self.bright_time() + pref.ljust(pad, ' '))
        return Message(ret + self)

    @staticmethod
    def bright_time():
        """Returns the current time formatted nicely, flanked by ANSI escape codes for bright text."""
        return Message(strftime("%d.%m %H:%M ", localtime())).white()


class LabHandler(events.PatternMatchingEventHandler):  # inheriting from watchdog's PatternMatchingEventHandler
    patterns = ['*.xdrx', '*.eds', '*.txt']  # Events are only generated for these file types.

    def __init__(self):
        super(LabHandler, self).__init__()
        self.recent_events = deque('ghi', maxlen=30)  # A list of recent events to prevent duplicate messages.
        self.error_message = deque(maxlen=1)
        self.v_counter = Counter(machine='Viia7')
        self.q_counter = Counter(machine='Qiaxcel')
        self._auto_export = True
        self._user_only = False
        self.user = os.getlogin()

    """
    The Observer passes events to the handler (this class), which then calls functions based on the type of event
    Event object properties:
    event.event_type
        'modified' | 'created' | 'moved' | 'deleted'
    event.is_directory
        True | False
    event.src_path
        path/to/observed/file
    """

    def on_modified(self, event):
        """Called when a modified event is detected. aka Viia7 events."""

        sleep(1)  # wait here to allow file to be fully written - prevents some errors with os.stat
        if '.eds' in event.src_path and event.src_path not in self.recent_events:  # .eds files we haven't seen recently
            if self.is_large_enough(event.src_path):  # this is here instead of ^ to prevent double error message
                with v_lock:
                    self.v_counter.count = self.notif(event, self.v_counter.count)

    def on_created(self, event):
        """Called when a new file is created. aka Qiaxcel/ Export events."""

        if '.xdrx' in event.src_path and event.src_path not in self.recent_events:  # .xdrx file we haven't seen
            with q_lock:
                self.q_counter.count = self.notif(event, self.q_counter.count)
        if '.txt' in event.src_path and self.user in event.src_path \
                and "Export" in event.src_path and self._auto_export:
            sleep(0.3)  # wait here to allow file to be fully written # increase if timeout happens a lot
            try:
                export.new(event.src_path)
            except OSError:  # if team drive is being slow, wait longer.
                sleep(4)
                export.new(event.src_path)
            except ValueError as e:  # When the exported file is bad
                print(e)

    def notif(self, event, x_counter):
        """Prints a notification about the event to console. May be normal or distinguished.
         Distinguished notifications flash the console window.
         Notif will be distinguished if users windows login is in path or x_counter = 1 or > 9 (now/ always setting)"""

        self.recent_events.append(event.src_path)  # Add to recent events queue to prevent duplicate notifications.
        machine, file = self.get_event_info(event)

        if os.getlogin() in event.src_path.lower() or x_counter == 1 or x_counter > 9:  # distinguished notif

            file = Message(file).green()
            message = ' {} has finished!'.format(file)
            ctypes.windll.user32.FlashWindow(ctypes.windll.kernel32.GetConsoleWindow(), True)  # Flash console window
            print(Message(message).timestamp(machine, distinguish=True))

        elif not self._user_only:  # non distinguished notification

            file = Message(file).white()
            message = ' {} has finished.'.format(file)
            if (self.q_counter.show and machine == 'Qiaxcel') or \
                    (self.v_counter.show and machine == 'Viia7'):
                print(Message(message).timestamp(machine))

        if x_counter == 1:  # If this was the run to notify on, inform that notification is now off.
            print(''.ljust(25, ' ') + machine + ' No longer notifying.')

        return x_counter - 1  # Increment counter down

    def auto_export(self):
        self._auto_export = not self._auto_export
        print(''.ljust(25, ' ') + Message('Auto export processing ON')) if self._auto_export \
            else print(''.ljust(25, ' ') + Message('Auto export processing OFF'))

    def user_only(self):
        self._user_only = not self._user_only
        print(''.ljust(25, ' ') + 'Displaying ' + Message('YOUR').white() + ' events only') if self._user_only \
            else print(''.ljust(25, ' ') + 'Displaying ' + Message('ALL').white() + ' events')

    def show_all(self):
        self.v_counter.show = True
        self.q_counter.show = True
        self._user_only = False
        print(''.ljust(25, ' ') + 'Displaying ' + Message('ALL').white() + ' events')

    @staticmethod
    def get_event_info(event):
        """Returns the machine name and file path."""
        machine = 'Viia7' if event.event_type == 'modified' else 'Qiaxcel'
        file = str(os.path.splitext(event.src_path)[0].split('\\')[-1])  # Get file name
        return machine, file

    @staticmethod
    def is_large_enough(path):
        """Determines if the file given by path is above 1300000 bytes"""
        try:
            return os.stat(path).st_size > 1300000
        except (FileNotFoundError, OSError) as e:
            file = str(os.path.splitext(path)[0].split('\\')[-1])
            print(Message(e).red())  # If not, print error, assume True.
            print(Message(file + " wasn't saved properly! You'll need to analyse "
                                 "and save the run again from the machine.").timestamp())

            return True  # Better to inform than not. I think this happens when .eds isn't saved or is deleted?


class Egel(object):
    _original = ''

    def grab(self):
        # Retrieves image from clipboard and makes sure the image is the right size
        new = ImageGrab.grabclipboard()
        # Alternate values may need to be added here, for some reason Shaheen's PC produces an image 1 px
        # taller than everyone else's (505px). I haven't tested other dpis
        assert (new.size[1] in {1575, 788, 504, 505, 394})  # These are the height values for different image dpi levels
        self._original = new

    def get(self):
        try:
            self.send_to_clipboard(self.crop(crop_type='standard'))
        except AttributeError:
            print(''.ljust(25, ' ') + 'There is no image loaded')
            pass

    def get_small(self):
        # send both scale and egel to clipboard, if using 'Office Clipboard', may paste both from clipboard history.
        try:
            self.send_to_clipboard(self.crop(crop_type='scale'), self.crop(crop_type='small'))
        except AttributeError:
            print(''.ljust(25, ' ') + 'There is no image loaded')
            pass

    def get_scale(self):
        try:
            self.send_to_clipboard(self.crop(crop_type='scale'))
        except AttributeError:
            print(''.ljust(25, ' ') + 'There is no image loaded')
            pass

    def crop(self, crop_type='standard'):
        # crop is a box within an image defined in pixels (left, top, right, bottom)
        crops = {'small': (self._original.size[1] / 5.54,       # Samples without scale
                           self._original.size[1] / 71.7,
                           self._original.size[0] - self._original.size[1] / 5.325,
                           self._original.size[1]),
                 'scale': (0,                                   # Scale only
                           self._original.size[1] / 71.7,
                           self._original.size[1] / 5.54,
                           self._original.size[1]),
                 'standard': (self._original.size[1] / 5.54,    # Samples with scale attached
                              self._original.size[1] / 71.7,
                              self._original.size[0] - self._original.size[1] / 168,
                              self._original.size[1])}
        img = self._original.crop(crops[crop_type])
        img = img.rotate(270, expand=True)
        img_out = BytesIO()
        img.convert("RGB").save(img_out, "BMP")
        img_final = img_out.getvalue()[14:]  # the first 15 bytes are header.
        img_out.close()
        return img_final

    @staticmethod
    def send_to_clipboard(*args):
        for item in args:  # Can send multiple images to clipboard one after another.
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, item)
            win32clipboard.CloseClipboard()
            sleep(0.05)  # small wait for Office Clipboard.
        print('Clipboard image processed.')


class ClipboardWatcher(Thread):
    """
    Thread that watches the clipboard for Qiaxcel images and edits them for pasting into summary files.
    """
    def __init__(self):
        super(ClipboardWatcher, self).__init__()
        self._wait = 1.5  # How often to check the clipboard, can be safely reduced if needed.
        self._paused = False
        self._stopping = False
        self.image = Egel()

    def run(self):
        """
        win32clipboard.GetClipboardSequenceNumber() changes every time the clipboard is used.
        We check if it has changed and attempt to process if it has.
        """
        recent_value = win32clipboard.GetClipboardSequenceNumber()
        while not self._stopping:
            tmp_value = win32clipboard.GetClipboardSequenceNumber()
            if not tmp_value == recent_value and not self._paused:  # if the clipboard has changed
                recent_value = tmp_value
                try:
                    self.image.grab()
                    self.image.get()
                except AttributeError:
                    pass  # given when clipboard object is not an image. Ignore
                except AssertionError:
                    pass  # given when image is not the right size. Ignore
            else:
                sleep(self._wait)

    def toggle(self):
        """
        Toggles image editing and reduces the poll rate.
        :return:
        """
        self._paused = not self._paused
        if self._paused:
            self._wait = 10
            print(Message(''.ljust(25, ' ') + 'Clipboard watcher OFF'))
        else:
            self._wait = 2.
            print(Message(''.ljust(25, ' ') + 'Clipboard watcher ON'))

    def stop(self):
        self._stopping = True


class InputLoop(Thread):
    """
    The main input loop the users sees, run in its own thread. Takes commands or file paths as
    input and calls appropriate method.
    """
    def __init__(self):
        super(InputLoop, self).__init__()
        self._stopping = False  # Kill switch to stop run() loop
        self.daemon = True
        self.instructions = {   # A reverse dictionary where the keys are the task and the key values are the commands
            labhandler.user_only: ['mine'],
            labhandler.show_all: ['all'],
            labhandler.auto_export: ['auto'],
            export.multi_off: ['done', 'stop'],
            export.multi_toggle: ['multi', 'mutli'],
            export.last_file: ['last', 'prev', 'previous', 'last file', 'lastfile'],
            export.to_file: ['to file', 'tofile', 'file'],
            egel_watcher.toggle: ['egels', 'images', 'clip'],
            egel_watcher.image.get_scale: ['scale'],
            egel_watcher.image.get: ['egel'],
            egel_watcher.image.get_small: ['small'],
            self.startup: ['install', 'setup', 'startup'],
            self.startup_remove: ['delete', 'uninstall'],
            self.stop: ['quit', 'exit', 'QQ', 'quti'],
            self.print_help: ['help', 'hlep'],
            watch.restart_observers: ['restart'],
        }

    def run(self):
        print('Running...\nEnter a command or type help for options')
        while not self._stopping:
            inp = self.get_input()
            if os.path.isfile(inp):
                try:
                    export.new(inp)
                except TypeError:  # If the txt file is bad, export.py returns a NoneType, which causes this exception
                    pass
                except ValueError as e:
                    print(e)
            else:
                inp = inp.lower()
                # Lookup command in instructions and call the method
                cont = False
                for key in self.instructions:
                    if inp in self.instructions[key]:
                        key()
                        cont = True
                        break
                if cont:
                    continue

                for char in inp:
                    if char.isdigit():  # if a digit is in input, set count = that digit and break.
                        count = int(char)
                        break  # break out of loop after first digit.
                else:
                    count = 0
                if 'q' in inp:
                    if 'hide' in inp:
                        labhandler.q_counter.show = ''
                        print(labhandler.q_counter.show)
                        continue
                    with q_lock:  # lock to make referencing variable shared between threads safe.
                        labhandler.q_counter.count = count
                        print(Message(labhandler.q_counter.notify_setting))
                if 'v' in inp:
                    if 'hide' in inp:
                        labhandler.v_counter.show = ''
                        print(labhandler.v_counter.show)
                        continue
                    with v_lock:
                        labhandler.v_counter.count = count
                        print(Message(labhandler.v_counter.notify_setting))

    @staticmethod
    def get_input():
        """
        Asks for user input. If blank, reads clipboard for input.
        :return:
        """
        inp = input('')
        if not inp:  # if blank input, read clipboard.
            inp = clipboard.clipboard_get()
        inp = os.path.normpath(inp.strip('\'"'))
        inp = inp.strip()
        return inp

    @staticmethod
    def print_help():
        help_dict = {
            'GenoTools____v' + __version__ + '____jb40': {
                Message('Monitors Qiaxcel and Viia7 and notifies when runs complete.\n'
                        '  Auto-processes Qiaxcel gel images and Viia7 export files.\n\n'
                        '   Files with your username will generate a notification.\n'
                        '    Commands may be given to notify you on other events.').normal(): ''},
            'Commands': {
                Message('    Press Enter to paste file path, or enter a command').normal(): ''},
            'Notifications': {
                'Q or V         ': ': Notifications for all events '.ljust(36, ' ') + '(Toggle)',
                'Q or V + ' + Message('digit ').yellow():
                    ": Notify after " + Message('[digit]').yellow() + " events.",
                'Q or V + ' + Message('hide  ').white():
                    ': Hide Qiaxcel or Viia7 events '.ljust(36, ' ') + '(Toggle)',
                'Mine': ': Display your events only '.ljust(36, ' ') + '(Toggle)',
                'All': ': Display all events'},
            'Auto-Processing': {
                'Auto': ': Auto-process Viia7 export files '.ljust(36, ' ') + '(Toggle)',
                'Multi': ': Start Multi export mode '.ljust(36, ' ') + '(Toggle)',
                'Done': ': Stop Multi export mode',
                'ToFile': ': Export to a pre-existing file',
                'Last': ': Export to the previous file',
                'Images': ': Auto-process Qiaxcel images       ' + '(Toggle)'},
            'Other': {
                'Install':   ': Start on Windows Startup',
                'Uninstall': ': Remove from Windows Startup',
                'Quit':      ': Exit the program'}}
        for heading in help_dict:
            print('\n' + Message(heading).white().center(68, '_') + '\n')
            for command in help_dict[heading]:
                print(Message(' ' + command).white2().ljust(25, ' ') + Message(help_dict[heading][command]))

    @staticmethod
    def startup(silent=False):
        InputLoop.startup_remove(silent=True)

        global local

        name = "Lab Helper.cmd"
        if not local:
            if not silent:
                print('You should install this locally before adding to Startup')
        else:
            name = "Lab Helper Local.cmd"
        if not silent:
            print("Installing to Startup...")
        startup_path = os.path.expanduser("~\\AppData\\Roaming\\Microsoft\\Windows"
                                          "\\Start Menu\\Programs\\Startup\\")
        with open(startup_path + name, "w+") as f:
            f.write("@echo off\n")
            f.write("cls\n")
            if '.py' in argv[0]:
                if os.path.isfile(os.getcwd() + '/Monitor.py'):
                    f.write("start /MIN python " + os.getcwd() + '/Monitor.py')
                else:
                    f.write("start /MIN python" + argv[0])
            else:
                if os.path.isfile(os.getcwd() + '/Monitor.py'):
                    f.write("start /MIN " + os.getcwd() + '/Monitor.exe')
                else:
                    f.write("start /MIN " + argv[0])
        if not silent:
            print(Message("Done!").green())

    @staticmethod
    def startup_remove(silent=False):
        if not silent:
            print('Removing from Startup...')
        startup_file = os.path.expanduser("~\\AppData\\Roaming\\Microsoft\\Windows\\"
                                          "Start Menu\\Programs\\Startup\\Lab Helper.cmd")
        startup_file_local = os.path.expanduser("~\\AppData\\Roaming\\Microsoft\\Windows\\"
                                                "Start Menu\\Programs\\Startup\\Lab Helper.cmd")
        try:
            os.remove(startup_file)
            if not silent:
                print(Message("Removed.").green())
        except FileNotFoundError:
            try:
                os.remove(startup_file_local)
            except FileNotFoundError:
                if not silent:
                    print('File not found!')

    def stop(self):
        self._stopping = True
        watch.stop_observe()


class Watcher(Thread):
    """
    This class contains all observer functionality and checks and changes the state.
    """
    def __init__(self):
        super().__init__()
        self._stopping = False
        self.obs = BaseObserver(emitter_class=MyEmitter, timeout=DEFAULT_OBSERVER_TIMEOUT)
        self.date = strftime("%b %Y", localtime())  # for checking when month changes

        self.q_watch = self.experiment_curr = self.export_curr = self.experiment_last = self.export_last = None
        self.set_watch()

    def set_watch(self):
        # Schedule observer watch locations
        self.q_watch = self.obs.schedule(labhandler, path=config['File paths']['QIAxcel'])
        self.experiment_curr = self.obs.schedule(labhandler, path=self.get_path("Experiments"))
        self.export_curr = self.obs.schedule(labhandler, path=self.get_path("Export"))

        experiment_path_last = self.get_path("Experiments", last_month=True)
        export_path_last = self.get_path("Export", last_month=True)
        self.experiment_last = self.obs.schedule(labhandler, path=experiment_path_last) \
            if experiment_path_last is not None else False
        self.export_last = self.obs.schedule(labhandler, path=export_path_last) \
            if export_path_last is not None else False

    def status(self):
        print('Observer Running') if self.obs.is_alive() else print('Observer Stopped')

    def run(self):
        global local
        self.start_observe()
        while not self._stopping:  # Check if something has changed
            if local:
                self.check_update_local()
            if self.date != strftime("%b %Y", localtime()):  # If month has changed.
                self.update_month()
            for n in range(600):  # a 10 minute loop
                if n % 30 == 0:  # every 30 seconds
                    self.check_update()
                if self._stopping:
                    break
                sleep(1)

    def update_month(self):
        """
        Called when the month has changed. Updates which folders are being watched for file changes.
        """
        self.date = strftime("%b %Y", localtime())  # Update month
        print('The month has changed to ' + self.date)

        if self.experiment_last:
            self.obs.unschedule(self.experiment_last)  # Unschedule last month watch
        if self.export_last:
            self.obs.unschedule(self.export_last)
        self.experiment_last = self.experiment_curr  # Make current month watch last month
        self.export_last = self.export_curr
        self.experiment_curr = self.obs.schedule(labhandler, self.get_path("Experiments"))  # Schedule new current month
        self.export_curr = self.obs.schedule(labhandler, self.get_path("Export"))

    @staticmethod
    def check_update():
        """
        A method to make this program close on other peoples machines by setting an update time frame
        in the config.ini. This is possible and necessary because this program is normally run from
        an exe on a network share, therefore cannot update if it is in use.
        """
        if os.path.isfile(os.getcwd() + '/config.ini'):  # may have changed.
            config.read(os.getcwd() + '/config.ini')
        else:
            config.read(os.path.normpath(os.path.dirname(argv[0]) + '/config.ini'))

        start = datetime.strptime(config['Update']['Start'], '%d.%m.%Y %H:%M')
        end = datetime.strptime(config['Update']['End'], '%d.%m.%Y %H:%M')
        if start < datetime.now() < end:
            print('Update in progress until ' + datetime.strftime(end, "%d.%m.%y %H:%M "))
            print('Closing in 10 seconds...')
            sleep(10)
            raise KeyboardInterrupt

    @staticmethod
    def check_update_local():
        master = configparser.ConfigParser()
        master.read(config['File paths']['master'] + '/config.ini')
        if not master['Update']['Version'] == __version__:
            print('Please update to the latest version of the program!')

    def get_path(self, folder, last_month=False):
        """
        Gets this or last month's file path for Experiments or Results export.
        Month abbreviations are sometimes set by people and therefore don't follow a consistent scheme.
        :param folder: str 'Experiments' or 'Export'
        :param last_month: bool if true returns path for last month.
        :return: str file path see build_path()
        """
        makedirs = True
        date_t = date.today()
        if last_month:
            makedirs = False
            first = date_t.replace(day=1)  # replace the day with the first of the month.
            date_t = first - timedelta(days=1)  # subtract one day to get the last day of last month
        month, year = date_t.strftime('%b %Y').split(' ')
        path = self.build_path(month, year, folder)
        if os.path.isdir(path):
            return path
        month_full = date_t.strftime('%B')
        path2 = self.build_path(month_full, year, folder)
        if os.path.isdir(path2):
            return path2
        if month == 'Sep':
            path2 = self.build_path('Sept', year, folder)
            if os.path.isdir(path2):
                return path2
        if makedirs:
            os.makedirs(path, exist_ok=True)
            return path
        print("Cant find last month's " + folder + " path")
        return None

    @staticmethod
    def build_path(month, year, folder):
        """
        Builds the file path from base path from config.ini and parameters
        :param month:
        :param year:
        :param folder: str 'Experiments' or 'Export'
        :return: str file path '\\file01-s0\\Team121\\Genotyping\\qPCR 2019\\Experiments\\Aug 2019'
        """
        if folder == "Experiments":
            return config['File paths']['Genotyping'] + 'qPCR ' + year + '\\Experiments\\' + month + ' ' + year
        if folder == "Export":
            return config['File paths']['Genotyping'] + 'qPCR ' + year + '\\Results Export\\' + month + ' ' + year

    def start_observe(self):
        self.obs.start()

    def stop_observe(self):
        self._stopping = True
        self.obs.unschedule_all()
        self.obs.stop()
        sleep(1)
        self.status()

    def restart_observers(self):
        self.obs.unschedule_all()
        self.obs.stop()
        sleep(2)
        self.status()

        self.obs = BaseObserver(emitter_class=MyEmitter, timeout=DEFAULT_OBSERVER_TIMEOUT)
        self.set_watch()
        self.obs.start()
        self.status()


class MyEmitter(observers.read_directory_changes.WindowsApiEmitter):
    """
    This class is used to catch an un-catchable exception in watchdog
    that would occur when the connection to the network drive was
    temporarily lost
    This can still sometimes cause a crash if the program is run from
    the network drive, but this is to do with the build method of the
    exe. A single exe build may fix but antivirus prevents it.
    Currently recommend installing to a local drive C:// - not U://
    """
    message = deque(maxlen=2)  # This is used by thread_print to prevent duplicate messages from threads.

    def queue_events(self, timeout):  # Subclass queue events - this is where the exception occurs
        try:
            super().queue_events(timeout)
        except OSError as e:    # Catch the exception and print error
            self.thread_print(str(e))
            self.thread_print('Lost connection to team drive!')
            connected = False
            while not connected:  # resume when connection to network drive is restored
                try:
                    self.on_thread_start()  # need to re-set the directory handle.
                    connected = True
                    self.thread_print('Reconnected!')
                except OSError:
                    sleep(10)
                    self.thread_print('Reconnecting...')

    def thread_print(self, msg):
        """
        Prevents duplicate print statements from threads. If it has been sent recently, it is not re sent.
        message is a deque with length 2.
        """
        if msg not in self.message:
            print(msg)
            self.message.append(msg)


if __name__ == '__main__':
    colorama_init()  # Init colorama to enable coloured text output via ANSI escape codes on windows console.
    q_lock = Lock()  # Locks used when reading or writing q_cnt or v_cnt since they are in multiple threads.
    v_lock = Lock()
    config = configparser.ConfigParser()
    if os.path.isfile(os.getcwd() + '/config.ini'):
        config.read(os.getcwd() + '/config.ini')  # config.ini = ANSI
    else:
        config.read(os.path.normpath(os.path.dirname(argv[0]) + '/config.ini'))

    local = True if win32file.GetDriveType(os.getcwd().split(':')[0] + ':') == 3 else False
    if not local:
        print('You may wish to install this program to your computer to prevent possible crashes')
    else:  # if we are running locally and a startup entry exists for the remote version, we should
        # replace it with the local version.
        if os.path.isfile(os.path.expanduser("~\\AppData\\Roaming\\Microsoft\\Windows\\"
                                             "Start Menu\\Programs\\Startup\\Lab Helper.cmd")):
            InputLoop.startup(silent=True)

    egel_watcher = ClipboardWatcher()  # Instantiate classes
    labhandler = LabHandler()
    export = Export.Export()

    watch = Watcher()
    in_loop = InputLoop()
    watch.start()  # Start threads
    egel_watcher.start()
    in_loop.start()

    try:
        while True:  # Check if something has changed
            for i in range(300):
                if not watch.is_alive():
                    raise KeyboardInterrupt
                sleep(4)
    except KeyboardInterrupt:  # on keyboard interrupt (Ctrl + C)
        watch.obs.stop()  # Stop observer + Threads (if alive)
        egel_watcher.stop()
        print('\nbye!')
