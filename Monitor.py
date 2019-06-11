import time
from datetime import datetime, date, timedelta
import ctypes
import os
from collections import deque
from threading import Lock, Thread
from sys import argv
import configparser
import PIL  # required by openpyxl to allow handling of xlsx files with images in them
try:
    import Export
    from watchdog import events, observers
    from watchdog.observers.api import DEFAULT_OBSERVER_TIMEOUT, BaseObserver
    from colorama import Fore, Style, init as colorama_init
    from pandas.io import clipboard
    from PIL import ImageGrab
    from io import BytesIO
    import win32clipboard
    import msvcrt
except ModuleNotFoundError:
    input("Please install the requirements! \nPress Enter to exit.")
    quit()


class Counter(object):
    def __init__(self, machine=''):
        self._count = 0
        self._show = True
        self._machine = machine

    @property
    def count(self):
        return self._count

    @count.setter
    def count(self, value):
        if value == 0:
            self._count = 0 if self._count > 0 else 1337
        else:
            self._count = value

    @property
    def show(self):
        return ''.ljust(19, ' ') + self._machine + (' Displaying' if self._show else ' Hiding') + ' events'

    @show.setter
    def show(self, value):
        self._show = value if value is True or value is False else not self._show

    @property
    def notify_setting(self):
        """return the current Notify setting"""
        if 0 < self.count < 10:
            return ''.ljust(19, ' ') + self._machine + ' Notifying in ' + str(self.count) + ' runs time'
        if self.count > 9:
            return ''.ljust(19, ' ') + self._machine + ' Notifying for all runs'
        elif self.count < 1:
            return ''.ljust(19, ' ') + self._machine + ' Notify OFF'


class Message(str):
    def white(self):
        return Message(Style.BRIGHT + self + Style.RESET_ALL)

    def green(self):
        return Message(Fore.GREEN + self + Style.RESET_ALL)

    def cyan(self):
        return Message(Fore.CYAN + self + Style.RESET_ALL)

    def magenta(self):
        return Message(Fore.MAGENTA + self + Style.RESET_ALL)

    def red(self):
        return Message(Fore.RED + self + Style.RESET_ALL)

    def yellow(self):
        return Message(Fore.YELLOW + self + Style.RESET_ALL)

    def grey(self):
        return Message(Fore.LIGHTBLACK_EX + self + Style.RESET_ALL)

    def blue(self):
        return Message(Fore.BLUE + self + Style.RESET_ALL)

    def light_blue(self):
        return Message(Fore.LIGHTBLUE_EX + self + Style.RESET_ALL)

    def light_green(self):
        return Message(Fore.LIGHTGREEN_EX + self + Style.RESET_ALL)

    def light_red(self):
        return Message(Fore.LIGHTRED_EX + self + Style.RESET_ALL)

    def light_cyan(self):
        return Message(Fore.LIGHTCYAN_EX + self + Style.RESET_ALL)

    def light_magenta(self):
        return Message(Fore.LIGHTMAGENTA_EX + self + Style.RESET_ALL)

    def timestamp(self, machine=None, distinguish=False):
        pad = 12  # The .ljust pad value- because colour is added as 0-width characters, this value changes.
        if machine:
            pad += 9
            if machine == 'Viia7':
                machine = Message(machine).cyan()
            else:
                machine = Message(machine).magenta()

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
        return Message(time.strftime("%d.%m %H:%M ", time.localtime())).white()


class LabHandler(events.PatternMatchingEventHandler):  # inheriting from watchdog's PatternMatchingEventHandler
    patterns = ['*.xdrx', '*.eds', '*.txt']            # Events are only generated for these file types.

    def __init__(self):
        super(LabHandler, self).__init__()
        self.recent_events = deque('ghi', maxlen=30)
        self.error_message = deque(maxlen=1)
        self.v_counter = Counter(machine='Viia7')
        self.q_counter = Counter(machine='Qiaxcel')
        self._auto_export = True
        self._user_only = False
        self.user = os.getlogin()
    """
    The Observer passes events to the handler, which then calls functions based on the type of event
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

        time.sleep(1)  # wait here to allow file to be fully written - prevents some errors with os.stat
        if '.eds' in event.src_path and event.src_path not in self.recent_events:  # .eds files we haven't seen recently
            if self.is_large_enough(event.src_path):  # this is here instead of ^ to prevent double error message
                with v_lock:
                    self.v_counter.count = self.notif(event, self.v_counter.count)

    def on_created(self, event):
        """Called when a new file is created. aka Qiaxcel/ Export events."""

        if '.xdrx' in event.src_path and event.src_path not in self.recent_events:  # .xdrx file we haven't seen
            with q_lock:
                self.q_counter.count = self.notif(event, self.q_counter.count)
        if '.txt' in event.src_path and self.user in event.src_path\
                and "Export" in event.src_path and self._auto_export:
            time.sleep(0.3)       # wait here to allow file to be fully written # increase if timeout happens a lot
            try:
                export.new(event.src_path)
            except OSError:     # if team drive is being slow, wait longer.
                time.sleep(4)
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
            print(Message(message).timestamp(machine))

        elif not self._user_only:                           # non distinguished notification

            file = Message(file).white()
            message = ' {} has finished.'.format(file)
            if (self.q_counter.show and machine == 'Qiaxcel') or \
                    (self.v_counter.show and machine == 'Viia7'):

                print(Message(message).timestamp(machine))

        if x_counter == 1:         # If this was the run to notify on, inform that notification is now off.
            print(''.ljust(19, ' ') + machine + ' No longer notifying.')

        return x_counter - 1         # Increment counter down

    def auto_export(self):
        self._auto_export = not self._auto_export
        print('Auto export processing ' + Message('ON').green()) if self._auto_export\
            else print('Auto export processing ' + Message('OFF').red())

    def user_only(self):
        self._user_only = not self._user_only
        print('Displaying ' + Message('YOUR').white() + ' events only') if self._user_only\
            else print('Displaying ' + Message('ALL').white() + ' events')

    def show_all(self):
        self.v_counter.show = True
        self.q_counter.show = True
        self._user_only = False
        print('Displaying ' + Message('ALL').white() + ' events')

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
        new = ImageGrab.grabclipboard()
        assert (new.size[1] in {1575, 788, 504, 394})
        self._original = new

    def get(self):
        self.send_to_clipboard(self.crop(crop_type='standard'))

    def get_small(self):
        # send both scale and egel to clipboard, if using 'Office Clipboard', may paste both from clipboard history.
        self.send_to_clipboard(self.crop(crop_type='scale'), self.crop(crop_type='small'))

    def get_scale(self):
        self.send_to_clipboard(self.crop(crop_type='scale'))

    def crop(self, crop_type='standard'):
        # crop is a box within an image defined in pixels (left, top, right, bottom)
        crops = {'small': (self._original.size[1] / 5.54,
                           self._original.size[1] / 71.7,
                           self._original.size[0] - self._original.size[1] / 5.325,
                           self._original.size[1]),
                 'scale': (0,
                           self._original.size[1] / 71.7,
                           self._original.size[1] / 5.54,
                           self._original.size[1]),
                 'standard': (self._original.size[1] / 5.54,
                              self._original.size[1] / 71.7,
                              self._original.size[0] - self._original.size[1] / 168,
                              self._original.size[1])}
        img = self._original.crop(crops[crop_type])
        img = img.rotate(270, expand=True)
        img_out = BytesIO()
        img.convert("RGB").save(img_out, "BMP")
        img_final = img_out.getvalue()[14:]
        img_out.close()
        return img_final

    @staticmethod
    def send_to_clipboard(*args):
        for item in args:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, item)
            win32clipboard.CloseClipboard()
            time.sleep(0.05)  # small wait for Office Clipboard.
        print('Clipboard image processed.')


class ClipboardWatcher(Thread):
    def __init__(self):
        super(ClipboardWatcher, self).__init__()
        self._wait = 2.
        self._paused = False
        self._stopping = False
        self.image = Egel()

    def run(self):
        print('Clipboard watcher starting...')
        recent_value = win32clipboard.GetClipboardSequenceNumber()
        while not self._stopping:
            tmp_value = win32clipboard.GetClipboardSequenceNumber()
            if not tmp_value == recent_value and not self._paused:  # if the clipboard has changed
                recent_value = tmp_value
                try:
                    self.image.grab()
                    self.image.get()
                except AttributeError:
                    pass    # given when clipboard object is not an image. Ignore
                except AssertionError:
                    pass    # given when image is not the right size. Ignore
            else:
                time.sleep(self._wait)

    def toggle(self):
        self._paused = not self._paused
        if self._paused:
            self._wait = 10
            print('Clipboard watcher' + Message('OFF').red())
        else:
            self._wait = 2.
            print('Clipboard watcher resumed')

    def stop(self):
        self._stopping = True


class InputLoop(Thread):
    def __init__(self):
        super(InputLoop, self).__init__()
        self._stopping = False
        self.daemon = True
        self.instructions = {
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
                             self.startup: ['install', 'setup'],
                             self.startup_remove: ['delete', 'uninstall'],
                             self.stop: ['quit', 'exit', 'QQ', 'quti'],
                             self.print_help: ['help', 'hlep']
                             }

    def run(self):
        print('Running...\nEnter a command or type help')
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
                        print(labhandler.q_counter.notify_setting)
                if 'v' in inp:
                    if 'hide' in inp:
                        labhandler.v_counter.show = ''
                        print(labhandler.v_counter.show)
                        continue
                    with v_lock:
                        labhandler.v_counter.count = count
                        print(labhandler.v_counter.notify_setting)

    @staticmethod
    def get_input():
        inp = input('')
        if not inp:  # if blank input, read clipboard.
            inp = clipboard.clipboard_get()
        inp = os.path.normpath(inp.strip('\'"'))
        inp = inp.strip()
        return inp

    @staticmethod
    def print_help():
        # TODO: Add colours?
        print('GenoTools - ver 11.06.2019 - jb40'.center(45, ' ') + '\n'.ljust(45, '-'))
        print('Monitors Qiaxcel and Viia7 and notifies on run completion, '
              'and auto-processes Viia7 export files and Qiaxcel images.\n'
              'Files with your username will generate a distinguished notification.\n'
              'Commands may be given to notify you on other events, '
              'e.g. All runs, only your runs, or in n runs time.\n')
        print('Commands'.center(45, ' ') + '\n'.ljust(45, '-'))
        print('Press Enter to paste file path, or enter a command\n')
        print('Notifications'.center(45, ' ') + '\n'.ljust(45, '-'))
        print('Q or V - no number   : Toggle distinguished notifications for all events')
        print('Q or V - with number : Distinguished notification after n events.')
        print('Q or V + "hide"      : Toggle hide Qiaxcel or Viia7 events ')
        print('Mine                 : Display your events only.')
        print('All                  : Display all events.')
        print('Auto Processing'.center(45, ' ') + '\n'.ljust(45, '-'))
        print('Auto                 : Toggle auto-processing of export files')
        print('Multi                : Start/Stop Multi export mode (Toggle)')
        print('Done                 : Stop Multi export mode')
        print('ToFile               : Export to a pre-existing file (Paste the path)')
        print('Last                 : Export to the previously exported file')
        print('Images               : Toggle auto-processing of Qiaxcel images')
        print(''.ljust(45, '-'))
        print('Install      Setup   : Start program on Windows Startup')
        print('Uninstall    Delete  : Remove from Windows Startup')
        print('Quit         Exit    : Exit the program')
        print(''.ljust(45, '-'))

    @staticmethod
    def startup():
        print("Installing to Startup...")
        startup_path = os.path.expanduser("~\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\\")
        name = "Lab Helper.cmd"
        with open(startup_path + name, "w+") as f:
            f.write("@echo off\n")
            f.write("cls\n")
            if '.py' in argv[0]:
                f.write("start /MIN python " + argv[0])
            else:
                f.write("start /MIN " + argv[0])
        print("Done!")

    @staticmethod
    def startup_remove():
        print('Removing from Startup...')
        startup_file = os.path.expanduser("~\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs"
                                          "\\Startup\\Lab Helper.cmd")
        try:
            os.remove(startup_file)
            print("Removed.")
        except FileNotFoundError:
            print('File not found!')

    def stop(self):
        self._stopping = True
        observer.stop()


class StatusCheck(object):
    def __init__(self):
        self.date = time.strftime("%b %Y", time.localtime())

        self.qiaxcel_path = config['File paths']['QIAxcel']
        self.experiment_path = self.get_path("Experiments")  # Path to current month
        self.export_path = self.get_path("Export")
        self.experiment_path_last = self.get_path("Experiments", last_month=True)  # Path to last month
        self.export_path_last = self.get_path("Export", last_month=True)

        self.qiaxcel_watch = observer.schedule(labhandler, path=self.qiaxcel_path)
        self.experiment_watch_curr = observer.schedule(labhandler, path=self.experiment_path)
        self.export_watch_curr = observer.schedule(labhandler, path=self.export_path)
        self.experiment_watch_last = observer.schedule(labhandler, path=self.experiment_path_last)\
            if self.experiment_path_last is not None else False
        self.export_watch_last = observer.schedule(labhandler, path=self.export_path_last) \
            if self.export_path_last is not None else False
        self.message = deque(maxlen=2)  # This is used by thread_print to prevent duplicate messages from threads.

    def update_month(self):
        self.date = time.strftime("%b %Y", time.localtime())  # Update month
        print('The month has changed to ' + self.date)
        self.experiment_path = self.get_path("Experiments")  # Get new path
        self.export_path = self.get_path("Export")
        if self.experiment_watch_last:
            observer.unschedule(self.experiment_watch_last)  # Unschedule last month watch
        if self.export_watch_last:
            observer.unschedule(self.export_watch_last)
        self.experiment_watch_last = self.experiment_watch_curr  # Make current month watch last month
        self.export_watch_last = self.export_watch_curr
        self.experiment_watch_curr = observer.schedule(labhandler, self.experiment_path)  # Schedule new current month
        self.export_watch_curr = observer.schedule(labhandler, self.export_path)

    @staticmethod
    def check_update():
        config.read(os.path.normpath(os.path.dirname(argv[0]) + '/config.ini'))  # may have changed.
        start = datetime.strptime(config['Update']['Start'], '%d.%m.%Y %H:%M')
        end = datetime.strptime(config['Update']['End'], '%d.%m.%Y %H:%M')
        if start < datetime.now() < end:
            print('Update in progress until ' + datetime.strftime(end, "%d.%m.%y %H:%M "))
            print('Closing in 10 seconds...')
            time.sleep(10)
            raise KeyboardInterrupt

    def get_path(self, folder, last_month=False):
        makedirs = True
        date_t = date.today()
        if last_month:
            makedirs = False
            first = date_t.replace(day=1)                     # replace the day with the first of the month.
            date_t = first - timedelta(days=1)                # subtract one day to get the last day of last month
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
        if folder == "Experiments":
            return config['File paths']['Genotyping'] + 'qPCR ' + year + '\\Experiments\\' + month + ' ' + year
        if folder == "Export":
            return config['File paths']['Genotyping'] + 'qPCR ' + year + '\\Results Export\\' + month + ' ' + year

    def thread_print(self, msg):
        """
        Prevents duplicate print statements from threads. If it has been sent recently, it is not re sent.
        message is a deque with length 2.
        """
        if msg not in self.message:
            print(msg)
            self.message.append(msg)


class MyEmitter(observers.read_directory_changes.WindowsApiEmitter):

    def queue_events(self, timeout):
        try:
            super().queue_events(timeout)
        except OSError as e:
            status.thread_print(str(e))
            status.thread_print('Lost connection to team drive!')
            connected = False
            while not connected:
                try:
                    self.on_thread_start()  # need to re-set the directory handle.
                    connected = True
                    status.thread_print('Reconnected!')
                except OSError:
                    time.sleep(10)
                    status.thread_print('Reconnecting...')


if __name__ == '__main__':
    colorama_init()  # Init colorama to enable coloured text output via ANSI escape codes on windows console.
    q_lock = Lock()  # Locks used when reading or writing q_cnt or v_cnt since they are in multiple threads.
    v_lock = Lock()
    # m_lock = Lock()
    # message = None
    config = configparser.ConfigParser()
    config.read(os.path.normpath(os.path.dirname(argv[0]) + '/config.ini'))  # config.ini = ANSI

    egel_watcher = ClipboardWatcher()  # Instantiate classes
    labhandler = LabHandler()
    export = Export.Export()
    observer = BaseObserver(emitter_class=MyEmitter, timeout=DEFAULT_OBSERVER_TIMEOUT)
    in_loop = InputLoop()
    status = StatusCheck()

    observer.start()  # Start threads
    egel_watcher.start()
    in_loop.start()

    try:
        while True:  # Check if something has changed
            if status.date != time.strftime("%b %Y", time.localtime()):  # If month has changed.
                status.update_month()
            for i in range(300):
                if not observer.is_alive():
                    raise KeyboardInterrupt
                status.check_update()
                time.sleep(4)
    except KeyboardInterrupt:  # on keyboard interrupt (Ctrl + C)
        observer.stop()  # Stop observer + Threads (if alive)
        egel_watcher.stop()
        print('\nbye!')
