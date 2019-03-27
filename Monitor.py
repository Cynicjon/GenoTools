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
    from colorama import Fore, Style, init as colorama_init
    from pandas.io import clipboard
    from PIL import ImageGrab
    from io import BytesIO
    import win32clipboard
except ModuleNotFoundError:
    input("Please install the requirements! \nPress Enter to exit.")
    quit()

    # TODO: Conditional formatting sometimes isn't inserted into the first sheet of a multi export
    #       Adjust the scale to fit exactly 2 rows. - Done, Needs testing
    #       Occasionally get numbers formatted as text error - if v large numbers present
    #       Update Help text with egel instructions

    # TODO: Upload to GitHub
    #       Push Egel update


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


class LabHandler(events.PatternMatchingEventHandler):  # inheriting from watchdog's PatternMatchingEventHandler
    patterns = ["*.xdrx", "*.eds", '*.txt']            # Events are only generated for these file types.
    recent_events = deque('ghi', maxlen=30)            # A queue to prevent duplicate events. 30 is arbitrary.
    viia7_str = Fore.CYAN + 'Viia7' + Style.RESET_ALL + ':  '       # Colour codes for Viia7.
    qiaxcel_str = Fore.MAGENTA + 'Qiaxcel' + Style.RESET_ALL + ':'  # Colour codes for Qiaxcel
    v_counter = Counter(machine=viia7_str)
    q_counter = Counter(machine=qiaxcel_str)
    _auto_export = True
    _user_only = False
    _multi_export = False
    _to_file = ''
    user = os.getlogin()

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
        if '.txt' in event.src_path and self.user in event.src_path \
                and "Export" in event.src_path and self._auto_export:
            time.sleep(0.3)       # wait here to allow file to be fully written # increase if timeout happens a lot
            try:
                file_out = Export.auto_to_xl(event.src_path, main_file=self._to_file)
            except OSError:     # if team drive is being slow, wait longer.
                time.sleep(4)
                file_out = Export.auto_to_xl(event.src_path, main_file=self._to_file)
            if self._multi_export and self.to_file == '':
                self._to_file = file_out
            if not self._multi_export:
                try:
                    os.startfile(file_out)
                except TypeError:   # If the txt file is bad, export.py returns a NoneType, which causes this exception
                    pass

    def notif(self, event, x_counter):
        """Prints a notification about the event to console. May be normal or distinguished.
         Distinguished notifications flash the console window.
         Notif will be distinguished if users windows login is in path or x_counter = 1 or > 9 (now/ always setting)"""

        self.recent_events.append(event.src_path)  # Add to recent events queue to prevent duplicate notifications.
        machine, file = self.get_event_info(event)

        if os.getlogin() in event.src_path.lower() or x_counter == 1 or x_counter > 9:  # distinguished notif
            file = Fore.GREEN + file + Style.RESET_ALL      # file name in green
            message = Fore.GREEN + '>>> {}'.format(machine).ljust(21, ' ') + ' {} has finished!'.format(file)
            ctypes.windll.user32.FlashWindow(ctypes.windll.kernel32.GetConsoleWindow(), True)  # Flash console window
            print(self.bright_time() + message)

        elif not self._user_only:                           # non distinguished notification
            file = Style.BRIGHT + file + Style.RESET_ALL    # file name in white
            message = ' -  {}'.format(machine).ljust(21, ' ') + ' {} has finished.'.format(file)
            if (self.q_counter.show and machine == self.qiaxcel_str) or \
                    (self.v_counter.show and machine == self.viia7_str):
                print(self.bright_time() + message)

        if x_counter == 1:         # If this was the run to notify on, inform that notification is now off.
            print(''.ljust(19, ' ') + machine + ' No longer notifying.')

        return x_counter - 1         # Increment counter down

    def auto_export(self):
        self._auto_export = not self._auto_export
        print('Auto export processing ON') if self._auto_export else print('Auto export processing OFF')

    def multi_export(self):
        self._multi_export = not self._multi_export
        if not self._multi_export and self._to_file:
            os.startfile(self._to_file)  # Try/except shouldn't be needed here.
            print('Multi Export processing complete: ' + os.path.split(self._to_file)[1])
            self._to_file = ''
        print('Multi export processing ON') if self._multi_export else print('Multi export processing OFF')

    @property
    def to_file(self):
        return self._to_file

    @to_file.setter  # Setter not needed since no logic??
    def to_file(self, path):
        self._to_file = path

    def user_only(self):
        self._user_only = not self._user_only
        print('Displaying your events ONLY') if self._user_only else print('Displaying ALL events')

    def show_all(self):
        self.v_counter.show = True
        self.q_counter.show = True
        self._user_only = False
        print('Displaying ALL events')

    def get_event_info(self, event):
        """Returns the machine name and file path."""
        machine = self.viia7_str if event.event_type == 'modified' else self.qiaxcel_str
        file = str(os.path.splitext(event.src_path)[0].split('\\')[-1])  # Get file name
        return machine, file

    @staticmethod
    def bright_time():
        """Returns the current time formatted nicely, flanked by ANSI escape codes for bright text."""
        return Style.BRIGHT + time.strftime("%d.%m.%y %H:%M ", time.localtime()) + Style.NORMAL

    def is_large_enough(self, path):
        """Determines if the file given by path is above 1300000 bytes"""
        try:
            return os.stat(path).st_size > 1300000
        except (FileNotFoundError, OSError) as e:
            file = str(os.path.splitext(path)[0].split('\\')[-1])
            print(Fore.RED, e, Style.RESET_ALL)  # If not, print error, assume True.
            print(self.bright_time() + ' -  {}'.format(self.viia7_str).ljust(21, ' ') + file + " wasn't saved properly!"
                  " You'll need to analyse and save the run again from the machine.")
            return True  # Better to inform than not. I think this happens when .eds isn't saved or is deleted?


class Egel(object):

    _original = ''

    def grab(self):
        new = ImageGrab.grabclipboard()
        assert (new.size[1] in {1575, 788, 504, 394})
        self._original = new

    def egel(self):
        self.send_to_clipboard(self.crop(crop_type='standard'))

    def small_egel(self):
        # send both scale and egel to clipboard, if using 'Office Clipboard', may paste both from clipboard history.
        self.send_to_clipboard(self.crop(crop_type='scale'), self.crop(crop_type='small'))

    def scale(self):
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
    def __init__(self, test, task):
        super(ClipboardWatcher, self).__init__()
        self._test = test
        self._task = task
        self._wait = 2.
        self._paused = False
        self._stopping = False

    def run(self):
        print('Clipboard watcher starting...')
        recent_value = win32clipboard.GetClipboardSequenceNumber()
        while not self._stopping:
            tmp_value = win32clipboard.GetClipboardSequenceNumber()
            if not tmp_value == recent_value and not self._paused:  # if the clipboard has changed
                recent_value = tmp_value
                try:
                    self._test()
                    self._task()
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
            print('Clipboard watcher stopped')
        else:
            self._wait = 2.
            print('Clipboard watcher resumed')

    def stop(self):
        self._stopping = True


def get_input():
    inp = input('')
    if not inp:                 # if blank input, read clipboard.
        inp = clipboard.clipboard_get()
    inp = os.path.normpath(inp.strip('\'"'))
    inp = inp.strip()
    return inp


def input_loop():  # TODO: finish this. Add option to turn off file start.
    """A loop that asks user for input. Responds to q (Qiaxcel), v (Viia7) and a single number or some combination.
    Any digit after the first is ignored. e.g: q, q5, v4, qv3. q54 == q5.
    no number will toggle notify all, a number will notify after n events."""
    instructions = {'mine': labhandler.user_only,
                    'all': labhandler.show_all,
                    'auto': labhandler.auto_export,
                    'multi': labhandler.multi_export,
                    'to file': to_file, 'tofile': to_file, 'file': to_file,
                    'egels': egel_watcher.toggle,
                    'scale': egel.scale, 'egel': egel.egel, 'small': egel.small_egel,
                    'install': startup, 'setup': startup, 'add': startup,
                    'remove': startup_remove, 'delete': startup_remove, 'uninstall': startup_remove,
                    'quit': quit_script, 'exit': quit_script, 'QQ': quit_script,
                    'help': print_help
                    }
    try:
        print('Running...\nEnter a command or type help')
        while True:
            inp = get_input()
            if os.path.isfile(inp):
                try:
                    os.startfile(Export.auto_to_xl(inp, main_file=labhandler.to_file))
                except TypeError:  # If the txt file is bad, export.py returns a NoneType, which causes this exception
                    pass
            else:
                inp = inp.lower()
                if inp in instructions:
                    instructions[inp]()
                    continue
                for char in inp:
                    if char.isdigit():      # if a digit is in input, set count = that digit and break.
                        count = int(char)
                        break               # break out of loop after first digit.
                else:
                    count = 0
                if 'q' in inp:
                    if 'hide' in inp:
                        labhandler.q_counter.show = ''
                        print(labhandler.q_counter.show)
                        continue
                    with q_lock:            # lock to make referencing variable shared between threads safe.
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
    except EOFError:                        # This error is given when exiting script only. Ignore.
        pass


def to_file():
    labhandler.multi_export()
    print('Enter a target file (.xlsx) or type stop to cancel')
    while True:
        inp = get_input()
        if os.path.isfile(inp) and inp[-5:] == '.xlsx':
            labhandler.to_file = inp
            break
        if inp.lower() == 'stop':
            labhandler.multi_export()
            break
        else:
            print('That isn\'t an excel file path!')
    if labhandler.to_file:
        print('Thanks. You can now export your files, or paste the file path here.')


def quit_script():
    print('Exiting program...')
    observer.stop()


def print_help():
    print('Lab Helper - ver 27.03.2019 - jb40'.center(45, ' ') + '\n'.ljust(45, '-'))
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
    print('Multi                : Start/Finish Multi export mode.')
    print('ToFile               : Export to a pre-existing file (paste the path)')
    print('Egels                : Toggle auto-processing of Qiaxcel images')
    print(''.ljust(45, '-'))
    print('Install      Setup   : Start program on Windows Startup')
    print('Uninstall    Delete  : Remove from Windows Startup')
    print('Quit         Exit    : Exit the program')
    print(''.ljust(45, '-'))


def get_path(folder, last_month=False):
    makedirs = True
    date_t = date.today()
    if last_month:
        makedirs = False
        first = date_t.replace(day=1)                     # replace the day with the first of the month.
        date_t = first - timedelta(days=1)                # subtract one day to get the last day of last month
    month, year = date_t.strftime('%b %Y').split(' ')
    path = build_path(month, year, folder)
    if os.path.isdir(path):
        return path
    month_full = date_t.strftime('%B')
    path2 = build_path(month_full, year, folder)
    if os.path.isdir(path2):
        return path2
    if month == 'Sep':
        path2 = build_path('Sept', year, folder)
        if os.path.isdir(path2):
            return path2
    if makedirs:
        os.makedirs(path, exist_ok=True)
        return path
    print("Cant find last month's " + folder + " path")
    return None


def build_path(month, year, folder):
    if folder == "Experiments":
        return config['File paths']['Genotyping'] + 'qPCR ' + year + '\\Experiments\\' + month + ' ' + year
    if folder == "Export":
        return config['File paths']['Genotyping'] + 'qPCR ' + year + '\\Results Export\\' + month + ' ' + year


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


def startup_remove():
    print('Removing from Startup...')
    startup_file = os.path.expanduser("~\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs"
                                      "\\Startup\\Lab Helper.cmd")
    try:
        os.remove(startup_file)
    except FileNotFoundError:
        pass
    print("Removed.")


def check_update():
    config.read(os.path.normpath(os.path.dirname(argv[0]) + '/config.ini'))  # may have changed.
    start = datetime.strptime(config['Update']['Start'], '%d.%m.%Y %H:%M')
    end = datetime.strptime(config['Update']['End'], '%d.%m.%Y %H:%M')
    if start < datetime.now() < end:
        print('Update in progress until ' + datetime.strftime(end, "%d.%m.%y %H:%M "))
        print('Closing in 10 seconds...')
        time.sleep(10)
        raise KeyboardInterrupt


if __name__ == '__main__':
    colorama_init()            # Init colorama to enable coloured text output via ANSI escape codes on windows console.
    q_lock = Lock()            # Locks used when reading or writing q_cnt or v_cnt since they are in multiple threads.
    v_lock = Lock()

    current_month_year = time.strftime("%b %Y", time.localtime())     # This is checked later to see if month changed.
    config = configparser.ConfigParser()
    config.read(os.path.normpath(os.path.dirname(argv[0]) + '/config.ini'))  # todo encoding utf-8?atm config.ini = ANSI
    qiaxcel_path = config['File paths']['QIAxcel']
    viia7_path = get_path("Experiments")                              # Path to current month Viia7 experiments.
    export_path = get_path("Export")
    viia7_path_last = get_path("Experiments", last_month=True)        # Path to last month Viia7 experiments.
    export_path_last = get_path("Export", last_month=True)

    egel = Egel()
    egel_watcher = ClipboardWatcher(egel.grab, egel.egel)

    labhandler = LabHandler()                                         # Instantiate Handler
    observer = observers.Observer()                                   # Instantiate Observer
    qiaxcel_watch = observer.schedule(labhandler, path=qiaxcel_path)  # Schedule watch locations, and use our Handler.
    viia7_watch_curr = observer.schedule(labhandler, path=viia7_path)
    export_watch_curr = observer.schedule(labhandler, path=export_path)

    viia7_watch_last = observer.schedule(labhandler, path=viia7_path_last) if viia7_path_last is not None else False
    export_watch_last = observer.schedule(labhandler, path=export_path_last)if export_path_last is not None else False

    observer.start()                                                  # Start Observer thread.
    egel_watcher.start()

    in_loop = Thread(target=input_loop)
    in_loop.daemon = True
    in_loop.start()

    try:
        while True:  # Check if something has changed
            if current_month_year != time.strftime("%b %Y", time.localtime()):  # If month has changed.
                current_month_year = time.strftime("%b %Y", time.localtime())   # Update month
                print('The month has changed to ' + current_month_year)
                viia7_path = get_path("Viia7")                                  # Get new  path
                export_path = get_path("Export")
                if viia7_watch_last:
                    observer.unschedule(viia7_watch_last)                       # Unschedule last month watch
                if export_watch_last:
                    observer.unschedule(export_watch_last)
                viia7_watch_last = viia7_watch_curr                             # Make current month watch last month
                export_watch_last = export_watch_curr
                viia7_watch_curr = observer.schedule(labhandler, viia7_path)    # Schedule new current month
                export_watch_curr = observer.schedule(labhandler, export_path)
            for i in range(600):
                if not observer.is_alive():
                    raise KeyboardInterrupt
                check_update()
                time.sleep(2)
    except KeyboardInterrupt:                                                   # on keyboard interrupt (Ctrl + C)
        observer.stop()                                                         # Stop observer + Threads (if alive)
        egel_watcher.stop()
        print('\nbye!')
