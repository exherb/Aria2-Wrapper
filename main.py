#!/usr/bin/env python
# coding=utf-8

import os
import sys
import subprocess
import json
import Tkinter as tk
import tkFileDialog as filedialog
from PIL import Image, ImageTk
import psutil


def _is_windows_x64():
    return 'PROGRAMFILES(X86)' in os.environ


def _get_aria2_bin():
    platform = sys.platform
    if hasattr(sys, 'frozen') and sys.platform == 'win32':
        basis = sys.executable
        if _is_windows_x64():
            platform = 'win64'
    else:
        basis = __file__
    return os.path.realpath(os.path.join(os.path.dirname(basis), 'aria2',
                                         platform,
                                         'aria2c' +
                                         ('.exe'
                                          if sys.platform == 'win32' else '')))


def _get_image(name):
    if hasattr(sys, 'frozen') and sys.platform == 'win32':
        basis = sys.executable
    else:
        basis = __file__
    return os.path.join(os.path.dirname(basis), 'images', name)


def _get_app_path():
    if hasattr(sys, 'frozen'):
        if sys.platform == 'win32':
            return sys.executable
        elif sys.platform == 'darwin':
            return os.path.dirname(os.path.dirname(
                                   os.path.dirname(sys.executable)))
    return __file__


def _registry_as_startup(app_path):
    if not hasattr(sys, 'frozen'):
        raise NotImplementedError()
    app_path = os.path.realpath(app_path)
    app_name, _ = os.path.splitext(os.path.basename(app_path))
    if sys.platform == 'win32':
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        startup_dir = shell.SpecialFolders("Startup")

        path = os.path.join(startup_dir,
                            app_name + '.lnk')
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = app_path
        shortcut.WorkingDirectory = os.path.dirname(app_path)
        shortcut.save()
    elif sys.platform == 'darwin':
        status = subprocess.call('osascript -e \'tell application ' +
                                 '"System Events" to make login item at end ' +
                                 'with properties ' +
                                 '{{path:"{}", hidden:false}}\''.
                                 format(app_path), shell=True)
        if status != 0:
            raise RuntimeError()
    else:
        raise NotImplementedError()


def _remove_startup(app_path):
    if not hasattr(sys, 'frozen'):
        raise NotImplementedError()
    app_name, ext = os.path.splitext(os.path.basename(app_path))
    if sys.platform == 'win32':
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        startup_dir = shell.SpecialFolders("Startup")
        path = os.path.join(startup_dir,
                            app_name + '.lnk')
        os.remove(path)
    elif sys.platform == 'darwin':
        status = subprocess.call('osascript -e \'tell application ' +
                                 '"System Events" to delete login item "{}"\''.
                                 format(app_name),
                                 shell=True)
        if status != 0:
            raise RuntimeError()
    else:
        raise NotImplementedError()


def _is_in_startup(app_path):
    app_name, _ = os.path.splitext(os.path.basename(app_path))
    if sys.platform == 'win32':
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        startup_dir = shell.SpecialFolders("Startup")
        path = os.path.join(startup_dir,
                            app_name + '.lnk')
        return os.path.exists(path)
    elif sys.platform == 'darwin':
        p = subprocess.Popen('osascript -e \'tell application ' +
                             '"System Events" to get the name of every ' +
                             'login item\'',
                             shell=True,
                             stdout=subprocess.PIPE,
                             stderr=subprocess.PIPE)
        out, _ = p.communicate()
        p.wait()
        return out and app_name in out
    else:
        raise NotImplementedError()


def _get_aria2_process(aria2_bin):
    for process in psutil.process_iter():
        try:
            if aria2_bin in process.cmdline()[0]:
                return process
        except Exception:
            pass
    return None


def _terminate_aria2_process(aria2_bin, wait=True):
    running_process = _get_aria2_process(aria2_bin)
    if running_process:
        running_process.terminate()


def _get_config_path(config_name):
    if sys.platform == 'win32':
        configs_path = os.path.join(os.path.dirname(_get_app_path()),
                                    'configs')
    elif sys.platform == 'darwin':
        configs_path = os.path.join(os.path.expanduser('~'), '.aria2-wrapper')
    else:
        raise NotImplementedError()
    if not os.path.exists(configs_path):
        os.makedirs(configs_path)
    return os.path.join(configs_path, config_name)


def _load_setting():
    settings_path = _get_config_path('settings.json')
    if not os.path.exists(settings_path):
        return {}
    return json.load(open(settings_path, 'r'))


def _save_setting(settings):
    json.dump(settings, open(_get_config_path('settings.json'), 'w'))


def _change_aria2_state(state, output_dir, rpc_secret):
    if not output_dir:
        output_dir = os.path.join(os.path.expanduser('~'),
                                  'Downloads')
    aria2_bin = _get_aria2_bin()
    _terminate_aria2_process(aria2_bin)
    if state:
        session_file = _get_config_path('aria2.session')
        args = [aria2_bin, '--enable-rpc',
                '--rpc-listen-all=true',
                '--rpc-allow-origin-all',
                '--continue=true',
                '--save-session={}'.
                format(session_file),
                '--dir={}'.format(output_dir)]
        if os.path.exists(session_file):
            args.append('--input-file={}'.format(session_file))
        if rpc_secret:
            args.append('--rpc-secret={}'.format(rpc_secret))
        subprocess.Popen(args, close_fds=True)


def _show_preferences():
    settings = _load_setting()

    window = tk.Tk()
    window.resizable(0, 0)
    window.title('Aria2 Desktop Wrapper')
    window.attributes('-topmost', 1)
    screenwidth = window.winfo_screenwidth()
    screenheight = window.winfo_screenheight()
    width = 405
    height = 200
    if sys.platform == 'win32':
        height = height + 30
    x = (screenwidth - width)*0.5
    y = (screenheight - height)*0.5
    window.geometry('{}x{}+{}+{}'.format(width, height, int(x), int(y)))
    window.is_picking_file = False

    store_dir = settings.get('dir',
                             os.path.join(os.path.expanduser('~'),
                                          'Downloads'))
    if not os.path.exists(store_dir):
        try:
            os.makedirs(store_dir)
        except Exception:
            store_dir = os.path.expanduser('~')
    store_dir = tk.StringVar(window, store_dir)
    store_rpc_secret = tk.StringVar(window, settings.get('rpc-secret', ''))

    run_with_system = tk.BooleanVar(window, settings.get('startup', False))
    aria2_started = tk.BooleanVar(window,
                                  _get_aria2_process(_get_aria2_bin())
                                  is not None)

    frame = tk.Frame(window, padx=30, pady=15)
    frame.pack(fill='both', expand=True)

    row = 0

    on_image = ImageTk.PhotoImage(Image.open(_get_image('on.jpg')))
    off_image = ImageTk.PhotoImage(Image.open(_get_image('off.jpg')))
    aria2_switcher = tk.Label(frame,
                              image=(on_image if aria2_started.get()
                                     else off_image))
    aria2_switcher.grid(row=row, column=0, sticky='w')
    aria2_switcher.on_image = on_image
    aria2_switcher.on_image = off_image

    on_text = 'Aria2 is running...'
    off_text = 'Aria2 is not running'
    aria2_state = tk.Label(frame,
                           text=(on_text
                                 if aria2_started.get() else
                                 off_text))
    aria2_state.grid(row=row, column=1, columnspan=2, sticky='w')

    def on_aria2_switched(event):
        state = not aria2_started.get()
        _change_aria2_state(state, store_dir.get(),
                            store_rpc_secret.get())
        if state:
            aria2_state['text'] = on_text
        else:
            aria2_state['text'] = off_text
        aria2_started.set(state)
        event.widget['image'] = on_image if state else off_image
    aria2_switcher.bind('<ButtonRelease-1>', on_aria2_switched)

    row += 1
    tk.Label(frame, text='Put downloads in:').grid(row=row, column=0,
                                                   sticky='w')
    tk.Entry(frame, textvariable=store_dir).grid(row=row, column=1,
                                                 sticky='we')

    def on_select_downloads_directory():
        if window.is_picking_file:
            return
        window.is_picking_file = True
        path = filedialog.askdirectory(title='Select downloads directory')
        print(path)
        if path:
            store_dir.set(path)
        window.is_picking_file = False
    tk.Button(frame, text='Set', command=on_select_downloads_directory).\
        grid(row=row, column=2, sticky='e', pady=10)

    row += 1
    tk.Label(frame, text='RPC Token:').grid(row=row, column=0,
                                            sticky='e')
    tk.Entry(frame, textvariable=store_rpc_secret).grid(row=row, column=1,
                                                        sticky='we')

    row += 1
    seperator = tk.Canvas(frame, height=2)
    seperator.create_line(0, 5, width, 5, fill='#e3e3e3', width=2)
    seperator.grid(row=row, column=0, columnspan=3, sticky='we')

    def on_startup_state_changed():
        status = run_with_system.get()
        try:
            app_path = _get_app_path()
            if status:
                _registry_as_startup(app_path)
            else:
                _remove_startup(app_path)
        except Exception:
            run_with_system.set(not status)

    row += 1
    tk.Checkbutton(frame,
                   variable=run_with_system,
                   command=on_startup_state_changed,
                   text='Start aria2 wrapper at login').grid(row=row, column=0,
                                                             columnspan=2,
                                                             sticky='w')

    def on_done():
        window.destroy()
    tk.Button(frame, text='Done', command=on_done).\
        grid(row=row, column=2, sticky='e', pady=10)

    def on_destroty(event):
        if event.widget != window:
            return
        settings['dir'] = store_dir.get()
        settings['rpc-secret'] = store_rpc_secret.get()
        settings['startup'] = run_with_system.get()
        _save_setting(settings)
    window.bind("<Destroy>", on_destroty)

    window.lift()
    window.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == 'preferences':
        _show_preferences()
    else:
        settings = _load_setting()
        if settings.get('startup', None) is None:
            try:
                _registry_as_startup(_get_app_path())
                settings['startup'] = True
                _save_setting(settings)
            except Exception:
                pass
        _change_aria2_state(True, settings.get('dir', None),
                            settings.get('rpc-secret', None))

        def _start_preferences():
            if hasattr(sys, 'frozen'):
                if sys.platform == 'darwin':
                    subprocess.Popen([os.path.join(_get_app_path(),
                                                   'Contents',
                                                   'MacOS',
                                                   'Arial2 Wrapper'),
                                      'preferences'], close_fds=True)
                elif sys.platform == 'win32':
                    subprocess.Popen([os.path.join(_get_app_path(),
                                                   'Arial2 Wrapper.exe'),
                                      'preferences'], close_fds=True)
            else:
                subprocess.Popen([os.path.realpath(_get_app_path()),
                                  'preferences'], close_fds=True)

        if sys.platform == 'darwin':
            import rumps
            rumps._NOTIFICATIONS = False

            class Aria2WrapperApp(rumps.App):
                def __init__(self):
                    super(Aria2WrapperApp, self).__init__('Aria2',
                                                          quit_button=None)
                    self.menu = ['Aria2', rumps.separator,
                                 'Preferences',
                                 rumps.separator,
                                 'Quit']
                    self.set_aria2_state(True)

                def set_aria2_state(self, state):
                    item = self.menu['Aria2']
                    item.state = state
                    if state:
                        item.title = 'Running...'
                        self.icon = 'images/menubar_on.png'
                    else:
                        item.title = 'Not running'
                        self.icon = 'images/menubar_off.png'

                def change_aria2_state(self, state):
                    settings = _load_setting()
                    _change_aria2_state(state, settings.get('dir', None),
                                        settings.get('rpc-secret', None))
                    self.set_aria2_state(state)

                @rumps.timer(1)
                def refresh_aria2_state(self, sender):
                    self.set_aria2_state(_get_aria2_process(
                                         _get_aria2_bin()) is not None)

                @rumps.clicked('Aria2')
                def aria2_switcher(self, sender):
                    self.change_aria2_state(not sender.state)

                @rumps.clicked('Preferences')
                def prefs(self, _):
                    _start_preferences()

                @rumps.clicked('Quit')
                def quit(self, sender):
                    _terminate_aria2_process(_get_aria2_bin(), False)
                    rumps.quit_application(sender)
            Aria2WrapperApp().run()
        elif sys.platform == 'win32':
            import win32api
            import win32con
            import win32gui_struct
            try:
                import winxpgui as win32gui
            except ImportError:
                import win32gui

            def non_string_iterable(obj):
                try:
                    iter(obj)
                except TypeError:
                    return False
                else:
                    return not isinstance(obj, str)

            class SysTrayIcon(object):
                QUIT = 'Quit'
                SPECIAL_ACTIONS = [QUIT]

                FIRST_ID = 1023

                def __init__(self,
                             icon,
                             hover_text,
                             menu_options,
                             on_quit=None,
                             default_menu_index=None,
                             window_class_name=None,):

                    self.icon = icon
                    self.hover_text = hover_text
                    self.on_quit = on_quit

                    menu_options = menu_options + (('Quit', None, self.QUIT),)
                    self._next_action_id = self.FIRST_ID
                    self.menu_actions_by_id = set()
                    self.menu_options = self.\
                        _add_ids_to_menu_options(list(menu_options))
                    self.menu_actions_by_id = dict(self.menu_actions_by_id)
                    del self._next_action_id

                    self.default_menu_index = (default_menu_index or 0)
                    self.window_class_name = window_class_name or\
                        "SysTrayIconPy"

                    message_map = {win32gui.
                                   RegisterWindowMessage("TaskbarCreated"):
                                   self.restart,
                                   win32con.WM_DESTROY: self.destroy,
                                   win32con.WM_COMMAND: self.command,
                                   win32con.WM_USER + 20: self.notify, }
                    # Register the Window class.
                    window_class = win32gui.WNDCLASS()
                    hinst = window_class.hInstance = win32gui.\
                        GetModuleHandle(None)
                    window_class.lpszClassName = self.window_class_name
                    window_class.style = win32con.CS_VREDRAW |\
                        win32con.CS_HREDRAW
                    window_class.hCursor = win32gui.\
                        LoadCursor(0, win32con.IDC_ARROW)
                    window_class.hbrBackground = win32con.COLOR_WINDOW
                    window_class.lpfnWndProc = message_map
                    classAtom = win32gui.RegisterClass(window_class)
                    # Create the Window.
                    style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
                    self.hwnd = win32gui.CreateWindow(classAtom,
                                                      self.window_class_name,
                                                      style,
                                                      0,
                                                      0,
                                                      win32con.CW_USEDEFAULT,
                                                      win32con.CW_USEDEFAULT,
                                                      0,
                                                      0,
                                                      hinst,
                                                      None)
                    win32gui.UpdateWindow(self.hwnd)
                    self.notify_id = None
                    self.refresh_icon()

                    win32gui.PumpMessages()

                def _add_ids_to_menu_options(self, menu_options):
                    result = []
                    for menu_option in menu_options:
                        option_text, option_icon, option_action = menu_option
                        if callable(option_action) or\
                           option_action in self.SPECIAL_ACTIONS:
                            self.menu_actions_by_id.add((self._next_action_id,
                                                         option_action))
                            result.append(menu_option +
                                          (self._next_action_id,))
                        elif non_string_iterable(option_action):
                            result.append((option_text,
                                           option_icon,
                                           self.
                                           _add_ids_to_menu_options
                                           (option_action),
                                           self._next_action_id))
                        else:
                            print('Unknown item', option_text, option_icon,
                                  option_action)
                        self._next_action_id += 1
                    return result

                def refresh_icon(self):
                    # Try and find a custom icon
                    hinst = win32gui.GetModuleHandle(None)
                    if os.path.isfile(self.icon):
                        icon_flags = win32con.LR_LOADFROMFILE |\
                            win32con.LR_DEFAULTSIZE
                        hicon = win32gui.LoadImage(hinst,
                                                   self.icon,
                                                   win32con.IMAGE_ICON,
                                                   0,
                                                   0,
                                                   icon_flags)
                    else:
                        print("Can't find icon file - using default.")
                        hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

                    if self.notify_id:
                        message = win32gui.NIM_MODIFY
                    else:
                        message = win32gui.NIM_ADD
                    self.notify_id = (self.hwnd,
                                      0,
                                      win32gui.NIF_ICON |
                                      win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                                      win32con.WM_USER+20,
                                      hicon,
                                      self.hover_text)
                    win32gui.Shell_NotifyIcon(message, self.notify_id)

                def restart(self, hwnd, msg, wparam, lparam):
                    self.refresh_icon()

                def destroy(self, hwnd, msg, wparam, lparam):
                    if self.on_quit:
                        self.on_quit(self)
                    nid = (self.hwnd, 0)
                    win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
                    win32gui.PostQuitMessage(0)  # Terminate the app.

                def notify(self, hwnd, msg, wparam, lparam):
                    if lparam == win32con.WM_LBUTTONDBLCLK:
                        self.execute_menu_option(self.default_menu_index +
                                                 self.FIRST_ID)
                    elif lparam == win32con.WM_RBUTTONUP:
                        self.show_menu()
                    elif lparam == win32con.WM_LBUTTONUP:
                        pass
                    return True

                def show_menu(self):
                    menu = win32gui.CreatePopupMenu()
                    self.create_menu(menu, self.menu_options)
                    #win32gui.SetMenuDefaultItem(menu, 1000, 0)

                    pos = win32gui.GetCursorPos()
                    # See http://msdn.microsoft.com/library/default.asp?url=
                    # /library/en-us/winui/menus_0hdi.asp
                    win32gui.SetForegroundWindow(self.hwnd)
                    win32gui.TrackPopupMenu(menu,
                                            win32con.TPM_LEFTALIGN,
                                            pos[0],
                                            pos[1],
                                            0,
                                            self.hwnd,
                                            None)
                    win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

                def create_menu(self, menu, menu_options):
                    for option_text, option_icon, option_action,\
                            option_id in menu_options[::-1]:
                        if option_icon:
                            option_icon = self.prep_menu_icon(option_icon)

                        if option_id in self.menu_actions_by_id:
                            item, extras = win32gui_struct.\
                                PackMENUITEMINFO(text=option_text,
                                                 hbmpItem=option_icon,
                                                 wID=option_id)
                            win32gui.InsertMenuItem(menu, 0, 1, item)
                        else:
                            submenu = win32gui.CreatePopupMenu()
                            self.create_menu(submenu, option_action)
                            item, extras = win32gui_struct.\
                                PackMENUITEMINFO(text=option_text,
                                                 hbmpItem=option_icon,
                                                 hSubMenu=submenu)
                            win32gui.InsertMenuItem(menu, 0, 1, item)

                def prep_menu_icon(self, icon):
                    # First load the icon.
                    ico_x = win32api.GetSystemMetrics(win32con.SM_CXSMICON)
                    ico_y = win32api.GetSystemMetrics(win32con.SM_CYSMICON)
                    hicon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON,
                                               ico_x, ico_y,
                                               win32con.LR_LOADFROMFILE)

                    hdcBitmap = win32gui.CreateCompatibleDC(0)
                    hdcScreen = win32gui.GetDC(0)
                    hbm = win32gui.CreateCompatibleBitmap(hdcScreen, ico_x,
                                                          ico_y)
                    hbmOld = win32gui.SelectObject(hdcBitmap, hbm)
                    # Fill the background.
                    brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)
                    win32gui.FillRect(hdcBitmap, (0, 0, 16, 16), brush)
                    win32gui.DrawIconEx(hdcBitmap, 0, 0, hicon, ico_x, ico_y,
                                        0, 0,
                                        win32con.DI_NORMAL)
                    win32gui.SelectObject(hdcBitmap, hbmOld)
                    win32gui.DeleteDC(hdcBitmap)

                    return hbm

                def command(self, hwnd, msg, wparam, lparam):
                    id = win32gui.LOWORD(wparam)
                    self.execute_menu_option(id)

                def execute_menu_option(self, id):
                    menu_action = self.menu_actions_by_id[id]
                    if menu_action == self.QUIT:
                        win32gui.DestroyWindow(self.hwnd)
                    else:
                        menu_action(self)

            menu_options = (('Aria2', None, None),
                            ('Preferences', None, _start_preferences))

            def set_aria2_state(systray, state):
                systray.state = state
                if state:
                    systray.icon = 'images/menubar_on.png'
                else:
                    systray.icon = 'images/menubar_off.png'
                systray.refresh_icon()

            def change_aria2_state(systray, state):
                settings = _load_setting()
                _change_aria2_state(state, settings.get('dir', None),
                                    settings.get('rpc-secret', None))
                set_aria2_state(state)

            def quit(_):
                _terminate_aria2_process(_get_aria2_bin(), False)

            SysTrayIcon('images/menubar_on.png', "Aria2 Wrapper",
                        menu_options,
                        on_quit=quit,
                        default_menu_index=1)
