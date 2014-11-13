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
        _change_aria2_state(True, settings['dir'],
                            settings.get('rpc-secret', None))
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
                    _change_aria2_state(state, settings['dir'],
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
                    if hasattr(sys, 'frozen'):
                        subprocess.Popen([os.path.join(_get_app_path(),
                                                       'Contents',
                                                       'MacOS',
                                                       'Arial2 Wrapper'),
                                          'preferences'], close_fds=True)
                    else:
                        subprocess.Popen([os.path.realpath(_get_app_path()),
                                          'preferences'], close_fds=True)

                @rumps.clicked('Quit')
                def quit(self, sender):
                    _terminate_aria2_process(_get_aria2_bin(), False)
                    rumps.quit_application(sender)
            Aria2WrapperApp().run()
