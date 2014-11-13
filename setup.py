#!/usr/bin/env python
# coding=utf-8

import sys
import os
from glob import glob
from setuptools import setup

name = "Arial2 Wrapper"
mainscript = 'main.py'


data_files = [('aria2/{}'.format(sys.platform),
               glob('aria2/{}/*'.format(sys.platform))),
              ('images/',
               glob('images/*'))]
if sys.platform == 'win32':
    data_files.append(('aria2/win64',
                       glob('aria2/win64'.format(sys.platform))))


if sys.platform == 'darwin':
    sys.path.append('/Library/Python/2.7/site-packages')
    extra_options = dict(
        setup_requires=['py2app'],
        app=[mainscript],
        options={'py2app': {'argv_emulation': False,
                            'iconfile': 'icos/icon.icns',
                            'plist': {
                                'CFBundleName': name,
                                'CFBundleShortVersionString': '1.0.0',
                                'CFBundleVersion': '1.0.0',
                                'LSUIElement': True
                                },
                            'packages': ['rumps']
                            }})
elif sys.platform == 'win32':
    import py2exe
    extra_options = dict(
        setup_requires=['py2exe'],
        windows=[{'script': mainscript,
                  'icon_resources': [(0, 'icos/icon.ico')]}],
        bundle_files=2
    )
setup(
    name=name,
    data_files=data_files,
    **extra_options
)

if sys.platform == 'win32':
    target = 'dist/{}.exe'.format(name)
    if os.path.exists(target):
        os.remove(target)
    os.rename('dist/gui.exe', target)
