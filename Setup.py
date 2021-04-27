import os
from distutils.core import setup
from glob import glob
import py2exe
import PySide2

dirname = os.path.dirname(PySide2.__file__)
plugin_path = os.path.join(dirname, 'plugins', 'platforms')
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = plugin_path

data_files = (
    ('', glob(r'C:\Windows\SYSTEM32\msvcp100.dll')),
    ('', glob(r'C:\Windows\SYSTEM32\msvcr100.dll')),
    ('platforms',
     glob(r'C:\repos_python\DistractionFreeOEC\venv\Lib\site-packages\PySide2\plugins\platforms\qwindows.dll')),
)

setup(
    name='Distraction Free OEC',
    windows=[{"script": 'App.py', 'dest_base': 'Distraction Free'}],
    data_files=data_files,
    # 'windows' means it's a GUI, 'console' 
    # It's a console program, 'service' a Windows' service, 'com_server' is for a COM server
    # You can add more and py2exe will compile them separately.
    options={
        # This is the list of options each module has, for example py2exe,
        # but for example, PyQt or django could also contain specific options
        'py2exe': {
            'dist_dir': 'dist',  # The output folder

            'compressed': True,  # If you want the program to be compressed to be as small as possible
            'includes': ['sys', 'os', 'openpyxl', 'PySide2',
                         'PySide2.QtGui', 'PySide2.QtCore', 'PySide2.QtWidgets'],
            # All the modules you need to be included,
            # I added packages such as PySide 
            # and psutil but also custom ones like modules
            # and utils inside it because py2exe 
            # guesses which modules are being used by
            # the file we want to compile, but not the imports,
            # so if you import something inside main.py which
            # also imports something, it might break.
        }
    },

)
