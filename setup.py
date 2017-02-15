from distutils.core import setup
import py2exe

setup(
    windows = [{
        'script': 'master.py',
        "icon_resources": [(1, "pycel.ico")],
    }],
    options = {
        "py2exe":{
            "unbuffered": True,
            "optimize": 2,
            "bundle_files": 1,
        }
    }
)
