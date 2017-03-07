from distutils.core import setup
import py2exe

setup(windows=[{
    'script':'master.py',
    'uac_info':'requireAdministrator',
},])
