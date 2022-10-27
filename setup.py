''' Supported platforms
    Linux x86_64
    Windows win32 and x86_64
    
    Requires setuptools >=43.0.0 for platform_system attribute
'''
import sys
from setuptools import setup

# OpenOPCService can only be installed on Windows systems
noservicepackage = len(sys.argv) > 2  and (sys.argv[2] == "--plat-name=manylinux1_x86_64" or sys.argv[2] == "--plat-name=win_amd64")

packagelist = ["OpenOPC"] if noservicepackage else ["OpenOPC", "OpenOPCService"]
entries= {'console_scripts': ['opc=OpenOPC.opc:main'] if noservicepackage else ['opc=OpenOPC.opc:main', 'opc-service=OpenOPCService.opcservice:main'] }
requires= ['Pyro5>=5.13.1'] if noservicepackage else ['Pyro5>=5.13.1', 'pywin32>=221; platform_system=="Windows"'] # pywin32>=221 allowed for Python 3.4 compatible version of pywin32

setup(packages=packagelist,
        install_requires=requires,
        entry_points=entries)
