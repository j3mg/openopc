from setuptools import setup

setup(
    description=" OPC (OLE for Process Control) toolkit for Python 3.x",
    install_requires=['Pyro4>=4.61'],
    keywords='python, opc, openopc',
    license='GPLv2',
    maintainer = 'Michal Kwiatkowski',
    maintainer_email = 'michal@trivas.pl',
    name="OpenOPC-Python3x",
    package_dir={'':'src'},
    py_modules=['OpenOPC'],
    url='https://github.com/mkwiatkowski/openopc',
    version="1.3.1",
)
