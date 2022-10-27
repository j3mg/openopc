python setup.py sdist

python setup.py bdist_wheel --plat-name=win32
rd /S/Q build

REM Build 64-bit specific wheels that do not include OpenOPCService

python setup.py bdist_wheel --plat-name=win_amd64
python setup.py bdist_wheel --plat-name=manylinux1_x86_64
