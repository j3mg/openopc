###########################################################################
#
# OpenOPC for Python Library Module
#
# Copyright (c) 2007-2012 Barry Barnreiter (barry_b@users.sourceforge.net)
# Copyright (c) 2014 Anton D. Kachalov (mouse@yandex.ru)
# Copyright (c) 2017 José A. Maita (jose.a.maita@gmail.com)
# Copyright (c) 2022 j3mg
#
###########################################################################

import os
from Pyro5 import __version__ as _pyro_version_
import Pyro5.core
import Pyro5.server
from OpenOPC.common import OPC_CLASS, OPC_CLIENT, OPC_SERVER, OPCError, TimeoutError # Required for the OpenOPC Gateway service
from OpenOPC.opcda import client as client

__version__ = '1.5.1'

def win32_check(): # can be mocked by pytest-mock
    return (os.name == 'nt')

def parse_version(ver):
    return tuple(map(int, (ver.split(".")))) # Works with Pyro5 numeric version strings


if parse_version("5.13.1") > parse_version(_pyro_version_):
    raise ImportError("Pyro version must be greater than 5.13")

def get_sessions(host='localhost', port=7766):
    """Return sessions in OpenOPC Gateway Service as GUID:host hash"""

    import Pyro5.client
    server_obj = Pyro5.client.Proxy("PYRO:opc@{0}:{1}".format(host, port))
    return server_obj.get_clients()

def open_client(host='localhost', port=7766):
    """Connect to the specified OpenOPC Gateway Service"""

    import Pyro5.client
    server_obj = Pyro5.client.Proxy("PYRO:opc@{0}:{1}".format(host, port))
    return server_obj.create_client()

