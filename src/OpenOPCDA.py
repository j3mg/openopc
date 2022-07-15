###########################################################################
#
# OpenOPC for Python OPC-DA Library Module
#
# A Windows only OPC-DA library module.
#
# Copyright (c) 2007-2012 Barry Barnreiter (barry_b@users.sourceforge.net)
# Copyright (c) 2014 Anton D. Kachalov (mouse@yandex.ru)
# Copyright (c) 2017 José A. Maita (jose.a.maita@gmail.com)
# Copyright (c) 2022 j3mg
#
###########################################################################
import os

def win32_check(): # can be mocked by pytest-mock
    return (os.name == 'nt')

win32com_found = False
win32_found = win32_check()

# Win32 only modules not needed for 'open' protocol mode
if win32_found:
    try:
        import win32com.client
        import win32com.server.util
        import win32event
        import pythoncom
        import pywintypes
        import SystemHealth

        # Win32 variant types
        pywintypes.datetime = pywintypes.TimeType
        vt = dict([(pythoncom.__dict__[vtype], vtype) for vtype in pythoncom.__dict__.keys() if vtype[:2] == "VT"])

        # Allow gencache to create the cached wrapper objects
        win32com.client.gencache.is_readonly = False

        # Under p2exe the call in gencache to __init__() does not happen
        # so we use Rebuild() to force the creation of the gen_py folder
        win32com.client.gencache.Rebuild(verbose=0)

    # So we can work on Windows in "open" protocol mode without the need for the win32com modules
    except ImportError:
        win32com_found = False
    else:
        from OpenOPCDAIO import ClientIO
        from OpenOPCDATools import ClientTools
        
        win32com_found = True

import sys
import time
import types
import string
import socket
import re
import Pyro5.core
import Pyro5.server

# OPC Constants
from Common import OPC_CLASS, OPC_SERVER, OPC_CLIENT, OPCError, get_error_str

@Pyro5.server.expose    # needed for 5.12
class client():
    def __init__(self, opc_class=None, client_name=None):
        """Instantiate OPC automation class"""
        
        self.opc_server = None
        self.opc_host = None
        self.client_name = client_name
        self._groups = {}
        self._group_tags = {}
        self._group_valid_tags = {}
        self._group_server_handles = {}
        self._group_handles_tag = {}
        self._group_hooks = {}
        # Members prefixed by __open are owned by the OPC Gateway Service
        self.__open_serv__ = None 
        self.__open_self__ = None
        self.__open_host__ = None
        self.__open_port__ = None
        self.__open_guid__ = None

        self.trace = None
        def win32_init(opc_class):
            pythoncom.CoInitialize()

            if opc_class == None:
                if 'OPC_CLASS' in os.environ:
                    opc_class = os.environ['OPC_CLASS']
                else:
                    opc_class = OPC_CLASS

            opc_class_list = opc_class.split(';')

            for i,c in enumerate(opc_class_list):
                try:
                    self._opc = win32com.client.gencache.EnsureDispatch(c, 0)
                    self.opc_class = c
                    break
                except pythoncom.com_error as err:
                    if i == len(opc_class_list)-1:
                        error_msg = 'Dispatch: %s' % get_error_str(err)
                        raise OPCError(error_msg)
            self._event = win32event.CreateEvent(None,0,0,None)

            self.clientIO = ClientIO()
            self.clientTools = ClientTools(self.opc_class, self.opc_host)

        self.win32os = win32com_found and win32_check()
        if self.win32os: win32_init(opc_class)

    def set_gateway_settings(self, service, instance, host, port, guid):
        """ For the use of the OPC Gateway Service """
        
        self.__open_serv__ = service 
        self.__open_self__ = instance
        self.__open_host__ = host
        self.__open_port__ = port
        self.__open_guid__ = guid
        if self.win32os: self.clientTools.set_gateway_settings(service, host, port, guid)

    def set_trace(self, trace):
        if self.__open_serv__ == None:
            self.trace = trace
            if self.win32os: self.clientIO.setTrace(trace)

    def connect(self, opc_server=None, opc_host='localhost'):
        """Connect to the specified OPC server"""
        
        def win32_connect(opc_server, opc_host):
            pythoncom.CoInitialize()
            if opc_server == None:
                # Initial connect using environment vars
                if self.opc_server == None:
                    if 'OPC_SERVER' in os.environ:
                        opc_server = os.environ['OPC_SERVER']
                    else:
                        opc_server = OPC_SERVER
                # Reconnect using previous server name
                else:
                    opc_server = self.opc_server
                    opc_host = self.opc_host

            opc_server_list = opc_server.split(';')
            connected = False

            for s in opc_server_list:
                if len(s):
                    try:
                        if self.trace: self.trace('Connect(%s,%s)' % (s, opc_host))
                        self._opc.Connect(s, opc_host)
                    except pythoncom.com_error as err:
                        if len(opc_server_list) == 1:
                            error_msg = 'Connect: %s' % self._get_error_str(err)
                            raise OPCError(error_msg)
                    else:
                        # Set client name since some OPC servers use it for security
                        try:
                            if self.client_name == None:
                                if 'OPC_CLIENT' in os.environ:
                                    self._opc.ClientName = os.environ['OPC_CLIENT']
                                else:
                                    self._opc.ClientName = OPC_CLIENT
                            else:
                                self._opc.ClientName = self.client_name
                        except:
                            pass
                        connected = True
                        break

            if not connected:
                raise OPCError('Connect: Cannot connect to any of the servers in the OPC_SERVER list')

            # With some OPC servers, the next OPC call immediately after Connect()
            # will occationally fail.  Sleeping for 1/100 second seems to fix this.
            time.sleep(0.01)

            self.opc_server = opc_server
            if opc_host == 'localhost':
                opc_host = socket.gethostname()
            self.opc_host = opc_host
            self.clientTools.set_opc_host(opc_host)
            return connected
        
        connected = False
        if self.win32os:
            connected = win32_connect(opc_server, opc_host)
        return connected

    def GUID(self):
        return self.__open_guid__

    def close(self, del_object=True):
        """Disconnect from the currently connected OPC server"""

        def win32_close(del_object):
            try:
                pythoncom.CoInitialize()
                self.remove(self.groups())

            except pythoncom.com_error as err:
                error_msg = 'Disconnect: %s' % get_error_str(err, self._opc)
                raise OPCError(error_msg)

            except OPCError:
                pass

            finally:
                if self.trace: self.trace('Disconnect()')
                self._opc.Disconnect()

                # Remove this object from the open gateway service
                if self.__open_serv__ and del_object:
                    self.__open_serv__.release_client(self.__open_self__)

        if self.win32os: win32_close(del_object)

    def groups(self):
        """Return a list of active tag groups"""
        return self._groups.keys()


    #
    # Read/Write functions
    #
    def iread(self, tags=None, group=None, size=None, pause=0, source='hybrid', update=-1, timeout=5000, sync=False, include_error=False, rebuild=False):
        if self.win32os:
            return self.clientIO.iread(self._opc, self.clientTools, tags, group, size, pause, source, update, timeout, sync, include_error, rebuild)
        else:
            return None

    def read(self, tags=None, group=None, size=None, pause=0, source='hybrid', update=-1, timeout=5000, sync=False, include_error=False, rebuild=False):
        if self.win32os:
            return self.clientIO.read(self._opc, self.clientTools, tags, group, size, pause, source, update, timeout, sync, include_error, rebuild)
        else:
            return None

    def _read_health(self, tags):
        if self.win32os:
            return self.clientIO._read_health(self.clientTools, tags)
        else:
            return None

    def iwrite(self, tag_value_pairs, size=None, pause=0, include_error=False):
        if self.win32os:
            return self.clientIO.iwrite(self._opc, self.clientTools, tag_value_pairs, size, pause, include_error)
        else:
            return None

    def write(self, tag_value_pairs, size=None, pause=0, include_error=False):
        if self.win32os:
            return self.clientIO.write(self._opc, self.clientTools, tag_value_pairs, size, pause, include_error)
        else:
            return None

    def remove(self, groups):
        if self.win32os:
            return self.clientIO.remove(self._opc, groups)
        else:
            return None

    def __getitem__(self, key):
        if self.win32os:
            return self.clientIO.__getitem__(self._opc, self.clientTools, key)
        else:
            return None

    def __setitem__(self, key, value):
        if self.win32os:
            return self.clientIO.__setitem__(self._opc, self.clientTools, key, value)
        else:
            return None
    #
    # General tools functions
    #
    def iproperties(self, tags, id=None):
        if self.win32os:
            return self.clientTools.iproperties(self._opc, tags, id)
        else:
            return None

    def properties(self, tags, id=None):
        if self.win32os:
            return self.clientTools.properties(self._opc, tags, id)
        else:
            return None

    def ilist(self, paths='*', recursive=False, flat=False, include_type=False):
        if self.win32os:
            return self.clientTools.ilist(self._opc, paths, recursive, flat, include_type)
        else:
            return None

    def list(self, paths='*', recursive=False, flat=False, include_type=False):
        if self.win32os:
            return self.clientTools.list(self._opc, paths, recursive, flat, include_type)
        else:
            return None

    def servers(self, opc_host='localhost'):
        if self.win32os:
            return self.clientTools.servers(self._opc, opc_host)
        else:
            return None

    def info(self):
        if self.win32os:
            return self.clientTools.info(self._opc)
        else:
            return None

    def ping(self):
        if self.win32os:
            return self.clientTools.ping(self._opc)
        else:
            return None

    def _get_error_str(self, err):
        if self.win32os:
            return get_error_str(err, self._opc)
        else:
            return None

    def _update_tx_time(self):
        if self.win32os:
            return self.clientTools._update_tx_time()
        else:
            return False

