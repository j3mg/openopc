###########################################################################
#
# OpenOPC Gateway Service
#
# A Windows service providing remote access to the OpenOPC library.
#
# Copyright (c) 2007-2012 Barry Barnreiter (barry_b@users.sourceforge.net)
# Copyright (c) 2014 Anton D. Kachalov (mouse@yandex.ru)
# Copyright (c) 2017 José A. Maita (jose.a.maita@gmail.com)
#
###########################################################################

import win32serviceutil
import win32service
import win32event
import servicemanager
import winerror
import select
import socket
import os
import sys
import time
import OpenOPC

try:
    import Pyro5.core
    import Pyro5.client
    import Pyro5.server
except ImportError:
    print('Pyro5 module required (https://pypi.python.org/pypi/Pyro5)')
    exit()

Pyro5.config.SERVERTYPE='thread'

opc_class = os.getenv('OPC_CLASS', OpenOPC.OPC_CLASS)
opc_gate_host = os.getenv('OPC_GATE_HOST', 'localhost')
opc_gate_port = int(os.getenv('OPC_GATE_PORT', 7766))


@Pyro5.server.expose    # needed for version 5.12
class opc(object):
    def __init__(self):
        self._remote_hosts = {}
        self._init_times = {}
        self._tx_times = {}

    def get_clients(self):
        """Return list of server instances as a list of (GUID,host,time) tuples"""
        hlist = []

        if self._pyroDaemon != None:
            reg1 = Pyro5.server.DaemonObject(self._pyroDaemon).registered()   # needed for version 5.12
            reg2 = [si for si in reg1 if si.find('obj_') == 0]
            reg = ["PYRO:{0}@{1}:{2}".format(obj, opc_gate_host, opc_gate_port) for obj in reg2]
            hosts = self._remote_hosts
            init_times = self._init_times
            tx_times = self._tx_times
            hlist = [(hosts[k] if k in hosts else '', init_times[k], tx_times[k]) for k in reg]

        return hlist

    def create_client(self, weakRegister=False): # Requires Pyro5 5.13 or above
        """Create a new OpenOPC instance in the Pyro server"""

        opc_obj = OpenOPC.client(opc_class)
        uri = self._pyroDaemon.register(opc_obj, weak=weakRegister)

        uuid = uri.__str__() # undocumented Pyro function
        opc_obj.set_gateway_settings(self, opc_obj, opc_gate_host, opc_gate_port, uuid)

        remote_ip = uuid # self.getLocalStorage().caller.addr[0]
        self._remote_hosts[uuid] = '%s' % (remote_ip)
        self._init_times[uuid] =  time.time()
        self._tx_times[uuid] =  time.time()
        return Pyro5.client.Proxy(uri)

    def release_client(self, obj):
        """Release an OpenOPC instance in the Pyro server"""
        """Called in client.close() method"""
        """REQUIRED unless object is created with weak register"""

        self._pyroDaemon.unregister(obj)
        del self._remote_hosts[obj.GUID()]
        del self._init_times[obj.GUID()]
        del self._tx_times[obj.GUID()]
        del obj

class OpcService(win32serviceutil.ServiceFramework):
    _svc_name_ = "zzzOpenOPCService"
    _svc_display_name_ = "OpenOPC Gateway Service"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)

    def SvcStop(self):
        servicemanager.LogInfoMsg('\n\nStopping service')
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogInfoMsg('\n\nStarting service on host %s port %d' % (opc_gate_host ,opc_gate_port))

        daemon = Pyro5.server.Daemon(host=opc_gate_host, port=opc_gate_port)
        daemon.register(opc(), "opc")

        socks = daemon.sockets
        while win32event.WaitForSingleObject(self.hWaitStop, 0) != win32event.WAIT_OBJECT_0:
            ins,outs,exs = select.select(socks,[],[],1)
            if ins:
                daemon.events(ins)

        daemon.shutdown()

def main():
    if len(sys.argv) == 1:
        try:
            evtsrc_dll = os.path.abspath(servicemanager.__file__)
            servicemanager.PrepareToHostSingle(OpcService)
            servicemanager.Initialize('zzzOpenOPCService', evtsrc_dll)
            servicemanager.StartServiceCtrlDispatcher()
        except win32service.error as details:
            if details.winerror == winerror.ERROR_FAILED_SERVICE_CONTROLLER_CONNECT:
                win32serviceutil.usage()
                print(' --foreground: Run OpenOPCService in foreground.')

    else:
        if sys.argv[1] == '--foreground':
            daemon = Pyro5.server.Daemon(host=opc_gate_host, port=opc_gate_port)
            daemon.register(opc(), 'opc')

            socks = set(daemon.sockets)
            while True:
                ins,outs,exs = select.select(socks,[],[],1)
                if ins:
                    daemon.events(ins)

            daemon.shutdown()
        else:
            win32serviceutil.HandleCommandLine(OpcService)

if __name__ == '__main__':
    main()
