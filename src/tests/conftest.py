"""
PyTest configuration file containing fixtures and monkeypatching functions and classes
"""
import Pyro5.client
import win32serviceutil
import win32event
import win32com.client
import pythoncom
import servicemanager
import pytest
import OpenOPCDA
import OpenOPCService
import time
import Common

__opcServer__ = None
__opcClient__ = None

def pytest_configure():
    pytest.opcServer = __opcServer__
    pytest.opcClient = __opcClient__ # reuse the OPC client for multiple tests to prevent Windows threading exception

def pytest_runtest_setup(item: pytest.Item) -> None:
    if item.name == 'test_read':
        pytest.opcClient = OpenOPCDA.client()
        pytest.opcClient.connect('KEPware.KEPserverEx.V4')
        pytest.opcClient.__open_guid__ = "PYRO:obj_dda39746d43a4c8b960022e6fbfd3137@127.0.0.1:7766"

def pytest_runtest_teardown(item: pytest.Item) -> None:
    if item.name == 'test_get_clients':
        pytest.opcClient.close()

class mock_daemon():
    def register(self, dummyobj, weak=False):
        uri = "PYRO:obj_dda39746d43a4c8b960022e6fbfd3137@127.0.0.1:7766"
        return uri
    def unregister(self, dummyobj):
        pass

@pytest.fixture
def mock_opc(monkeypatch):
    class _mock_opc():
        def __init__(self):
            self._remote_hosts = {}
            self._init_times = {}
            self._tx_times = {}
            self._pyroDaemon = mock_daemon()
            
    monkeypatch.setattr(OpenOPCService.opc, "__init__", _mock_opc.__init__)
    
@pytest.fixture
def mock_pyro_client(monkeypatch):
    class _mock_pyro(OpenOPCDA.client):
        def __init__(self, uri):
            super().__init__()
            self.__open_serv__ = pytest.opcServer 
            self.__open_self__ = self
            self.__open_host__ = "127.0.0.1"
            self.__open_port__ = 7766
            self.__open_guid__ = "PYRO:obj_dda39746d43a4c8b960022e6fbfd3137@127.0.0.1:7766"
        def connect(self, opc_server, opc_host):
            return True

    monkeypatch.setattr("Pyro5.client.Proxy", _mock_pyro)

@pytest.fixture
def mock_pyro_daemon(monkeypatch):
    class _mock_pyro():
        def __init__(self, daemon):
            pass
        def registered(self):
            return ["Pyro.Daemon", "opc", "obj_dda39746d43a4c8b960022e6fbfd3137"]
            
    monkeypatch.setattr("Pyro5.server.DaemonObject", _mock_pyro)
    
@pytest.fixture
def mock_opc_service(monkeypatch):
    class _mock_service(OpenOPCService.opc):
        def __init__(self, pyroobj):
            OpenOPCService.opc.__init__(self)
            self._pyroDaemon = mock_daemon()
        def get_clients(self):
            uuid = "PYRO:obj_dda39746d43a4c8b960022e6fbfd3137@127.0.0.1:7766"
            remote_ip = uuid
            self._remote_hosts[uuid] = '%s' % (remote_ip)
            self._init_times[uuid] =  time.time()
            self._tx_times[uuid] =  time.time()
            return super().get_clients()
        def create_client(self, weakRegister=False):
            return pytest.opcClient

    monkeypatch.setattr("Pyro5.client.Proxy", _mock_service)

    
@pytest.fixture
def mock_win32_service(monkeypatch):
    class _mock_service_framework():
        def __init__(self, args):
            pass

        def ReportServiceStatus(self, arg):
            pass
    
    class _mock_service_manager():
        def LogInfoMsg(self, *argv):
            pass
            
    class _mock_win32event():
        def SetEvent(self):
            pass
        def WaitForSingleObject(self, evt):
            return win32event.WAIT_OBJECT_0
    
    class _mock_daemon():
        def __init__(self, host, port):
            pass
        def register(self, arg1, arg2):
            pass
        def shutdown(self):
            pass
        @property
        def sockets(self):
            return None
        
    monkeypatch.setattr(win32serviceutil.ServiceFramework, "__init__", _mock_service_framework.__init__)
    monkeypatch.setattr(win32serviceutil.ServiceFramework, "ReportServiceStatus", _mock_service_framework.ReportServiceStatus)
    monkeypatch.setattr(servicemanager, "LogInfoMsg", _mock_service_manager.LogInfoMsg)
    monkeypatch.setattr(win32event, "SetEvent", _mock_win32event.SetEvent)
    monkeypatch.setattr(win32event, "WaitForSingleObject", _mock_win32event.WaitForSingleObject)
    monkeypatch.setattr("Pyro5.server.Daemon", _mock_daemon)
