"""
Unit tests for OpenOPC.open_client
Requires:
        pytest
        KEPware.KEPserverEx.V4 OPC simulator
        Matrikon.OPC.Automation dll registered
"""
import pytest
import OpenOPC

def test_open_connect(mock_opc_service, mock_pyro_daemon):
    gateway = "127.0.0.1"
    opchost = "127.0.0.1"
    opc = OpenOPC.open_client(gateway)
    connected = opc.connect('KEPware.KEPserverEx.V4', opchost)
    assert(connected == True)
 
def test_guid(mock_opc_service, mock_pyro_daemon):
    gateway = "127.0.0.1"
    opchost = "127.0.0.1"
    opc = OpenOPC.open_client(gateway)
    guid = opc.GUID()
    assert(guid == "PYRO:obj_dda39746d43a4c8b960022e6fbfd3137@127.0.0.1:7766")
 
def test_get_clients(mock_opc_service, mock_pyro_daemon):
    gateway = "127.0.0.1"
    opchost = "127.0.0.1"
    opc = OpenOPC.open_client(gateway)
    connected = opc.connect('KEPware.KEPserverEx.V4', opchost)
    list = OpenOPC.get_sessions(gateway)
    assert(len(list))
    