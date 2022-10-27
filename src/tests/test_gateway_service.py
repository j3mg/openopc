"""
Unit tests for OpenOPCService.opc and OpenOPCService.OpenOPCService class
Requires:
        pytest
        pytest.opcClient from test_client.test_read
        KEPware.KEPserverEx.V4 OPC simulator
        Matrikon.OPC.Automation dll registered
"""
import pytest
import OpenOPCService

def test_server_connect(mock_opc, mock_pyro_client, mock_pyro_daemon):
    gateway = "127.0.0.1"
    opchost = "127.0.0.1"
    server = OpenOPCService.opc()
    pytest.opcServer = server
    taglistni = ['Server Variables.Monitor Value 1', 'Server Variables.Monitor Value 2']
    serverclient = server.create_client(gateway)
    connected = serverclient.connect('KEPware.KEPserverEx.V4', opchost)
    assert(connected == True)
    
def test_server_get_clients(mock_opc, mock_pyro_client, mock_pyro_daemon):
    gateway = "127.0.0.1"
    opchost = "127.0.0.1"
    server = OpenOPCService.opc()
    pytest.opcServer = server
    taglistni = ['Server Variables.Monitor Value 1', 'Server Variables.Monitor Value 2']
    serverclient = server.create_client(gateway)
    connected = serverclient.connect('KEPware.KEPserverEx.V4', opchost)
    clientList = server.get_clients()
    assert(len(clientList))

def test_release_client(mock_opc, mock_pyro_client, mock_pyro_daemon):
    gateway = "127.0.0.1"
    opchost = "127.0.0.1"
    server = OpenOPCService.opc()
    pytest.opcServer = server
    serverclient = server.create_client(gateway)
    server.release_client(serverclient)
    
def test_service(mock_win32_service):
    service = OpenOPCService.opcservice.OpcService('dummy')
    service.SvcDoRun()
    service.SvcStop()
    assert(True)