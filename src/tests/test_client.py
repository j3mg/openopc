"""
Unit tests for OpenOPC.client class
Requires:
        pytest
        pytest-mock
        KEPware.KEPserverEx.V4 OPC simulator
        Matrikon.OPC.Automation dll registered
"""
import os
import pytest
import OpenOPC
import pythoncom
from OpenOPC.common import OPCError

def test_badclassconnect():
    with pytest.raises( (pythoncom.com_error, Exception) ) as exc_info:
        opc = OpenOPC.client(opc_class='badclass')
        connected = opc.connect()
    assert('Dispatch: Invalid class string' in str(exc_info.value))

def test_badserverconnect():
    with pytest.raises( (pythoncom.com_error, Exception) ) as exc_info:
        opc = OpenOPC.client()
        connected = opc.connect(opc_server='dummy')
    assert('Connect: -2147467259' in str(exc_info.value))

def test_noserverconnect():
    with pytest.raises(BaseException) as exc_info:
        opc = OpenOPC.client()
        connected = opc.connect(opc_server="")
    assert('Connect: Cannot connect to any of the servers in the OPC_SERVER list' in str(exc_info.value))

def test_noconnecttoserverinlist():
    with pytest.raises(BaseException) as exc_info:
        opc = OpenOPC.client()
        connected = opc.connect(opc_server="opc.deltav.1;AIM.OPC.1;Yokogawa.ExaopcDAEXQ.1")
    assert('Connect: Cannot connect to any of the servers in the OPC_SERVER list' in str(exc_info.value))

def test_nonwindowsos(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    opc = OpenOPC.client()
    connected = opc.connect()
    if connected:
        opc.close()
    assert (connected == False)

def test_nonwindowsiread(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    opc = OpenOPC.client()
    response = opc.iread(taglistkep)
    assert(response == None)

def test_nonwindowsread(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    opc = OpenOPC.client()
    response = opc.read(taglistkep)
    assert(response == None)

def test_nonwindowsremove(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    opc = OpenOPC.client()
    response = opc.remove(groups='Channel_1.Device_1')
    assert(response == None)

def test_nonwindowswrite(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    tagPair = [('Channel_1.Device_1.Tag_1', 8), ('Channel_1.Device_1.Tag_2', 62), ('Channel_1.Device_1.Tag_3', 74)]
    opc = OpenOPC.client()
    response = opc.write(tagPair)
    assert(response == None)

def test_nonwindowsiwrite(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    tagPair = [('Channel_1.Device_1.Tag_1', 8), ('Channel_1.Device_1.Tag_2', 62), ('Channel_1.Device_1.Tag_3', 74)]
    opc = OpenOPC.client()
    response = opc.iwrite(tagPair)
    assert(response == None)

def test_nonwindowsproperties(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    opc = OpenOPC.client()
    response = opc.properties(taglistkep)
    assert(response == None)

def test_nonwindowsiproperties(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    opc = OpenOPC.client()
    response = opc.iproperties(taglistkep)
    assert(response == None)

def test_nonwindowsoslist(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    opc = OpenOPC.client()
    response = opc.list()
    assert(response == None)
    
def test_nonwindowsosilist(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    opc = OpenOPC.client()
    response = opc.ilist()
    assert(response == None)
    
def test_nonwindowsos_servers(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    opc = OpenOPC.client()
    response = opc.servers('127.0.0.1')
    assert(response == None)

def test_nonwindowsos_getitem(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    tagkep = 'Channel_1.Device_1.Tag_1'
    opc = OpenOPC.client()
    response = opc.__getitem__(tagkep)
    assert(response == None)
    
def test_nonwindowsos_setitem(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    tagkep = 'Channel_1.Device_1.Tag_1'
    opc = OpenOPC.client()
    response = opc.__setitem__(tagkep, 10)
    assert(response == None)
    
def test_nonwindowsoshealth(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    healthtags = ['@MemFree', '@MemUsed', '@MemTotal', '@MemPercent', '@DiskFree', '@SineWave', '@SawWave']
    opc = OpenOPC.client()
    response = opc._read_health(healthtags)
    assert(response == None)
    
def test_nonwindowsos_info(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    opc = OpenOPC.client()
    response = opc.info()
    assert(response == None)

def test_nonwindowsos_ping(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    opc = OpenOPC.client()
    response = opc.ping()
    assert(response == None)

def test_nonwindowsos_get_error_str(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    opc = OpenOPC.client()
    response = opc._get_error_str(err='dummy')
    assert(response == None)

def test_nonwindowsos_update_tx_time(mocker):
    mocker.patch.object(OpenOPC.opcda, 'win32com_found', False)
    opc = OpenOPC.client()
    response = opc._update_tx_time()
    assert(response == False)

def test_connect():
    opc = OpenOPC.client(client_name='OPCTest')
    connected = opc.connect()
    opc.close()
    assert(connected == True)
    
def test_connectnoenvironment(monkeypatch):
    monkeypatch.delenv("OPC_CLASS", raising=True)
    monkeypatch.delenv("OPC_SERVER", raising=True)
    monkeypatch.delenv("OPC_CLIENT", raising=True)
    opc = OpenOPC.client()
    connected = opc.connect()
    opc.close()
    assert(connected == True)
    
def test_reconnect():
    opc = OpenOPC.client()
    connected = opc.connect()
    opc.close()
    connected = opc.connect()
    assert(connected == True)


#
# An OPC client is reused for the tests below
#    
def test_read():
    def traceFunc(str):
        print(str)
    pytest.opcClient.set_trace(traceFunc)
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    tags = pytest.opcClient.read(taglistkep)
    assert(len(tags) == 4)

def test_readsized():
    def traceFunc(str):
        print(str)
    pytest.opcClient.set_trace(traceFunc)
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1']
    tags = pytest.opcClient.read(taglistkep, size=2)
    assert(len(tags) == 2)

def test_readsync():
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    tags = pytest.opcClient.read(taglistkep, sync=True)
    assert(len(tags) == 4)

def test_readincludeerror():
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    tags = pytest.opcClient.read(taglistkep, include_error=True)
    assert(len(tags) == 4 and 'The operation completed successfully.' in tags[0][4])

def test_readsingleincludeerror():
    tagkep = 'Channel_1.Device_1.Bool_1'
    tag = pytest.opcClient.read(tagkep, include_error=True)
    assert(len(tag) == 4 and 'The operation completed successfully.' in tag[3] )

def test_groupinitializeread():
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    tags = pytest.opcClient.read(taglistkep, group='Device1Group')
    assert(len(tags) == 4 and tags[0][2] == 'Good' and tags[1][2] == 'Good' and tags[2][2] == 'Good' and tags[3][2] == 'Good')

def test_twosetoftagsreads():
    taglistkep = ['Channel_1.Device_2.Bool_1', 'Channel_1.Device_2.Bool_2', 'Channel_1.Device_2.Bool_5', 'Channel_1.Device_2.Bool_6', 'Channel_1.Device_2.Bool_7', 'Channel_1.Device_2.Bool_8']
    tags = pytest.opcClient.read(taglistkep, group='Device1and2Group')
    assert(len(tags) == 6 and tags[0][2] == 'Good' and tags[1][2] == 'Good' and tags[2][2] == 'Good' and tags[3][2] == 'Good' and tags[4][2] == 'Good' and tags[5][2] == 'Good')

def test_device1groupread():
    tagkep = ['Channel_1.Device_1.Bool_1'] # needs to contain a tag
    tags = pytest.opcClient.read(tagkep, group='Device1Group')
    # result same as test_groupinitializeread
    assert(len(tags) == 4 and tags[0][2] == 'Good' and tags[1][2] == 'Good' and tags[2][2] == 'Good' and tags[3][2] == 'Good')

def test_device1groupreadcache():
    tagkep = ['Channel_1.Device_1.Bool_1'] # needs to contain a tag
    tags = pytest.opcClient.read(tagkep, group='Device1Group', source='cache')
    # result same as test_groupinitializeread
    assert(len(tags) == 4 and tags[0][2] == 'Good' and tags[1][2] == 'Good' and tags[2][2] == 'Good' and tags[3][2] == 'Good')

def test_device1groupreadcachesync():
    tagkep = ['Channel_1.Device_1.Bool_1'] # needs to contain a tag
    tags = pytest.opcClient.read(tagkep, group='Device1Group', source='cache', sync=True)
    # result same as test_groupinitializeread
    assert(len(tags) == 4 and tags[0][2] == 'Good' and tags[1][2] == 'Good' and tags[2][2] == 'Good' and tags[3][2] == 'Good')

def test_device1groupreadsync():
    tagkep = ['Channel_1.Device_1.Bool_1'] # needs to contain a tag
    tags = pytest.opcClient.read(tagkep, group='Device1Group', sync=True)
    # result same as test_groupinitializeread
    assert(len(tags) == 4 and tags[0][2] == 'Good' and tags[1][2] == 'Good' and tags[2][2] == 'Good' and tags[3][2] == 'Good')

def test_readrebuildgroup():
    taglistkep = ['Channel_1.Device_2.Bool_6', 'Channel_1.Device_2.Bool_7', 'Channel_1.Device_2.Bool_8', 'Channel_1.Device_2.Bool_9', 'Channel_1.Device_2.Bool_10']
    tags = pytest.opcClient.read(taglistkep, group='Device1and2Group', rebuild=True)
    assert(len(tags) == 5 and tags[0][2] == 'Good' and tags[1][2] == 'Good' and tags[2][2] == 'Good' and tags[3][2] == 'Good' and tags[4][2] == 'Good')

def test_readinvalidtags():
    def traceFunc(str):
        print(str)
    pytest.opcClient.set_trace(traceFunc)
    invalidtaglistkep = ['Channel_20.Device_1.Bool_1', 'Channel_20.Device_1.Tag_1', 'Channel_20.Device_1.Tag_2', 'Channel_20.Device_1.Tag_3']
    invalidresults = pytest.opcClient.read(invalidtaglistkep)
    assert(len(invalidresults) == 4 and invalidresults[0][1] == None and invalidresults[0][2] == 'Error')

def test_readnontags():
    with pytest.raises( TypeError ) as exc_info:
        taglistkep = [123.45, 67.89]
        response = pytest.opcClient.read(taglistkep)
    
    assert("read(): 'tags' parameter must be a string or a list of strings" in str(exc_info.value))

def test_readmixedstyle():
    with pytest.raises( TypeError ) as exc_info: # unsupported request
        taglistkep = ['Channel_20.Device_1.Bool_1', '@MemUsed']
        response = pytest.opcClient.read(taglistkep)
    
    assert("read(): system health and OPC tags cannot be included in the same group" in str(exc_info.value))

def test_readhealth():
    healthtags = ['@CpuUsage', '@MemFree', '@MemUsed', '@MemTotal', '@MemPercent', '@DiskFree', '@SineWave', '@SawWave']
    healthlist = pytest.opcClient.read(healthtags)
    assert(healthlist[0][0] == '@CpuUsage' and healthlist[1][0] == '@MemFree' and healthlist[2][0] == '@MemUsed' and  healthlist[3][0] == '@MemTotal' and
           healthlist[4][0] == '@MemPercent' and  healthlist[5][0] == '@DiskFree' and healthlist[6][0] == '@SineWave' and healthlist[7][0] == '@SawWave')

def test_iread():
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    tags = pytest.opcClient.iread(taglistkep)
    opclist = list(tags)
    assert(len(opclist) == 4 and opclist[0][2] == 'Good') # test for 4 tags and quality of first tag

def test_groupremove():
    removed = pytest.opcClient.remove(groups='Device1and2Group')
    assert(removed)

def test_write():
    tagPair = [('Channel_1.Device_1.Tag_1', 8), ('Channel_1.Device_1.Tag_2', 62), ('Channel_1.Device_1.Tag_3', 74)]
    status = pytest.opcClient.write(tagPair)
    assert(status[0][1] == 'Success' and  status[1][1] == 'Success' and status[2][1] == 'Success')
    
def test_writesized():
    tagPair = [('Channel_1.Device_1.Tag_1', 8), ('Channel_1.Device_1.Tag_2', 62)]
    status = pytest.opcClient.write(tagPair, size=2)
    assert(status[0][1] == 'Success' and  status[1][1] == 'Success')
    
def test_writewitherror():
    tagPair = [('Channel_1.Device_1.Tag_1', 8), ('Channel_1.Device_1.Tag_2', 62), ('Channel_1.Device_1.Tag_3', 74)]
    status = pytest.opcClient.write(tagPair, include_error=True)
    assert(status[0][1] == 'Success' and  status[1][1] == 'Success' and status[2][1] == 'Success')
    
def test_badwrite():
    with pytest.raises( TypeError ) as exc_info:
        tagTriple = (8, 'Channel_1.Device_1.Tag_1')
        status = pytest.opcClient.write(tagTriple)
    assert("write(): 'tag_value_pairs' parameter must be a (tag, value) tuple or a list of (tag,value) tuples" in str(exc_info.value))
    
def test_singlewrite():
    tagPair = ('Channel_1.Device_1.Tag_1', 8)
    status = pytest.opcClient.write(tagPair)
    assert(status == 'Success')
    
def test_iwrite():
    tagPair = [('Channel_1.Device_1.Tag_1', 8), ('Channel_1.Device_1.Tag_2', 62), ('Channel_1.Device_1.Tag_3', 74)]
    status = list(pytest.opcClient.iwrite(tagPair))
    assert(status[0][1] == 'Success' and  status[1][1] == 'Success' and status[2][1] == 'Success')

def test_getitem():
    tagkep = 'Channel_1.Device_1.Tag_1'
    value = pytest.opcClient.__getitem__(tagkep)
    assert(value > 0)

def test_setitem():
    tagkep = 'Channel_1.Device_1.Tag_1'
    tagvaluepair = pytest.opcClient.__setitem__(tagkep, 10)
    assert(tagvaluepair[0] == 'Channel_1.Device_1.Tag_1')
    
def test_properties():
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    props = pytest.opcClient.properties(taglistkep)
    assert(props[0][0] == 'Channel_1.Device_1.Bool_1')

def test_badtagproperties():
#    opcException = BaseException("test_badtagproperties")
    with pytest.raises( (TypeError, BaseException) ) as exc_info:
        taglistkep = ['Channel1.Device_1.Bool_1', 'Channel1.Device_1.Tag_1']
        props = pytest.opcClient.properties(taglistkep)
    assert("properties: -1073479672" in str(exc_info.value)) # error 0xC0040008: The item ID is not syntactically valid

def test_invalidtypeproperties():
    with pytest.raises( TypeError ) as exc_info:
        taglistkep = [1.1, 123]
        props = pytest.opcClient.properties(taglistkep)
    assert("properties(): 'tags' parameter must be a string or a list of strings" in str(exc_info.value)) # error 0xC0040008

def test_notaglistediproperties():
    props = pytest.opcClient.properties(tags=None)
    assert(len(props) == 0)

def test_singletagproperties():
    tagkep = 'Channel_1.Device_1.Bool_1'
    props = pytest.opcClient.properties(tagkep)
    assert(props[0][2] == 'Channel_1.Device_1.Bool_1')

def test_singletagpropertieswithid():
    tagkep = 'Channel_1.Device_1.Bool_1'
    prop = pytest.opcClient.properties(tagkep, id=1)
    assert(prop == 11)

def test_singletagpropertieswithids():
    tagkep = 'Channel_1.Device_1.Bool_1'
    props = pytest.opcClient.properties(tagkep, id=(1,2))
    assert(len(props) == 2 and props[0][0] == 1 and props[1][0] == 2)

def test_propertieswithid():
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    props = pytest.opcClient.properties(taglistkep, id=1)
    assert(props[0][0] == 'Channel_1.Device_1.Bool_1')

def test_propertieswithids():
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    props = pytest.opcClient.properties(taglistkep, id=(1, 2))
    assert(len(props) == 8 and props[0][0] == 'Channel_1.Device_1.Bool_1' and props[7][0] == 'Channel_1.Device_1.Tag_3')

def test_iproperties():
    taglistkep = ['Channel_1.Device_1.Bool_1', 'Channel_1.Device_1.Tag_1', 'Channel_1.Device_1.Tag_2', 'Channel_1.Device_1.Tag_3']
    props = list(pytest.opcClient.iproperties(taglistkep))
    assert(props[0][0] == 'Channel_1.Device_1.Bool_1')

def test_list():
    opclist = pytest.opcClient.list()
    assert(opclist[2] == 'Channel_1')

def test_listbadpath():
    opclist = pytest.opcClient.list(paths="Channel1")
    assert(len(opclist) == 0)

def test_listdevices():
    opclist = pytest.opcClient.list(paths=["Channel_1", "Channel_2", "Channel_3"])
    assert(len(opclist) == 5 and opclist[1] == 'Device_1' and opclist[4] == 'Device_4')

def test_listdevicewithrecursion():
    opclist = pytest.opcClient.list(paths="Channel_1.Device_1", recursive=True)
    assert(len(opclist) == 40 and opclist[0] == 'Channel_1.Device_1._System._DeviceId' and opclist[1] == 'Channel_1.Device_1._System._Enabled')

def test_listchannels():
    opclist = pytest.opcClient.list(paths="Channel*")
    assert(len(opclist) == 5 and opclist[1] == 'Channel_1' and opclist[4] == 'Channel_4')

def test_listdeviceswithtype():
    opclist = pytest.opcClient.list(paths=["Channel_1", "Channel_2", "Channel_3"], include_type=True)
    assert(len(opclist) == 5 and opclist[1][0] == 'Device_1' and opclist[1][1] == 'Branch' and opclist[4][0] == 'Device_4' and opclist[4][1] == 'Branch')

def test_listwithbadpath():
    with pytest.raises(TypeError) as exc_info:
        response = pytest.opcClient.list(paths=[1234])
    assert("list(): 'paths' parameter must be a string or a list of strings" in str(exc_info.value))

def test_flatlist():
    opclist = pytest.opcClient.list(flat=True)
    assert(len(opclist) == 483) # specific to KEPWareServerEx simdemo.opf

def test_ilist():
    opclist = list(pytest.opcClient.ilist())
    assert(opclist[2] == 'Channel_1')

def test_servers():
    serverlist = pytest.opcClient.servers('127.0.0.1') # Works with Matrikon.OPC.Automation dll
    assert(len(serverlist) > 0)

def test_info():
    infolist = pytest.opcClient.info()
    assert(len(infolist) > 0)

def test_ping():
    result = pytest.opcClient.ping()
    assert(result)

def test_health():
    healthtags = ['@MemFree', '@MemUsed', '@MemTotal', '@MemPercent', '@DiskFree', '@SineWave', '@SawWave']
    healthlist = pytest.opcClient._read_health(healthtags)
    assert(len(healthlist) == 7)

def test_taskhealth():
    healthtags = ['@TaskMem(KEPServerEx)', '@TaskCpu(KEPServerEx)', '@TaskExists(KEPServerEx)']
    #healthtags = ['@TaskMem(KEPServerEx)', '@TaskExists(KEPServerEx)']
    healthlist = pytest.opcClient._read_health(healthtags)
    assert(len(healthlist) == 3 and healthlist[0][2] == 'Good' and healthlist[1][2] == 'Good' and healthlist[2][2] == 'Good')

def test_updatetxtime():
    updated = pytest.opcClient._update_tx_time()
    assert(updated == False) # only available for gateway sessions
