###########################################################################
#
# OpenOPC for Python OPC-DA Library File
#
# General functions file.
#
# Copyright (c) 2007-2012 Barry Barnreiter (barry_b@users.sourceforge.net)
# Copyright (c) 2014 Anton D. Kachalov (mouse@yandex.ru)
# Copyright (c) 2017 José A. Maita (jose.a.maita@gmail.com)
# Copyright (c) 2022 j3mg
#
###########################################################################
import os

try:
    import pythoncom # Only used by get_error_str on Python 32-bit systems
except:
    pass


# OPC Constants
ACCESS_RIGHTS = (0, 'Read', 'Write', 'Read/Write')
BROWSER_TYPE = (0, 'Hierarchical', 'Flat')
OPC_CLASS = 'OPC DA Automation Wrapper 2.02;Matrikon.OPC.Automation;Graybox.OPC.DAWrapper;HSCOPC.Automation;RSI.OPCAutomation;OPC.Automation'
OPC_SERVER = 'Hci.TPNServer;HwHsc.OPCServer;opc.deltav.1;AIM.OPC.1;Yokogawa.ExaopcDAEXQ.1;OSI.DA.1;OPC.PHDServerDA.1;Aspen.Infoplus21_DA.1;National Instruments.OPCLabVIEW;RSLinx OPC Server;KEPware.KEPServerEx.V4;Matrikon.OPC.Simulation;Prosys.OPC.Simulation;CCOPC.XMLWrapper.1;OPC.SimaticHMI.CoRtHmiRTm.1'
OPC_CLIENT = 'OpenOPC'
OPC_QUALITY = ('Bad', 'Uncertain', 'Unknown', 'Good')
OPC_STATUS = (0, 'Running', 'Failed', 'NoConfig', 'Suspended', 'Test')
SOURCE_CACHE = 1
SOURCE_DEVICE = 2

def exceptional(func, alt_return=None, alt_exceptions=(Exception,), final=None, catch=None):
   """Turns exceptions into an alternative return value"""

   def _exceptional(*args, **kwargs):
      try:
         try:
            return func(*args, **kwargs)
         except alt_exceptions:
            return alt_return
         except:
            if catch: return catch(sys.exc_info(), lambda:func(*args, **kwargs))
            raise
      finally:
         if final: final()
   return _exceptional

def get_error_str(error, _opc=None):
    """Return the error string for a OPC or COM error code"""

    hr, msg, exc, arg = error.args

    if exc == None:
        error_str = str(msg)
    else:
        scode = exc[5]

        if _opc is not None:
            try:
                opc_err_str = unicode(_opc.GetErrorString(scode)).strip('\r\n')
            except:
                opc_err_str = None
        else:
            opc_err_str = None

        try:
            com_err_str = unicode(pythoncom.GetScodeString(scode)).strip('\r\n')
        except:
            com_err_str = None

        # OPC error codes and COM error codes are overlapping concepts,
        # so we combine them together into a single error message.

        if opc_err_str == None and com_err_str == None:
            error_str = str(scode)
        elif opc_err_str == com_err_str:
            error_str = opc_err_str
        elif opc_err_str == None:
            error_str = com_err_str
        elif com_err_str == None:
            error_str = opc_err_str
        else:
            error_str = '%s (%s)' % (opc_err_str, com_err_str)

    return error_str

def quality_str(quality_bits):
    """Convert OPC quality bits to a descriptive string"""

    quality = (quality_bits >> 6) & 3
    return OPC_QUALITY[quality]

def type_check(tags):
    """Perform a type check on a list of tags"""

    if type(tags) in (list, tuple):
        single = False
    elif tags == None:
        tags = []
        single = False
    else:
        tags = [tags]
        single = True

    if len([t for t in tags if type(t) not in (str,bytes)]) == 0:
        valid = True
    else:
        valid = False

    return tags, single, valid

def tags2trace(tags):
    """Convert a list tags into a formatted string suitable for the trace callback log"""
    arg_str = ''
    for i,t in enumerate(tags[1:]):
        if i > 0: arg_str += ','
        arg_str += '%s' % t
    return arg_str

def wild2regex(string):
   """Convert a Unix wildcard glob into a regular expression"""
   return string.replace('.','[.]').replace('*','.*').replace('?','.').replace('!','^')

class TimeoutError(Exception):
    def __init__(self, txt):
        Exception.__init__(self, txt)

def OPCError(msg):
    return Exception("OPCError", msg) # Pyro5 cannot serialize custom exception types
