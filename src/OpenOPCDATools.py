###########################################################################
#
# OpenOPC for Python OPC-DA Tools Library Module
#
# A Windows only OPC-DA library module.
#
# Copyright (c) 2007-2012 Barry Barnreiter (barry_b@users.sourceforge.net)
# Copyright (c) 2014 Anton D. Kachalov (mouse@yandex.ru)
# Copyright (c) 2017 José A. Maita (jose.a.maita@gmail.com)
# Copyright (c) 2022 j3mg
#
###########################################################################
import re
import SystemHealth
import pythoncom
import pywintypes
import time
from Common import exceptional, get_error_str, quality_str, type_check, wild2regex, ACCESS_RIGHTS, BROWSER_TYPE, OPCError, OPC_QUALITY, OPC_STATUS

class ClientTools():
    def __init__(self, opc_class, opc_host):
        self.opc_class = opc_class
        self.opc_host = opc_host
        self.__open_serv__ = None 
        self.__open_host__ = None
        self.__open_port__ = None
        self.__open_guid__ = None
        self._prev_serv_time = None

    def set_opc_host(self, host):
        self.opc_host = host

    def set_gateway_settings(self, service, host, port, guid):
        self.__open_serv__ = service 
        self.__open_host__ = host
        self.__open_port__ = port
        self.__open_guid__ = guid

    def iproperties(self, _opc, tags, id=None):
        """Iterable version of properties()"""

        try:
            self._update_tx_time()
            pythoncom.CoInitialize()

            tags, single_tag, valid = type_check(tags)
            if not valid:
                raise TypeError("properties(): 'tags' parameter must be a string or a list of strings")

            try:
                id.remove(0)
                include_name = True
            except:
                include_name = False

            if id != None:
                descriptions= []

                if isinstance(id, list) or isinstance(id, tuple):
                    property_id = list(id)
                    single_property = False
                else:
                    property_id = [id]
                    single_property = True

                for i in property_id:
                    descriptions.append('Property id %d' % i)
            else:
                single_property = False

            properties = []

            for tag in tags:

                if id == None:
                    descriptions = []
                    property_id = []
                    count, property_id, descriptions, datatypes = _opc.QueryAvailableProperties(tag)

                    # Remove bogus negative property id (not sure why this sometimes happens)
                    tag_properties = list(map(lambda x, y: (x, y), property_id, descriptions))
                    property_id = [p for p, d in tag_properties if p > 0]
                    descriptions = [d for p, d in tag_properties if p > 0]

                property_id.insert(0, 0)
                values = []
                errors = []
                values, errors = _opc.GetItemProperties(tag, len(property_id)-1, property_id)

                property_id.pop(0)
                values = [str(v) if type(v) == pywintypes.TimeType else v for v in values]

                # Replace variant id with type strings
                try:
                    i = property_id.index(1)
                    values[i] = vt[values[i]]
                except:
                    pass

                # Replace quality bits with quality strings
                try:
                    i = property_id.index(3)
                    values[i] = quality_str(values[i])
                except:
                    pass

                # Replace access rights bits with strings
                try:
                    i = property_id.index(5)
                    values[i] = ACCESS_RIGHTS[values[i]]
                except:
                    pass

                if id != None:
                    if single_property:
                        if single_tag:
                            tag_properties = values
                        else:
                            tag_properties = [values]
                    else:
                        tag_properties = list(map(lambda x, y: (x, y), property_id, values))
                else:
                    tag_properties = list(map(lambda x, y, z: (x, y, z), property_id, descriptions, values))
                    tag_properties.insert(0, (0, 'Item ID (virtual property)', tag))

                if include_name:    tag_properties.insert(0, (0, tag))
                if not single_tag:  tag_properties = [tuple([tag] + list(p)) for p in tag_properties]

                for p in tag_properties: yield p

        except pythoncom.com_error as err:
            error_msg = 'properties: %s' % get_error_str(err, _opc)
            raise OPCError(error_msg)

    def properties(self, _opc, tags, id=None):
        """Return list of property tuples (id, name, value) for the specified tag(s) """

        if type(tags) not in (list, tuple) and type(id) not in (type(None), list, tuple):
            single = True
        else:
            single = False

        props = self.iproperties(_opc, tags, id)

        if single:
            return list(props)[0]
        else:
            return list(props)

    def ilist(self, _opc, paths='*', recursive=False, flat=False, include_type=False):
        """Iterable version of list()"""

        try:
            self._update_tx_time()
            pythoncom.CoInitialize()

            try:
                browser = _opc.CreateBrowser()
            # For OPC servers that don't support browsing
            except:
                return {}

            paths, single, valid = type_check(paths)
            if not valid:
                raise TypeError("list(): 'paths' parameter must be a string or a list of strings")

            if len(paths) == 0: paths = ['*']
            nodes = {}

            for path in paths:

                if flat:
                    browser.MoveToRoot()
                    browser.Filter = ''
                    browser.ShowLeafs(True)

                    pattern = re.compile('^%s$' % wild2regex(path) , re.IGNORECASE)
                    matches = filter(pattern.search, browser)
                    if include_type:  matches = [(x, node_type) for x in matches]

                    for node in matches: yield node
                    continue

                queue = []
                queue.append(path)

                while len(queue) > 0:
                    tag = queue.pop(0)

                    browser.MoveToRoot()
                    browser.Filter = ''
                    pattern = None

                    path_str = '/'
                    path_list = tag.replace('.','/').split('/')
                    path_list = [p for p in path_list if len(p) > 0]
                    found_filter = False
                    path_postfix = '/'

                    for i, p in enumerate(path_list):
                        if found_filter:
                            path_postfix += p + '/'
                        elif p.find('*') >= 0:
                            pattern = re.compile('^%s$' % wild2regex(p) , re.IGNORECASE)
                            found_filter = True
                        elif len(p) != 0:
                            pattern = re.compile('^.*$')
                            browser.ShowBranches()

                            # Branch node, so move down
                            if len(browser) > 0:
                                try:
                                    browser.MoveDown(p)
                                    path_str += p + '/'
                                except:
                                    if i < len(path_list)-1: return
                                    pattern = re.compile('^%s$' % wild2regex(p) , re.IGNORECASE)

                            # Leaf node, so append all remaining path parts together
                            # to form a single search expression
                            else:
                                p = string.join(path_list[i:], '.')
                                pattern = re.compile('^%s$' % wild2regex(p) , re.IGNORECASE)
                                break

                    browser.ShowBranches()

                    if len(browser) == 0:
                        browser.ShowLeafs(False)
                        lowest_level = True
                        node_type = 'Leaf'
                    else:
                        lowest_level = False
                        node_type = 'Branch'

                    matches = filter(pattern.search, browser)

                    if not lowest_level and recursive:
                        queue += [path_str + x + path_postfix for x in matches]
                    else:
                        if lowest_level:  matches = [exceptional(browser.GetItemID,x)(x) for x in matches]
                        if include_type:  matches = [(x, node_type) for x in matches]
                        for node in matches:
                            if not node in nodes: yield node
                            nodes[node] = True

        except pythoncom.com_error as err:
            error_msg = 'list: %s' % get_error_str(err, _opc)
            raise OPCError(error_msg)

    def list(self, _opc, paths='*', recursive=False, flat=False, include_type=False):
        """Return list of item nodes at specified path(s) (tree browser)"""

        nodes = self.ilist(_opc, paths, recursive, flat, include_type)
        return list(nodes)

    def servers(self, _opc, opc_host='localhost'):
        """Return list of available OPC servers"""

        try:
            pythoncom.CoInitialize()
            servers = _opc.GetOPCServers(opc_host)
            servers = [s for s in servers if s != None]
            return servers

        except pythoncom.com_error as err:
            error_msg = 'servers: %s' % get_error_str(err, _opc)
            raise OPCError(error_msg)

    def info(self, _opc):
        """Return list of (name, value) pairs about the OPC server"""

        try:
            self._update_tx_time()
            pythoncom.CoInitialize()

            info_list = []

            if self.__open_serv__:
                mode = 'OpenOPC'
            else:
                mode = 'DCOM'

            info_list += [('Protocol', mode)]

            if mode == 'OpenOPC':
                info_list += [('Gateway Host', '%s:%s' % (self.__open_host__, self.__open_port__))]
                info_list += [('Gateway Version', '%s' % __version__)]
            info_list += [('Class', self.opc_class)]
            info_list += [('Client Name', _opc.ClientName)]
            info_list += [('OPC Host', self.opc_host)]
            info_list += [('OPC Server', _opc.ServerName)]
            info_list += [('State', OPC_STATUS[_opc.ServerState])]
            info_list += [('Version', '%d.%d (Build %d)' % (_opc.MajorVersion, _opc.MinorVersion, _opc.BuildNumber))]

            try:
                browser = _opc.CreateBrowser()
                browser_type = BROWSER_TYPE[browser.Organization]
            except:
                browser_type = 'Not Supported'

            info_list += [('Browser', browser_type)]
            info_list += [('Start Time', str(_opc.StartTime))]
            info_list += [('Current Time', str(_opc.CurrentTime))]
            info_list += [('Vendor', _opc.VendorInfo)]

            return info_list

        except pythoncom.com_error as err:
            error_msg = 'info: %s' % get_error_str(err, _opc)
            raise OPCError(error_msg)

    def ping(self, _opc):
        """Check if we are still talking to the OPC server"""
        try:
            # Convert OPC server time to milliseconds
            timestr = str(_opc.CurrentTime)
            hours = timestr[11:12]
            minutes = timestr[14:15]
            seconds = timestr[17:]
            totalseconds = float(hours)*3600.0 + float(minutes)*60.0 + float(seconds[:9])
            opc_serv_time = int(totalseconds) * 1000000.0
            if opc_serv_time == self._prev_serv_time:
                return False
            else:
                self._prev_serv_time = opc_serv_time
                return True
        except pythoncom.com_error:
            return False

    def _update_tx_time(self):
        """Update the session's last transaction time in the Gateway Service"""
        if self.__open_serv__:
            self.__open_serv__._tx_times[self.__open_guid__] = time.time()
            return True
        return False;
