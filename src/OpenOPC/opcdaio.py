###########################################################################
#
# OpenOPC for Python OPC-DA IO Library file
#
# A Windows only OPC-DA library file.
#
# Copyright (c) 2007-2012 Barry Barnreiter (barry_b@users.sourceforge.net)
# Copyright (c) 2014 Anton D. Kachalov (mouse@yandex.ru)
# Copyright (c) 2017 José A. Maita (jose.a.maita@gmail.com)
# Copyright (c) 2022 j3mg
#
###########################################################################
import re
import OpenOPC.systemhealth
import win32com.client
import win32com.server.util
import win32event
import pythoncom
import pywintypes
import time
from multiprocessing import Queue
from OpenOPC.common import get_error_str, quality_str, type_check, tags2trace, TimeoutError, OPCError, SOURCE_CACHE, SOURCE_DEVICE, OPC_QUALITY

current_client = None

class GroupEvents:
    def __init__(self):
        self.client = current_client

    def OnDataChange(self, TransactionID, NumItems, ClientHandles, ItemValues, Qualities, TimeStamps):
        self.client.callback_queue.put((TransactionID, ClientHandles, ItemValues, Qualities, TimeStamps))

class ClientIO():
    def __init__(self):
        self.trace = None

        self.callback_queue = Queue()

        self._tx_id = 0
        # On reconnect we need to remove the old group names from OpenOPC's internal
        # cache since they are now invalid
        self._groups = {}
        self._group_tags = {}
        self._group_valid_tags = {}
        self._group_server_handles = {}
        self._group_handles_tag = {}
        self._group_hooks = {}
        self.cpu = None

    def setTrace(self, trace):
        self.trace = trace

    def iread(self, _opc, clientTools, tags=None, group=None, size=None, pause=0, source='hybrid', update=-1, timeout=5000, sync=False, include_error=False, rebuild=False):
        """Iterable version of read()"""

        def add_items(tags):
            names = list(tags)

            names.insert(0,0)
            errors = []

            if self.trace: self.trace('Validate(%s)' % tags2trace(names))

            try:
                errors = opc_items.Validate(len(names)-1, names)
            except:
                pass

            valid_tags = []
            valid_values = []
            client_handles = []

            if not sub_group in self._group_handles_tag:
                self._group_handles_tag[sub_group] = {}
                n = 0
            elif len(self._group_handles_tag[sub_group]) > 0:
                n = max(self._group_handles_tag[sub_group]) + 1
            else:
                n = 0

            for i, tag in enumerate(tags):
                if errors[i] == 0:
                    valid_tags.append(tag)
                    client_handles.append(n)
                    self._group_handles_tag[sub_group][n] = tag
                    n += 1
                elif include_error:
                    error_msgs[tag] = _opc.GetErrorString(errors[i])

                if self.trace and errors[i] != 0: self.trace('%s failed validation' % tag)

            client_handles.insert(0,0)
            valid_tags.insert(0,0)
            server_handles = []
            errors = []

            if self.trace: self.trace('AddItems(%s)' % tags2trace(valid_tags))

            try:
                server_handles, errors = opc_items.AddItems(len(client_handles)-1, valid_tags, client_handles)
            except:
                pass

            valid_tags_tmp = []
            server_handles_tmp = []
            valid_tags.pop(0)

            if not sub_group in self._group_server_handles:
                self._group_server_handles[sub_group] = {}

            for i, tag in enumerate(valid_tags):
                if errors[i] == 0:
                    valid_tags_tmp.append(tag)
                    server_handles_tmp.append(server_handles[i])
                    self._group_server_handles[sub_group][tag] = server_handles[i]
                elif include_error:
                    error_msgs[tag] = _opc.GetErrorString(errors[i])

            valid_tags = valid_tags_tmp
            server_handles = server_handles_tmp

            return valid_tags, server_handles

        def remove_items(tags):
            if self.trace: self.trace('RemoveItems(%s)' % tags2trace(['']+tags))
            server_handles = [self._group_server_handles[sub_group][tag] for tag in tags]
            server_handles.insert(0,0)
            errors = []

            try:
                errors = opc_items.Remove(len(server_handles)-1, server_handles)
            except pythoncom.com_error as err:
                error_msg = 'RemoveItems: %s' % get_error_str(err, _opc)
                raise OPCError(error_msg)

        try:
            clientTools._update_tx_time()
            pythoncom.CoInitialize()

            if include_error:
                sync = True

            if sync:
                update = -1

            tags, single, valid = type_check(tags)
            if not valid:
                raise TypeError("iread(): 'tags' parameter must be a string or a list of strings")

            # Group exists
            if group in self._groups and not rebuild:
                num_groups = self._groups[group]
                data_source = SOURCE_CACHE

            # Group non-existant
            else:
                if size:
                    # Break-up tags into groups of 'size' tags
                    tag_groups = [tags[i:i+size] for i in range(0, len(tags), size)]
                else:
                    tag_groups = [tags]

                num_groups = len(tag_groups)
                data_source = SOURCE_DEVICE

            results = []

            for gid in range(num_groups):
                if gid > 0 and pause > 0: time.sleep(pause/1000.0)

                error_msgs = {}
                opc_groups = _opc.OPCGroups
                opc_groups.DefaultGroupUpdateRate = update

                # Anonymous group
                if group == None:
                    try:
                        if self.trace: self.trace('AddGroup()')
                        opc_group = opc_groups.Add()
                    except pythoncom.com_error as err:
                        error_msg = 'AddGroup: %s' % get_error_str(err, _opc)
                        raise OPCError(error_msg)
                    sub_group = group
                    new_group = True
                else:
                    sub_group = '%s.%d' % (group, gid)

                    # Existing named group
                    try:
                        if self.trace: self.trace('GetOPCGroup(%s)' % sub_group)
                        opc_group = opc_groups.GetOPCGroup(sub_group)
                        new_group = False

                    # New named group
                    except:
                        try:
                            if self.trace: self.trace('AddGroup(%s)' % sub_group)
                            opc_group = opc_groups.Add(sub_group)
                        except pythoncom.com_error as err:
                            error_msg = 'AddGroup: %s' % get_error_str(err, _opc)
                            raise OPCError(error_msg)
                        self._groups[str(group)] = len(tag_groups)
                        new_group = True

                opc_items = opc_group.OPCItems

                if new_group:
                    opc_group.IsSubscribed = 1
                    opc_group.IsActive = 1
                    if not sync:
                        if self.trace: self.trace('WithEvents(%s)' % opc_group.Name)
                        global current_client
                        current_client = self
                        self._group_hooks[opc_group.Name] = win32com.client.WithEvents(opc_group, GroupEvents)

                    tags = tag_groups[gid]

                    valid_tags, server_handles = add_items(tags)

                    self._group_tags[sub_group] = tags
                    self._group_valid_tags[sub_group] = valid_tags

                # Rebuild existing group
                elif rebuild:
                    tags = tag_groups[gid]

                    valid_tags = self._group_valid_tags[sub_group]
                    add_tags = [t for t in tags if t not in valid_tags]
                    del_tags = [t for t in valid_tags if t not in tags]

                    if len(add_tags) > 0:
                        valid_tags, server_handles = add_items(add_tags)
                        valid_tags = self._group_valid_tags[sub_group] + valid_tags

                    if len(del_tags) > 0:
                        remove_items(del_tags)
                        valid_tags = [t for t in valid_tags if t not in del_tags]

                    self._group_tags[sub_group] = tags
                    self._group_valid_tags[sub_group] = valid_tags

                    if source == 'hybrid': data_source = SOURCE_DEVICE

                # Existing group
                else:
                    tags = self._group_tags[sub_group]
                    valid_tags = self._group_valid_tags[sub_group]
                    if sync:
                        server_handles = [item.ServerHandle for item in opc_items]

                tag_value = {}
                tag_quality = {}
                tag_time = {}
                tag_error = {}

                # Sync Read
                if sync:
                    values = []
                    errors = []
                    qualities = []
                    timestamps= []

                    if len(valid_tags) > 0:
                        server_handles.insert(0,0)

                        if source != 'hybrid':
                            data_source = SOURCE_CACHE if source == 'cache' else SOURCE_DEVICE

                        if self.trace: self.trace('SyncRead(%s)' % data_source)

                        try:
                            values, errors, qualities, timestamps = opc_group.SyncRead(data_source, len(server_handles)-1, server_handles)
                        except pythoncom.com_error as err:
                            error_msg = 'SyncRead: %s' % get_error_str(err, _opc)
                            raise OPCError(error_msg)

                        for i,tag in enumerate(valid_tags):
                            tag_value[tag] = values[i]
                            tag_quality[tag] = qualities[i]
                            tag_time[tag] = timestamps[i]
                            tag_error[tag] = errors[i]

                # Async Read
                else:
                    if len(valid_tags) > 0:
                        if self._tx_id >= 0xFFFF:
                            self._tx_id = 0
                        self._tx_id += 1

                        if source != 'hybrid':
                            data_source = SOURCE_CACHE if source == 'cache' else SOURCE_DEVICE

                        if self.trace: self.trace('AsyncRefresh(%s)' % data_source)

                        try:
                            opc_group.AsyncRefresh(data_source, self._tx_id)
                        except pythoncom.com_error as err:
                            error_msg = 'AsyncRefresh: %s' % get_error_str(err, _opc)
                            raise OPCError(error_msg)

                        tx_id = 0
                        start = time.time() * 1000

                        while tx_id != self._tx_id:
                            now = time.time() * 1000
                            if now - start > timeout:
                                raise TimeoutError('Callback: Timeout waiting for data')

                            if self.callback_queue.empty():
                                pythoncom.PumpWaitingMessages()
                            else:
                                tx_id, handles, values, qualities, timestamps = self.callback_queue.get()

                        for i,h in enumerate(handles):
                            tag = self._group_handles_tag[sub_group][h]
                            tag_value[tag] = values[i]
                            tag_quality[tag] = qualities[i]
                            tag_time[tag] = timestamps[i]

                for tag in tags:
                    if tag in tag_value:
                        if (not sync and len(valid_tags) > 0) or (sync and tag_error[tag] == 0):
                            value = tag_value[tag]
                            if type(value) == pywintypes.TimeType:
                                value = str(value)
                            quality = quality_str(tag_quality[tag])
                            timestamp = str(tag_time[tag])
                        else:
                            value = None
                            quality = 'Error'
                            timestamp = None
                        if include_error:
                            error_msgs[tag] = _opc.GetErrorString(tag_error[tag]).strip('\r\n')
                    else:
                        value = None
                        quality = 'Error'
                        timestamp = None
                        if include_error and not tag in error_msgs:
                            error_msgs[tag] = ''

                    if single:
                        if include_error:
                            yield (value, quality, timestamp, error_msgs[tag])
                        else:
                            yield (value, quality, timestamp)
                    else:
                        if include_error:
                            yield (tag, value, quality, timestamp, error_msgs[tag])
                        else:
                            yield (tag, value, quality, timestamp)

                if group == None:
                    try:
                        if not sync and opc_group.Name in self._group_hooks:
                            if self.trace: self.trace('CloseEvents(%s)' % opc_group.Name)
                            self._group_hooks[opc_group.Name].close()

                        if self.trace: self.trace('RemoveGroup(%s)' % opc_group.Name)
                        opc_groups.Remove(opc_group.Name)

                    except pythoncom.com_error as err:
                        error_msg = 'RemoveGroup: %s' % get_error_str(err, _opc)
                        raise OPCError(error_msg)

        except pythoncom.com_error as err:
            error_msg = 'read: %s' % get_error_str(err, _opc)
            raise OPCError(error_msg)

    def read(self, _opc, clientTools, tags=None, group=None, size=None, pause=0, source='hybrid', update=-1, timeout=5000, sync=False, include_error=False, rebuild=False):
        """Return list of (value, quality, time) tuples for the specified tag(s)"""

        tags_list, single, valid = type_check(tags)
        if not valid:
            raise TypeError("read(): 'tags' parameter must be a string or a list of strings")

        num_health_tags = len([t for t in tags_list if t[:1] == '@'])
        num_opc_tags = len([t for t in tags_list if t[:1] != '@'])

        if num_health_tags > 0:
            if num_opc_tags > 0:
                raise TypeError("read(): system health and OPC tags cannot be included in the same group")
            results = self._read_health(clientTools, tags)
        else:
            results = self.iread(_opc, clientTools, tags, group, size, pause, source, update, timeout, sync, include_error, rebuild)

        if single:
            return list(results)[0]
        else:
            return list(results)

    def _read_health(self, clientTools, tags):
        """Return values of special system health monitoring tags"""

        clientTools._update_tx_time()
        tags, single, valid = type_check(tags)

        time_str = time.strftime('%x %H:%M:%S')
        results = []

        for t in tags:
            if   t == '@MemFree':      value = OpenOPC.systemhealth.mem_free()
            elif t == '@MemUsed':      value = OpenOPC.systemhealth.mem_used()
            elif t == '@MemTotal':     value = OpenOPC.systemhealth.mem_total()
            elif t == '@MemPercent':   value = OpenOPC.systemhealth.mem_percent()
            elif t == '@DiskFree':     value = OpenOPC.systemhealth.disk_free()
            elif t == '@SineWave':     value = OpenOPC.systemhealth.sine_wave()
            elif t == '@SawWave':      value = OpenOPC.systemhealth.saw_wave()

            elif t == '@CpuUsage':
                if self.cpu == None:
                    self.cpu = OpenOPC.systemhealth.CPU()
                    time.sleep(0.1)
                value = self.cpu.get_usage()

            else:
                value = None

                m = re.match('@TaskMem(((.*?)))', t)
                if m:
                    image_name = m.group(1)
                    value = OpenOPC.systemhealth.task_mem(image_name)

                m = re.match('@TaskCpu(((.*?)))', t)
                if m:
                    image_name = m.group(1)
                    value = OpenOPC.systemhealth.task_cpu(image_name)

                m = re.match('@TaskExists(((.*?)))', t)
                if m:
                    image_name = m.group(1)
                    value = OpenOPC.systemhealth.task_exists(image_name)

            if value == None:
                quality = 'Error'
            else:
                quality = 'Good'

            if single:
                results.append((value, quality, time_str))
            else:
                results.append((t, value, quality, time_str))

        return results

    def iwrite(self, _opc, clientTools, tag_value_pairs, size=None, pause=0, include_error=False):
        """Iterable version of write()"""

        try:
            clientTools._update_tx_time()
            pythoncom.CoInitialize()

            def _valid_pair(p):
                if type(p) in (list, tuple) and len(p) >= 2 and type(p[0]) in (str,bytes):
                    return True
                else:
                    return False

            if type(tag_value_pairs) not in (list, tuple):
                raise TypeError("write(): 'tag_value_pairs' parameter must be a (tag, value) tuple or a list of (tag,value) tuples")

            if tag_value_pairs == None:
                tag_value_pairs = ['']
                single = False
            elif type(tag_value_pairs[0]) in (str,bytes):
                tag_value_pairs = [tag_value_pairs]
                single = True
            else:
                single = False

            invalid_pairs = [p for p in tag_value_pairs if not _valid_pair(p)]
            if len(invalid_pairs) > 0:
                raise TypeError("write(): 'tag_value_pairs' parameter must be a (tag, value) tuple or a list of (tag,value) tuples")

            names = [tag[0] for tag in tag_value_pairs]
            tags = [tag[0] for tag in tag_value_pairs]
            values = [tag[1] for tag in tag_value_pairs]

            # Break-up tags & values into groups of 'size' tags
            if size:
                name_groups = [names[i:i+size] for i in range(0, len(names), size)]
                tag_groups = [tags[i:i+size] for i in range(0, len(tags), size)]
                value_groups = [values[i:i+size] for i in range(0, len(values), size)]
            else:
                name_groups = [names]
                tag_groups = [tags]
                value_groups = [values]

            num_groups = len(tag_groups)

            status = []

            for gid in range(num_groups):
                if gid > 0 and pause > 0: time.sleep(pause/1000.0)

                opc_groups = _opc.OPCGroups
                opc_group = opc_groups.Add()
                opc_items = opc_group.OPCItems

                names = name_groups[gid]
                tags = tag_groups[gid]
                values = value_groups[gid]

                names.insert(0,0)
                errors = []

                try:
                    errors = opc_items.Validate(len(names)-1, names)
                except:
                    pass

                n = 1
                valid_tags = []
                valid_values = []
                client_handles = []
                error_msgs = {}

                for i, tag in enumerate(tags):
                    if errors[i] == 0:
                        valid_tags.append(tag)
                        valid_values.append(values[i])
                        client_handles.append(n)
                        error_msgs[tag] = ''
                        n += 1
                    elif include_error:
                        error_msgs[tag] = _opc.GetErrorString(errors[i])

                client_handles.insert(0,0)
                valid_tags.insert(0,0)
                server_handles = []
                errors = []

                try:
                    server_handles, errors = opc_items.AddItems(len(client_handles)-1, valid_tags, client_handles)
                except:
                    pass

                valid_tags_tmp = []
                valid_values_tmp = []
                server_handles_tmp = []
                valid_tags.pop(0)

                for i, tag in enumerate(valid_tags):
                    if errors[i] == 0:
                        valid_tags_tmp.append(tag)
                        valid_values_tmp.append(valid_values[i])
                        server_handles_tmp.append(server_handles[i])
                        error_msgs[tag] = ''
                    elif include_error:
                        error_msgs[tag] = _opc.GetErrorString(errors[i])

                valid_tags = valid_tags_tmp
                valid_values = valid_values_tmp
                server_handles = server_handles_tmp

                server_handles.insert(0,0)
                valid_values.insert(0,0)
                errors = []

                if len(valid_values) > 1:
                    try:
                        errors = opc_group.SyncWrite(len(server_handles)-1, server_handles, valid_values)
                    except:
                        pass

                n = 0
                for tag in tags:
                    if tag in valid_tags:
                        if errors[n] == 0:
                            status = 'Success'
                        else:
                            status = 'Error'
                        if include_error:  error_msgs[tag] = _opc.GetErrorString(errors[n])
                        n += 1
                    else:
                        status = 'Error'

                    # OPC servers often include newline and carriage return characters
                    # in their error message strings, so remove any found.
                    if include_error:  error_msgs[tag] = error_msgs[tag].strip('\r\n')

                    if single:
                        if include_error:
                            yield (status, error_msgs[tag])
                        else:
                            yield status
                    else:
                        if include_error:
                            yield (tag, status, error_msgs[tag])
                        else:
                            yield (tag, status)

                opc_groups.Remove(opc_group.Name)

        except pythoncom.com_error as err:
            error_msg = 'write: %s' % get_error_str(err, _opc)
            raise OPCError(error_msg)

    def write(self, _opc, clientTools, tag_value_pairs, size=None, pause=0, include_error=False):
        """Write list of (tag, value) pair(s) to the server"""

        if type(tag_value_pairs) in (list, tuple) and type(tag_value_pairs[0]) in (list, tuple):
            single = False
        else:
            single = True

        status = self.iwrite(_opc, clientTools, tag_value_pairs, size, pause, include_error)

        if single:
            return list(status)[0]
        else:
            return list(status)

    def remove(self, _opc, groups):
        """Remove the specified tag group(s)"""

        try:
            pythoncom.CoInitialize()
            opc_groups = _opc.OPCGroups

            if type(groups) in (str,bytes):
                groups = [groups]
                single = True
            else:
                single = False

            status = []
            groups_deleted = False

            for group in groups:
                if group in self._groups:
                    for i in range(self._groups[group]):
                        sub_group = '%s.%d' % (group, i)

                        if sub_group in self._group_hooks:
                            if self.trace: self.trace('CloseEvents(%s)' % sub_group)
                            self._group_hooks[sub_group].close()

                        try:
                            if self.trace: self.trace('RemoveGroup(%s)' % sub_group)
                            errors = opc_groups.Remove(sub_group)
                            if errors == None:
                                groups_deleted = True
                        except pythoncom.com_error as err:
                            error_msg = 'RemoveGroup: %s' % get_error_str(err, _opc)
                            raise OPCError(error_msg)

                        del(self._group_tags[sub_group])
                        del(self._group_valid_tags[sub_group])
                        del(self._group_handles_tag[sub_group])
                        del(self._group_server_handles[sub_group])
                    del(self._groups[group])
            return groups_deleted

        except pythoncom.com_error as err:
            error_msg = 'remove: %s' % get_error_str(err, _opc)
            raise OPCError(error_msg)

    def __getitem__(self, _opc, clientTools, key):
        """Read single item (tag as dictionary key)"""
        value, quality, time_str = self.read(_opc, clientTools, key)
        return value

    def __setitem__(self, _opc, clientTools, key, value):
        """Write single item (tag as dictionary key)"""
        self.write(_opc, clientTools, (key, value))
        return (key, value)
