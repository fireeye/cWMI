########################################################################
# Copyright 2018 FireEye Inc.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
########################################################################

import os
import re
import keyword
import argparse


# only converts interface classes that are children of IUnknown
def convert_interface(infile, outfile, interface_name, overwrite):
    output_data = ''
    interface = {}

    with open(infile) as input_file:
        search_line = '{:s} : public IUnknown'.format(interface_name)
        interface_started = False
        method_started = False
        method_idx = 3
        method = {}
        for line in input_file.readlines():
            # look for our interface
            line_stripped = line.strip()
            if line_stripped and line_stripped.endswith(search_line):
                if not interface_started:
                    if overwrite or not os.path.exists(outfile):
                        output_data += 'import ctypes\n'
                        output_data += 'from ctypes import wintypes\n\n'
                        output_data += 'from .winerr import *\n'
                        output_data += 'from .wintype import *\n'
                        output_data += 'from . import winapi\n'
                        output_data += 'from . import com\n\n\n'
                        output_data += '_In_ = 1\n'
                        output_data += '_Out_ = 2\n\n'
                    else:
                        output_data += '\n\n'
                    interface_started = True
            elif line_stripped and interface_started:

                if line_stripped.endswith('};'):
                    # were done parsing interface
                    break

                # gather interface data
                if line_stripped.startswith('virtual') and not method_started:
                    # get name and return type
                    method['return_type'] = line_stripped.split(' ')[1]
                    method['name'] = line_stripped.split(' ')[3].replace('(', '')
                    method_started = True

                    # special case for void methods
                    if line_stripped.endswith(') = 0;'):
                        method['params'] = {}
                        interface[method_idx] = method
                        method = {}
                        method_idx += 1
                        method_started = False
                elif method_started:
                    param = {}
                    # processing parameters, out takes precedence
                    if '[out]' in line_stripped:
                        param['direction'] = '_Out_'
                    else:
                        param['direction'] = '_In_'

                    name_idx = -1
                    type_idx = name_idx - 1

                    if line_stripped.endswith(','):
                        char = ','
                    else:
                        name_idx -= 2
                        type_idx = name_idx - 1
                        char = ')'

                    if '* *' in line_stripped:
                        type_idx -= 1

                    param['name'] = line_stripped.split(' ')[name_idx].replace(char, '')
                    param['type'] = line_stripped.split(' ')[type_idx].replace(char, '')

                    if '* *' in line_stripped:
                        param['name'] = param['name'].replace('*', '**')

                    param['converted_name'] = convert_param(method, param['name'])
                    param['call_name'] = param['converted_name']

                    pointer_count = len(param['name'].split('*')) - 1
                    param['pointer_count'] = pointer_count

                    if pointer_count:
                        param['type'] += '{:s}'.format('*' * pointer_count)
                        param['name'] = param['name'].replace('*', '')

                    if 'params' in method:
                        method['params'].append(param)
                    else:
                        method['params'] = [param]

                    if line_stripped.endswith(') = 0;'):
                        # we are now done parsing this method
                        interface[method_idx] = method
                        method = {}
                        method_idx += 1
                        method_started = False

    # set output data base on gathered interface information. must be sorted
    idx_data = ''
    iface_data = 'class {:s}(com.IUnknown):\n'.format(interface_name)
    method_data = ''
    for idx, method in sorted(interface.items()):
        idx_str = '{:s}_{:s}_Idx'.format(interface_name, method['name'])
        idx_data += '{:s} = {:d}\n'.format(idx_str, idx)
        # for each interface, add method parameters
        # build method
        in_params_str = build_in_params_str(method['params'])
        method_data += '\n    def {:s}({:s}):\n'.format(method['name'], in_params_str)
        method_data += build_prototype(method['params'])
        method_data += build_paramflags(method['params'])
        pt_str = '        _{:s} = prototype('.format(method['name'])
        sp = len(pt_str)
        method_data += '{:s}{:s},\n{:s}\'{:s}\',\n{:s}paramflags)\n'.format(pt_str,
                                                                            idx_str,
                                                                            ' ' * sp,
                                                                            method['name'],
                                                                            ' ' * sp)
        method_data += '        _{:s}.errcheck = winapi.RAISE_NON_ZERO_ERR\n'.format(method['name'])
        bstr_alloc_str = build_bstr_alloc(method['params'])
        method_data += bstr_alloc_str
        method_data += build_call(method, True if bstr_alloc_str else False)
        method_data += build_bstr_free(method['params'], True if bstr_alloc_str else False)
        method_data += build_return(method['params'])

    idx_data += '\n\n'
    if not method_data:
        method_data += '    pass'

    # stitch results
    output_data += idx_data + iface_data + method_data

    if output_data:
        mode = 'a'
        if overwrite:
            mode = 'w'
        with open(outfile, mode) as output_file:
            output_file.write(output_data)


def build_bstr_alloc(params):
    bstr_alloc_str = ''
    for param in params:
        if param['type'] == 'BSTR':
            param['call_name'] = '{:s}_bstr'.format(param['converted_name'])
            #  if network_resource is not None else None
            bstr_alloc_str += '        {:s} = winapi.SysAllocString({:s}) if {:s} is not None else None\n'.format(
                param['call_name'],
                param['converted_name'],
                param['converted_name'])
    return bstr_alloc_str


def build_bstr_free(params, add_finally):
    if add_finally:
        bstr_free_str = '        finally:\n'
    else:
        bstr_free_str = ''
    for param in params:
        if param['type'] == 'BSTR':
            bstr_free_str += '            if {:s} is not None:\n                winapi.SysFreeString({:s})\n'.format(
                param['call_name'],
                param['call_name'])
    return bstr_free_str


def build_call(method, add_try):
    call_str = ''
    ret_count = 1
    ret_ptrs = []
    try_str = ''
    if add_try:
        try_str = '        try:\n    '
    for param in method['params']:
        if param['direction'] == '_Out_':
            ret_obj_name = 'return_obj' if ret_count < 2 else 'return_obj{:d}'.format(ret_count)
            ret_ptrs.append(ret_obj_name)
            ret_count += 1

    if ret_ptrs:
        call_str += '{:s}        {:s} = _{:s}(self.this'.format(try_str,
                                                                ', '.join([ret_ptr for ret_ptr in ret_ptrs]),
                                                                method['name'])
    else:
        call_str += '{:s}        _{:s}(self.this'.format(try_str, method['name'])

    sp = len(call_str.split('(')[0]) - len(try_str) + (4 if try_str else 0)
    for param in method['params']:
        if param['direction'] == '_In_':
            in_call_str = param['call_name']
            if '*' in param['type']:
                if param['type'].startswith('IWbem'):
                    in_call_str = '{:s}.this if {:s} else None'.format(param['call_name'],
                                                                       param['call_name'])
                else:
                    in_call_str = 'ctypes.byref({:s}) if {:s} else None'.format(param['call_name'],
                                                                                param['call_name'])
            call_str += ',\n{:s}{:s}'.format(' ' * (sp + 1), in_call_str)

    call_str += '\n{:s})\n'.format(' ' * (sp + 1))
    return call_str


def build_return(params):
    return_str = ''
    try_str = '        try:\n'
    except_str = '        except WindowsError:\n            '
    ret_count = 1
    ret_objs = []
    for param in params:
        if param['direction'] == '_Out_':
            ret_obj_name = 'return_obj' if ret_count < 2 else 'return_obj{:d}'.format(ret_count)
            ret_objs.append(ret_obj_name)
            ret_count += 1
            if param['pointer_count'] > 1:
                if param['type'] == 'SAFEARRAY**':
                    return_str += '        {:s} = ctypes.cast(wintypes.LPVOID({:s}), ctypes.POINTER(winapi.SAFEARRAY))\n'.format(  # NOQA
                        ret_obj_name,
                        ret_obj_name)
                    # return_str += '        {:s} = ctypes.cast(wintypes.LPVOID({:s}.contents.value), ctypes.POINTER(winapi.SAFEARRAY))\n'.format(  # NOQA
                    #     ret_obj_name,
                    #     ret_obj_name)
                else:
                    # return_str += '{:s}            {:s} = {:s}(wintypes.LPVOID({:s}.contents.value))\n{:s}{:s} = None\n'.format(  # NOQA
                    #     try_str,
                    #     ret_obj_name,
                    #     convert_type(param['type'].replace('*', '').strip()),
                    #     ret_obj_name,
                    #     except_str,
                    #     ret_obj_name)
                    return_str += '{:s}            {:s} = {:s}({:s})\n{:s}{:s} = None\n'.format(
                        try_str,
                        ret_obj_name,
                        convert_type(param['type'].replace('*', '').strip()),
                        ret_obj_name,
                        except_str,
                        ret_obj_name)

            elif param['pointer_count'] == 1 and 'BSTR' in param['type']:
                return_str += '        {:s} = winapi.convert_bstr_to_str({:s})\n'.format(ret_obj_name,
                                                                                         ret_obj_name)
    if ret_objs:
        return_str += '        return {:s}\n'.format(', '.join([ret_obj for ret_obj in ret_objs]))
    return return_str


def convert_type(type_str):
    if type_str == 'long':
        ret_str = 'ctypes.c_long'
    elif type_str == 'LPWSTR':
        ret_str = 'wintypes.LPWSTR'
    elif type_str == 'LPCWSTR':
        ret_str = 'wintypes.LPCWSTR'
    elif type_str == 'LONG':
        ret_str = 'wintypes.LONG'
    elif type_str == 'ULONG':
        ret_str = 'wintypes.ULONG'
    elif type_str == 'VARIANT':
        ret_str = 'winapi.VARIANT'
    elif type_str == 'SAFEARRAY':
        ret_str = 'winapi.SAFEARRAY'
    elif '**' in type_str or '* *' in type_str:
        ret_str = 'ctypes.POINTER(wintypes.LPVOID)'
    elif '*' in type_str:
        ret_str = 'ctypes.POINTER({:s})'.format(convert_type(type_str.replace('*', '').strip()))
    else:
        ret_str = type_str
    return ret_str


def build_prototype(params):
    proto = '        prototype = ctypes.WINFUNCTYPE(HRESULT'
    for param in params:
        proto += ',\n{:s}{:s}'.format(' ' * 39, convert_type(param['type']))
    proto += ')\n\n'
    return proto


def build_paramflags(params):
    param_flags = '        paramflags = ('
    started = False
    for param in params:
        if started:
            sp = ' ' * 22
        else:
            sp = ''

        started = True
        param_flag = '{:s}({:s}, \'{:s}\'),\n'.format(sp,
                                                      param['direction'],
                                                      param['name'])

        # if param['direction'] == '_In_':
        #     started = True
        #     param_flag = '{:s}({:s}, \'{:s}\'),\n'.format(sp, param['direction'], param['name'])
        # else:
        #     if param['pointer_count'] > 1:
        #         param_flag = '{:s}({:s}, \'{:s}\', ctypes.pointer(wintypes.LPVOID(None))),\n'.format(sp,
        #                                                                                              param['direction'],
        #                                                                                              param['name'])
        #     else:
        #         param_flag = '{:s}({:s}, \'{:s}\'),\n'.format(sp,
        #                                                       param['direction'],
        #                                                       param['name'])
        param_flags += param_flag
    if len(params):
        param_flags += '{:s})\n\n'.format(' ' * 22)
    else:
        param_flags += ')\n\n'
    return param_flags


def build_in_params_str(params):
    in_params = 'self'
    for param in params:
        if param['direction'] == '_In_':
            in_params += ', '
            in_params += param['converted_name']
    return in_params


def convert_param(method, param):
    # remove notation, split by upper, convert to lowercase
    param_sanitized = param.replace('*', '')
    substr = param_sanitized
    try:
        substr = re.search('([A-Z]\w+)', param_sanitized).group(1)
    except:
        pass
    case_re = re.compile(r'((?<=[a-z0-9])[A-Z]|(?!^)[A-Z](?=[a-z]))')
    converted_param = case_re.sub(r'_\1', substr).lower()
    if converted_param in keyword.kwlist or converted_param in dir(__builtins__):
        converted_param += '_param'
    # check for duplicates. if seen, append number to end
    if 'params' in method and len([param for param in method['params'] if param['name'] == converted_param]):
        param_names = [param['name'] for param in method['params']]
        for x in range(2, 10):
            count_name = '{:s}{:d}'.format(converted_param, x)
            if count_name not in param_names:
                converted_param = count_name
                break
    return converted_param


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Convert C++ interface into python module for use by cWMI')
    parser.add_argument('--infile', help='Input header file that contains the interface', required=True)
    parser.add_argument('--outfile', help='Output module file to write to', required=True)
    parser.add_argument('--interface', help='Name of the interface (case sensitive)', required=True)
    parser.add_argument('--overwrite', action='store_true', help='Overwrite existing file')
    args = parser.parse_args()
    print('Converting interface {:s}'.format(args.interface))
    convert_interface(args.infile,
                      args.outfile,
                      args.interface,
                      args.overwrite)
    print('Done!')
