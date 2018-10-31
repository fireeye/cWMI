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

from .wmi import *
from .winapi import *


def create_variant(val, v_type=None):
    '''
    Creates a VARIANT instance from a given value

    Args:
        val (any): The value to use when creating the VARIANT instance
        v_type (int): Variant type

    Returns:
        Initialized VARIANT instance
    '''
    var = winapi.VARIANT()
    winapi.SET_VT(var, winapi.VT_NULL)
    if v_type:
        if v_type == winapi.VT_I4:
            winapi.SET_VT(var, winapi.VT_I4)
            winapi.V_VAR(var).lVal = ctypes.c_int32(val)
        elif v_type == winapi.VT_R4:
            winapi.SET_VT(var, winapi.VT_R4)
            winapi.V_VAR(var).fltVal = ctypes.c_float(val)
        elif v_type == winapi.VT_LPWSTR:
            winapi.SET_VT(var, winapi.VT_LPWSTR)
            winapi.V_VAR(var).bstrVal = val
        elif v_type == winapi.VT_BSTR:
            winapi.SET_VT(var, winapi.VT_BSTR)
            bstr = winapi.SysAllocString(val)
            winapi.V_VAR(var).bstrVal = bstr
        else:
            raise NotImplemented()
    else:
        if isinstance(val, int):
            winapi.SET_VT(var, winapi.VT_I4)
            winapi.V_VAR(var).lVal = ctypes.c_int32(val)
        elif isinstance(val, float):
            winapi.SET_VT(var, winapi.VT_R4)
            winapi.V_VAR(var).fltVal = ctypes.c_float(val)
        elif isinstance(val, str):
            winapi.SET_VT(var, winapi.VT_BSTR)
            bstr = winapi.SysAllocString(val)
            winapi.V_VAR(var).bstrVal = bstr
        else:
            raise NotImplemented()
    return var


def destroy_variant(var):
    '''
    Destroys an instance of a VARIANT

    Args:
        var (VARIANT): Instance to destroy

    Returns:
        Nothing
    '''
    if winapi.V_VT(var) == winapi.VT_BSTR:
        winapi.SysFreeString(winapi.V_VAR3(var).bstrVal)


def obj_to_dict(obj, include_var_type=False, flags=0):
    '''
    Converts WMI object to python dictionary

    Args:
        obj (any): The object to convert
        include_var_type (bool): Include variant type
        flags (int): Flags to pass into WMI API(s)

    Returns:
        Dictionary containing object information
    '''
    ret = {}
    obj.BeginEnumeration(flags)
    while True:
        try:
            prop_name, var, _, _ = obj.Next(flags)
            if include_var_type:
                ret[prop_name] = winapi.V_TO_VT_DICT(var)
            else:
                ret[prop_name] = winapi.V_TO_TYPE(var)
        except WindowsError:
            break
    obj.EndEnumeration()
    return ret


def safe_array_to_list(safe_array, element_type):
    '''
    Converts SAFEARRAY structure to python list

    Args:
        safe_array (SAFEARRAY): Array structure to convert
        element_type (any): Type of the elements contained in the structure

    Returns:
        List containing the converted array elements
    '''
    ret = []
    data = winapi.SafeArrayAccessData(safe_array)
    str_array = ctypes.cast(data, ctypes.POINTER(element_type))
    try:
        for i in range(safe_array.contents.cbElements):
            ret.append(str_array[i])
    finally:
        winapi.SafeArrayUnaccessData(safe_array)
    winapi.SafeArrayDestroy(safe_array)
    return ret


def get_object_info(namespace, obj_name, values, user=None, password=None, include_var_type=False, flags=0):
    '''
    Gets desired values from an object.

    Args:
        namespace (str): Namespace to connect to
        obj_name (str): Path to object instance/class
        values (list): List of values to pull from object
        user (str): Username to use when connecting to remote system
        password (str): Password to use for remote connection
        include_var_type (bool): Include variant type
        flags (int): Flags to pass into WMI API(s)

    Returns:
        Dictionary containing object data
    '''
    ret = {}
    with WMI(namespace, user, password) as svc:
        with svc.GetObject(obj_name, flags, None) as obj:
            for value in values:
                var, _, _ = obj.Get(value, flags)
                if include_var_type:
                    ret[value] = winapi.V_TO_VT_DICT(var)
                else:
                    ret[value] = winapi.V_TO_TYPE(var)
    return ret


def get_all_object_info(namespace, obj_name, user=None, password=None, include_var_type=False, flags=0):
    '''
    Gets all data from an object

    Args:
        namespace (str): Namespace to connect to
        obj_name (str): Path to object instance/class
        user (str): Username to use when connecting to remote system
        password (str): Password to use for remote connection
        include_var_type (bool): Include variant type
        flags (int): Flags to pass into WMI API(s)

    Returns:
        Dictionary containing object data
    '''
    with WMI(namespace, user, password) as svc:
        with svc.GetObject(obj_name, flags, None) as obj:
            ret = obj_to_dict(obj, include_var_type, flags)
    return ret


def get_object_names(namespace, obj_name, user=None, password=None, flags=0):
    '''
    Gets all names from a given object

    Args:
        namespace (str): Namespace to connect to
        obj_name (str): Path to object instance/class
        user (str): Username to use when connecting to remote system
        password (str): Password to use for remote connection
        flags (int): Flags to pass into WMI API(s)

    Returns:
        List containing object names
    '''
    with WMI(namespace, user, password) as svc:
        with svc.GetObject(obj_name, flags, None) as obj:
            ret = safe_array_to_list(obj.GetNames(None, flags, None), ctypes.c_wchar_p)
    return ret


def query(namespace, query_str, query_type='WQL', user=None, password=None, include_var_type=False, flags=0):
    '''
    Performs a WMI query

    Args:
        namespace (str): Namespace to connect to
        query_str (str): String to use for the query
        query_type (str): Path to object instance/class
        user (str): Username to use when connecting to remote system
        password (str): Password to use for remote connection
        include_var_type (bool): Include variant type
        flags (int): Flags to pass into WMI API(s)

    Returns:
        Dictionary containing all objects returned by query
    '''
    ret = {}
    with WMI(namespace, user, password) as svc:
        with svc.ExecQuery(query_type, query_str, flags, None) as enum:
            # do query, then get all data
            while True:
                try:
                    with enum.Next(winapi.WBEM_INFINITE) as obj:
                        obj.BeginEnumeration(flags)

                        inst_data = {}
                        while True:
                            try:
                                prop_name, var, _, _ = obj.Next(flags)
                                if include_var_type:
                                    inst_data[prop_name] = winapi.V_TO_VT_DICT(var)
                                else:
                                    inst_data[prop_name] = winapi.V_TO_TYPE(var)
                            except WindowsError:
                                break

                        obj.EndEnumeration()
                        var, _, _ = obj.Get('__RELPATH', flags)
                        inst_relpath = winapi.V_TO_STR(var)
                        # use instance relative path for key
                        ret[inst_relpath] = inst_data
                except WindowsError:
                    break
    return ret


def does_object_exist(namespace, obj_name, user=None, password=None, flags=0):
    '''
    Tests if a given object exists

    Args:
        namespace (str): Namespace to connect to
        obj_name (str): Path to object instance/class
        user (str): Username to use when connecting to remote system
        password (str): Password to use for remote connection
        flags (int): Flags to pass into WMI API(s)

    Returns:
        True if object exists, or False if not
    '''
    ret = True
    with WMI(namespace, user, password) as svc:
        try:
            with svc.GetObject(obj_name, flags, None):
                pass
        except WindowsError:
            ret = False
    return ret


def get_method_names(namespace, obj_name, user=None, password=None, flags=0):
    '''
    Gets all method names from a given object

    Args:
        namespace (str): Namespace to connect to
        obj_name (str): Path to object instance/class
        user (str): Username to use when connecting to remote system
        password (str): Password to use for remote connection
        flags (int): Flags to pass into WMI API(s)

    Returns:
        List of method names
    '''
    ret = []
    with WMI(namespace, user, password) as svc:
        with svc.GetObject(obj_name, flags, None) as obj:
            obj.BeginMethodEnumeration(flags)
            while True:
                try:
                    method_name, in_param, out_param = obj.NextMethod(flags)
                    ret.append(method_name)
                except WindowsError:
                    break
            obj.EndMethodEnumeration()
    return ret


def get_method_info(namespace, obj_name, user=None, password=None, flags=0):
    '''
    Gets method information for a given object/method

    Args:
        namespace (str): Namespace to connect to
        obj_name (str): Path to object instance/class
        user (str): Username to use when connecting to remote system
        password (str): Password to use for remote connection
        flags (int): Flags to pass into WMI API(s)

    Returns:
        Dictionary containing method parameter info
    '''
    ret = {}
    with WMI(namespace, user, password) as svc:
        with svc.GetObject(obj_name, flags, None) as obj:
            obj.BeginMethodEnumeration(flags)

            while True:
                try:
                    method_name, in_sig, out_sig = obj.NextMethod(flags)
                    in_sig_vals = {}
                    if in_sig:
                        in_sig_vals = obj_to_dict(in_sig)
                        in_sig.Release()

                    out_sig_vals = {}
                    if out_sig:
                        out_sig_vals = obj_to_dict(out_sig)
                        out_sig.Release()

                    ret[method_name] = {'in_signature': in_sig_vals,
                                        'out_signature': out_sig_vals}
                except WindowsError:
                    break

            obj.EndMethodEnumeration()
    return ret


def call_method(namespace, obj_path, method_name, input_params=None, user=None, password=None, flags=0):
    '''
    Calls object method

    Args:
        namespace (str): Namespace to connect to
        obj_path (str): Path to object instance/class
        method_name (str): Name of the method to call
        input_params (dict): Method input parameters
        user (str): Username to use when connecting to remote system
        password (str): Password to use for remote connection
        flags (int): Flags to pass into WMI API(s)

    Returns:
        Dictionary containing data returned from the called method
    '''
    ret = {}
    with WMI(namespace, user, password) as svc:
        class_name = obj_path
        if '.' in class_name:
            class_name = class_name.split('.')[0]

        with svc.GetObject(class_name, flags, None) as obj:
            if input_params:
                in_obj_param, out_obj_param = obj.GetMethod(method_name, flags)
                if in_obj_param:
                    for prop, var in input_params.items():
                        if isinstance(var, dict) or not isinstance(var, VARIANT):
                            if isinstance(var, dict):
                                if 'type' in var:
                                    in_var = create_variant(var['value'], var['type'])
                                else:
                                    raise Exception('Variant type must be specified')
                            else:
                                in_var = create_variant(var)
                            in_obj_param.Put(prop, flags, in_var, 0)
                            destroy_variant(in_var)
                        else:
                            in_obj_param.Put(prop, flags, var, 0)
            else:
                in_obj_param, out_obj_param = obj.GetMethod(method_name, flags)

            if out_obj_param:
                out_obj_param.Release()

            out_obj = svc.ExecMethod(obj_path, method_name, flags, None, in_obj_param)
            if in_obj_param:
                in_obj_param.Release()
            if out_obj:
                ret = obj_to_dict(out_obj)
                out_obj.Release()
    return ret
