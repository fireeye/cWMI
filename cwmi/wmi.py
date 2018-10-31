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


import ctypes
from ctypes import wintypes

from . import com
from . import winapi
from .wintype import HRESULT, BSTR, CIMTYPE


_In_ = 1
_Out_ = 2

IWbemObjectSink_Indicate_Idx = 3
IWbemObjectSink_SetStatus_Idx = 4


class IWbemObjectSink(com.IUnknown):

    def Indicate(self, object_count, obj_array):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'lObjectCount'),
                      (_In_, 'apObjArray'),
                      )

        _Indicate = prototype(IWbemObjectSink_Indicate_Idx,
                              'Indicate',
                              paramflags)
        _Indicate.errcheck = winapi.RAISE_NON_ZERO_ERR
        _Indicate(self.this,
                  object_count,
                  obj_array.this if obj_array else None
                  )

    def SetStatus(self, flags, result, param, obj_param):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       HRESULT,
                                       BSTR,
                                       ctypes.POINTER(IWbemClassObject))

        paramflags = ((_In_, 'lFlags'),
                      (_In_, 'hResult'),
                      (_In_, 'strParam'),
                      (_In_, 'pObjParam'),
                      )

        _SetStatus = prototype(IWbemObjectSink_SetStatus_Idx,
                               'SetStatus',
                               paramflags)
        _SetStatus.errcheck = winapi.RAISE_NON_ZERO_ERR
        param_bstr = winapi.SysAllocString(param) if param is not None else None
        try:
            _SetStatus(self.this,
                       flags,
                       result,
                       param_bstr,
                       obj_param.this if obj_param else None
                       )
        finally:
            if param_bstr is not None:
                winapi.SysFreeString(param_bstr)


IWbemQualifierSet_Get_Idx = 3
IWbemQualifierSet_Put_Idx = 4
IWbemQualifierSet_Delete_Idx = 5
IWbemQualifierSet_GetNames_Idx = 6
IWbemQualifierSet_BeginEnumeration_Idx = 7
IWbemQualifierSet_Next_Idx = 8
IWbemQualifierSet_EndEnumeration_Idx = 9


class IWbemQualifierSet(com.IUnknown):

    def Get(self, name, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(winapi.VARIANT),
                                       ctypes.POINTER(ctypes.c_long))

        paramflags = ((_In_, 'wszName'),
                      (_In_, 'lFlags'),
                      (_Out_, 'pVal'),
                      (_Out_, 'plFlavor'),
                      )

        _Get = prototype(IWbemQualifierSet_Get_Idx,
                         'Get',
                         paramflags)
        _Get.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj, return_obj2 = _Get(self.this,
                                       name,
                                       flags
                                       )
        return return_obj, return_obj2

    def Put(self, name, val, flavor):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.POINTER(winapi.VARIANT),
                                       ctypes.c_long)

        paramflags = ((_In_, 'wszName'),
                      (_In_, 'pVal'),
                      (_In_, 'lFlavor'),
                      )

        _Put = prototype(IWbemQualifierSet_Put_Idx,
                         'Put',
                         paramflags)
        _Put.errcheck = winapi.RAISE_NON_ZERO_ERR
        _Put(self.this,
             name,
             ctypes.byref(val) if val else None,
             flavor
             )

    def Delete(self, name):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR)

        paramflags = ((_In_, 'wszName'),
                      )

        _Delete = prototype(IWbemQualifierSet_Delete_Idx,
                            'Delete',
                            paramflags)
        _Delete.errcheck = winapi.RAISE_NON_ZERO_ERR
        _Delete(self.this,
                name
                )

    def GetNames(self, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'lFlags'),
                      (_Out_, 'pNames'),
                      )

        _GetNames = prototype(IWbemQualifierSet_GetNames_Idx,
                              'GetNames',
                              paramflags)
        _GetNames.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetNames(self.this,
                               flags
                               )
        return_obj = ctypes.cast(wintypes.LPVOID(return_obj), ctypes.POINTER(winapi.SAFEARRAY))
        return return_obj

    def BeginEnumeration(self, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long)

        paramflags = ((_In_, 'lFlags'),
                      )

        _BeginEnumeration = prototype(IWbemQualifierSet_BeginEnumeration_Idx,
                                      'BeginEnumeration',
                                      paramflags)
        _BeginEnumeration.errcheck = winapi.RAISE_NON_ZERO_ERR
        _BeginEnumeration(self.this,
                          flags
                          )

    def Next(self, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(BSTR),
                                       ctypes.POINTER(winapi.VARIANT),
                                       ctypes.POINTER(ctypes.c_long))

        paramflags = ((_In_, 'lFlags'),
                      (_Out_, 'pstrName'),
                      (_Out_, 'pVal'),
                      (_Out_, 'plFlavor'),
                      )

        _Next = prototype(IWbemQualifierSet_Next_Idx,
                          'Next',
                          paramflags)
        _Next.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj, return_obj2, return_obj3 = _Next(self.this,
                                                     flags
                                                     )
        return_obj = winapi.convert_bstr_to_str(return_obj)
        return return_obj, return_obj2, return_obj3

    def EndEnumeration(self):
        prototype = ctypes.WINFUNCTYPE(HRESULT)

        paramflags = ()

        _EndEnumeration = prototype(IWbemQualifierSet_EndEnumeration_Idx,
                                    'EndEnumeration',
                                    paramflags)
        _EndEnumeration.errcheck = winapi.RAISE_NON_ZERO_ERR
        _EndEnumeration(self.this
                        )


IWbemClassObject_GetQualifierSet_Idx = 3
IWbemClassObject_Get_Idx = 4
IWbemClassObject_Put_Idx = 5
IWbemClassObject_Delete_Idx = 6
IWbemClassObject_GetNames_Idx = 7
IWbemClassObject_BeginEnumeration_Idx = 8
IWbemClassObject_Next_Idx = 9
IWbemClassObject_EndEnumeration_Idx = 10
IWbemClassObject_GetPropertyQualifierSet_Idx = 11
IWbemClassObject_Clone_Idx = 12
IWbemClassObject_GetObjectText_Idx = 13
IWbemClassObject_SpawnDerivedClass_Idx = 14
IWbemClassObject_SpawnInstance_Idx = 15
IWbemClassObject_CompareTo_Idx = 16
IWbemClassObject_GetPropertyOrigin_Idx = 17
IWbemClassObject_InheritsFrom_Idx = 18
IWbemClassObject_GetMethod_Idx = 19
IWbemClassObject_PutMethod_Idx = 20
IWbemClassObject_DeleteMethod_Idx = 21
IWbemClassObject_BeginMethodEnumeration_Idx = 22
IWbemClassObject_NextMethod_Idx = 23
IWbemClassObject_EndMethodEnumeration_Idx = 24
IWbemClassObject_GetMethodQualifierSet_Idx = 25
IWbemClassObject_GetMethodOrigin_Idx = 26


class IWbemClassObject(com.IUnknown):

    def GetQualifierSet(self):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_Out_, 'ppQualSet'),
                      )

        _GetQualifierSet = prototype(IWbemClassObject_GetQualifierSet_Idx,
                                     'GetQualifierSet',
                                     paramflags)
        _GetQualifierSet.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetQualifierSet(self.this
                                      )
        try:
            return_obj = IWbemQualifierSet(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def Get(self, name, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(winapi.VARIANT),
                                       ctypes.POINTER(CIMTYPE),
                                       ctypes.POINTER(ctypes.c_long))

        paramflags = ((_In_, 'wszName'),
                      (_In_, 'lFlags'),
                      (_Out_, 'pVal'),
                      (_Out_, 'pType'),
                      (_Out_, 'plFlavor'),
                      )

        _Get = prototype(IWbemClassObject_Get_Idx,
                         'Get',
                         paramflags)
        _Get.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj, return_obj2, return_obj3 = _Get(self.this,
                                                    name,
                                                    flags
                                                    )
        return return_obj, return_obj2, return_obj3

    def Put(self, name, flags, val, type_param):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(winapi.VARIANT),
                                       CIMTYPE)

        paramflags = ((_In_, 'wszName'),
                      (_In_, 'lFlags'),
                      (_In_, 'pVal'),
                      (_In_, 'Type'),
                      )

        _Put = prototype(IWbemClassObject_Put_Idx,
                         'Put',
                         paramflags)
        _Put.errcheck = winapi.RAISE_NON_ZERO_ERR
        _Put(self.this,
             name,
             flags,
             ctypes.byref(val) if val else None,
             type_param
             )

    def Delete(self, name):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR)

        paramflags = ((_In_, 'wszName'),
                      )

        _Delete = prototype(IWbemClassObject_Delete_Idx,
                            'Delete',
                            paramflags)
        _Delete.errcheck = winapi.RAISE_NON_ZERO_ERR
        _Delete(self.this,
                name
                )

    def GetNames(self, qualifier_name, flags, qualifier_val):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(winapi.VARIANT),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'wszQualifierName'),
                      (_In_, 'lFlags'),
                      (_In_, 'pQualifierVal'),
                      (_Out_, 'pNames'),
                      )

        _GetNames = prototype(IWbemClassObject_GetNames_Idx,
                              'GetNames',
                              paramflags)
        _GetNames.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetNames(self.this,
                               qualifier_name,
                               flags,
                               ctypes.byref(qualifier_val) if qualifier_val else None
                               )
        return_obj = ctypes.cast(wintypes.LPVOID(return_obj), ctypes.POINTER(winapi.SAFEARRAY))
        return return_obj

    def BeginEnumeration(self, enum_flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long)

        paramflags = ((_In_, 'lEnumFlags'),
                      )

        _BeginEnumeration = prototype(IWbemClassObject_BeginEnumeration_Idx,
                                      'BeginEnumeration',
                                      paramflags)
        _BeginEnumeration.errcheck = winapi.RAISE_NON_ZERO_ERR
        _BeginEnumeration(self.this,
                          enum_flags
                          )

    def Next(self, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(BSTR),
                                       ctypes.POINTER(winapi.VARIANT),
                                       ctypes.POINTER(CIMTYPE),
                                       ctypes.POINTER(ctypes.c_long))

        paramflags = ((_In_, 'lFlags'),
                      (_Out_, 'strName'),
                      (_Out_, 'pVal'),
                      (_Out_, 'pType'),
                      (_Out_, 'plFlavor'),
                      )

        _Next = prototype(IWbemClassObject_Next_Idx,
                          'Next',
                          paramflags)
        _Next.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj, return_obj2, return_obj3, return_obj4 = _Next(self.this,
                                                                  flags
                                                                  )
        return_obj = winapi.convert_bstr_to_str(return_obj)
        return return_obj, return_obj2, return_obj3, return_obj4

    def EndEnumeration(self):
        prototype = ctypes.WINFUNCTYPE(HRESULT)

        paramflags = ()

        _EndEnumeration = prototype(IWbemClassObject_EndEnumeration_Idx,
                                    'EndEnumeration',
                                    paramflags)
        _EndEnumeration.errcheck = winapi.RAISE_NON_ZERO_ERR
        _EndEnumeration(self.this
                        )

    def GetPropertyQualifierSet(self, property_param):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'wszProperty'),
                      (_Out_, 'ppQualSet'),
                      )

        _GetPropertyQualifierSet = prototype(IWbemClassObject_GetPropertyQualifierSet_Idx,
                                             'GetPropertyQualifierSet',
                                             paramflags)
        _GetPropertyQualifierSet.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetPropertyQualifierSet(self.this,
                                              property_param
                                              )
        try:
            return_obj = IWbemQualifierSet(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def Clone(self):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_Out_, 'ppCopy'),
                      )

        _Clone = prototype(IWbemClassObject_Clone_Idx,
                           'Clone',
                           paramflags)
        _Clone.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _Clone(self.this
                            )
        try:
            return_obj = IWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def GetObjectText(self, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(BSTR))

        paramflags = ((_In_, 'lFlags'),
                      (_Out_, 'pstrObjectText'),
                      )

        _GetObjectText = prototype(IWbemClassObject_GetObjectText_Idx,
                                   'GetObjectText',
                                   paramflags)
        _GetObjectText.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetObjectText(self.this,
                                    flags
                                    )
        return_obj = winapi.convert_bstr_to_str(return_obj)
        return return_obj

    def SpawnDerivedClass(self, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'lFlags'),
                      (_Out_, 'ppNewClass'),
                      )

        _SpawnDerivedClass = prototype(IWbemClassObject_SpawnDerivedClass_Idx,
                                       'SpawnDerivedClass',
                                       paramflags)
        _SpawnDerivedClass.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _SpawnDerivedClass(self.this,
                                        flags
                                        )
        try:
            return_obj = IWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def SpawnInstance(self, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'lFlags'),
                      (_Out_, 'ppNewInstance'),
                      )

        _SpawnInstance = prototype(IWbemClassObject_SpawnInstance_Idx,
                                   'SpawnInstance',
                                   paramflags)
        _SpawnInstance.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _SpawnInstance(self.this,
                                    flags
                                    )
        try:
            return_obj = IWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def CompareTo(self, flags, compare_to):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemClassObject))

        paramflags = ((_In_, 'lFlags'),
                      (_In_, 'pCompareTo'),
                      )

        _CompareTo = prototype(IWbemClassObject_CompareTo_Idx,
                               'CompareTo',
                               paramflags)
        _CompareTo.errcheck = winapi.RAISE_NON_ZERO_ERR
        _CompareTo(self.this,
                   flags,
                   compare_to.this if compare_to else None
                   )

    def GetPropertyOrigin(self, name):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.POINTER(BSTR))

        paramflags = ((_In_, 'wszName'),
                      (_Out_, 'pstrClassName'),
                      )

        _GetPropertyOrigin = prototype(IWbemClassObject_GetPropertyOrigin_Idx,
                                       'GetPropertyOrigin',
                                       paramflags)
        _GetPropertyOrigin.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetPropertyOrigin(self.this,
                                        name
                                        )
        return_obj = winapi.convert_bstr_to_str(return_obj)
        return return_obj

    def InheritsFrom(self, ancestor):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR)

        paramflags = ((_In_, 'strAncestor'),
                      )

        _InheritsFrom = prototype(IWbemClassObject_InheritsFrom_Idx,
                                  'InheritsFrom',
                                  paramflags)
        _InheritsFrom.errcheck = winapi.RAISE_NON_ZERO_ERR
        _InheritsFrom(self.this,
                      ancestor
                      )

    def GetMethod(self, name, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(wintypes.LPVOID),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'wszName'),
                      (_In_, 'lFlags'),
                      (_Out_, 'ppInSignature'),
                      (_Out_, 'ppOutSignature'),
                      )

        _GetMethod = prototype(IWbemClassObject_GetMethod_Idx,
                               'GetMethod',
                               paramflags)
        _GetMethod.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj, return_obj2 = _GetMethod(self.this,
                                             name,
                                             flags
                                             )
        try:
            return_obj = IWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        try:
            return_obj2 = IWbemClassObject(return_obj2)
        except WindowsError:
            return_obj2 = None
        return return_obj, return_obj2

    def PutMethod(self, name, flags, in_signature, out_signature):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemClassObject),
                                       ctypes.POINTER(IWbemClassObject))

        paramflags = ((_In_, 'wszName'),
                      (_In_, 'lFlags'),
                      (_In_, 'pInSignature'),
                      (_In_, 'pOutSignature'),
                      )

        _PutMethod = prototype(IWbemClassObject_PutMethod_Idx,
                               'PutMethod',
                               paramflags)
        _PutMethod.errcheck = winapi.RAISE_NON_ZERO_ERR
        _PutMethod(self.this,
                   name,
                   flags,
                   in_signature.this if in_signature else None,
                   out_signature.this if out_signature else None
                   )

    def DeleteMethod(self, name):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR)

        paramflags = ((_In_, 'wszName'),
                      )

        _DeleteMethod = prototype(IWbemClassObject_DeleteMethod_Idx,
                                  'DeleteMethod',
                                  paramflags)
        _DeleteMethod.errcheck = winapi.RAISE_NON_ZERO_ERR
        _DeleteMethod(self.this,
                      name
                      )

    def BeginMethodEnumeration(self, enum_flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long)

        paramflags = ((_In_, 'lEnumFlags'),
                      )

        _BeginMethodEnumeration = prototype(IWbemClassObject_BeginMethodEnumeration_Idx,
                                            'BeginMethodEnumeration',
                                            paramflags)
        _BeginMethodEnumeration.errcheck = winapi.RAISE_NON_ZERO_ERR
        _BeginMethodEnumeration(self.this,
                                enum_flags
                                )

    def NextMethod(self, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(BSTR),
                                       ctypes.POINTER(wintypes.LPVOID),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'lFlags'),
                      (_Out_, 'pstrName'),
                      (_Out_, 'ppInSignature'),
                      (_Out_, 'ppOutSignature'),
                      )

        _NextMethod = prototype(IWbemClassObject_NextMethod_Idx,
                                'NextMethod',
                                paramflags)
        _NextMethod.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj, return_obj2, return_obj3 = _NextMethod(self.this,
                                                           flags
                                                           )
        return_obj = winapi.convert_bstr_to_str(return_obj)
        try:
            return_obj2 = IWbemClassObject(return_obj2)
        except WindowsError:
            return_obj2 = None
        try:
            return_obj3 = IWbemClassObject(return_obj3)
        except WindowsError:
            return_obj3 = None
        return return_obj, return_obj2, return_obj3

    def EndMethodEnumeration(self):
        prototype = ctypes.WINFUNCTYPE(HRESULT)

        paramflags = ()

        _EndMethodEnumeration = prototype(IWbemClassObject_EndMethodEnumeration_Idx,
                                          'EndMethodEnumeration',
                                          paramflags)
        _EndMethodEnumeration.errcheck = winapi.RAISE_NON_ZERO_ERR
        _EndMethodEnumeration(self.this
                              )

    def GetMethodQualifierSet(self, method):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'wszMethod'),
                      (_Out_, 'ppQualSet'),
                      )

        _GetMethodQualifierSet = prototype(IWbemClassObject_GetMethodQualifierSet_Idx,
                                           'GetMethodQualifierSet',
                                           paramflags)
        _GetMethodQualifierSet.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetMethodQualifierSet(self.this,
                                            method
                                            )
        try:
            return_obj = IWbemQualifierSet(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def GetMethodOrigin(self, method_name):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.POINTER(BSTR))

        paramflags = ((_In_, 'wszMethodName'),
                      (_Out_, 'pstrClassName'),
                      )

        _GetMethodOrigin = prototype(IWbemClassObject_GetMethodOrigin_Idx,
                                     'GetMethodOrigin',
                                     paramflags)
        _GetMethodOrigin.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetMethodOrigin(self.this,
                                      method_name
                                      )
        return_obj = winapi.convert_bstr_to_str(return_obj)
        return return_obj


IEnumWbemClassObject_Reset_Idx = 3
IEnumWbemClassObject_Next_Idx = 4
IEnumWbemClassObject_NextAsync_Idx = 5
IEnumWbemClassObject_Clone_Idx = 6
IEnumWbemClassObject_Skip_Idx = 7


class IEnumWbemClassObject(com.IUnknown):

    def Reset(self):
        prototype = ctypes.WINFUNCTYPE(HRESULT)

        paramflags = ()

        _Reset = prototype(IEnumWbemClassObject_Reset_Idx,
                           'Reset',
                           paramflags)
        _Reset.errcheck = winapi.RAISE_NON_ZERO_ERR
        _Reset(self.this
               )

    def Next(self, timeout):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       wintypes.ULONG,
                                       ctypes.POINTER(wintypes.LPVOID),
                                       ctypes.POINTER(wintypes.ULONG))

        paramflags = ((_In_, 'lTimeout'),
                      (_In_, 'uCount'),
                      (_Out_, 'apObjects'),
                      (_Out_, 'puReturned'),
                      )

        _Next = prototype(IEnumWbemClassObject_Next_Idx,
                          'Next',
                          paramflags)
        _Next.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj, return_obj2 = _Next(self.this,
                                        timeout,
                                        1
                                        )
        try:
            return_obj = IWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def NextAsync(self, count, sink):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.ULONG,
                                       ctypes.POINTER(IWbemObjectSink))

        paramflags = ((_In_, 'uCount'),
                      (_In_, 'pSink'),
                      )

        _NextAsync = prototype(IEnumWbemClassObject_NextAsync_Idx,
                               'NextAsync',
                               paramflags)
        _NextAsync.errcheck = winapi.RAISE_NON_ZERO_ERR
        _NextAsync(self.this,
                   count,
                   sink.this if sink else None
                   )

    def Clone(self):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_Out_, 'ppEnum'),
                      )

        _Clone = prototype(IEnumWbemClassObject_Clone_Idx,
                           'Clone',
                           paramflags)
        _Clone.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _Clone(self.this
                            )
        try:
            return_obj = IEnumWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def Skip(self, timeout, count):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       wintypes.ULONG)

        paramflags = ((_In_, 'lTimeout'),
                      (_In_, 'nCount'),
                      )

        _Skip = prototype(IEnumWbemClassObject_Skip_Idx,
                          'Skip',
                          paramflags)
        _Skip.errcheck = winapi.RAISE_NON_ZERO_ERR
        _Skip(self.this,
              timeout,
              count
              )


IWbemCallResult_GetResultObject_Idx = 3
IWbemCallResult_GetResultString_Idx = 4
IWbemCallResult_GetResultServices_Idx = 5
IWbemCallResult_GetCallStatus_Idx = 6


class IWbemCallResult(com.IUnknown):

    def GetResultObject(self, timeout):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'lTimeout'),
                      (_Out_, 'ppResultObject'),
                      )

        _GetResultObject = prototype(IWbemCallResult_GetResultObject_Idx,
                                     'GetResultObject',
                                     paramflags)
        _GetResultObject.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetResultObject(self.this,
                                      timeout
                                      )
        try:
            return_obj = IWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def GetResultString(self, timeout):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(BSTR))

        paramflags = ((_In_, 'lTimeout'),
                      (_Out_, 'pstrResultString'),
                      )

        _GetResultString = prototype(IWbemCallResult_GetResultString_Idx,
                                     'GetResultString',
                                     paramflags)
        _GetResultString.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetResultString(self.this,
                                      timeout
                                      )
        return_obj = winapi.convert_bstr_to_str(return_obj)
        return return_obj

    def GetResultServices(self, timeout):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'lTimeout'),
                      (_Out_, 'ppServices'),
                      )

        _GetResultServices = prototype(IWbemCallResult_GetResultServices_Idx,
                                       'GetResultServices',
                                       paramflags)
        _GetResultServices.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetResultServices(self.this,
                                        timeout
                                        )
        try:
            return_obj = IWbemServices(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def GetCallStatus(self, timeout):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(ctypes.c_long))

        paramflags = ((_In_, 'lTimeout'),
                      (_Out_, 'plStatus'),
                      )

        _GetCallStatus = prototype(IWbemCallResult_GetCallStatus_Idx,
                                   'GetCallStatus',
                                   paramflags)
        _GetCallStatus.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetCallStatus(self.this,
                                    timeout
                                    )
        return return_obj


IWbemContext_Clone_Idx = 3
IWbemContext_GetNames_Idx = 4
IWbemContext_BeginEnumeration_Idx = 5
IWbemContext_Next_Idx = 6
IWbemContext_EndEnumeration_Idx = 7
IWbemContext_SetValue_Idx = 8
IWbemContext_GetValue_Idx = 9
IWbemContext_DeleteValue_Idx = 10
IWbemContext_DeleteAll_Idx = 11


class IWbemContext(com.IUnknown):

    def Clone(self):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_Out_, 'ppNewCopy'),
                      )

        _Clone = prototype(IWbemContext_Clone_Idx,
                           'Clone',
                           paramflags)
        _Clone.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _Clone(self.this
                            )
        try:
            return_obj = IWbemContext(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def GetNames(self, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'lFlags'),
                      (_Out_, 'pNames'),
                      )

        _GetNames = prototype(IWbemContext_GetNames_Idx,
                              'GetNames',
                              paramflags)
        _GetNames.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetNames(self.this,
                               flags
                               )
        return_obj = ctypes.cast(wintypes.LPVOID(return_obj), ctypes.POINTER(winapi.SAFEARRAY))
        return return_obj

    def BeginEnumeration(self, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long)

        paramflags = ((_In_, 'lFlags'),
                      )

        _BeginEnumeration = prototype(IWbemContext_BeginEnumeration_Idx,
                                      'BeginEnumeration',
                                      paramflags)
        _BeginEnumeration.errcheck = winapi.RAISE_NON_ZERO_ERR
        _BeginEnumeration(self.this,
                          flags
                          )

    def Next(self, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(BSTR),
                                       ctypes.POINTER(winapi.VARIANT))

        paramflags = ((_In_, 'lFlags'),
                      (_Out_, 'pstrName'),
                      (_Out_, 'pValue'),
                      )

        _Next = prototype(IWbemContext_Next_Idx,
                          'Next',
                          paramflags)
        _Next.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj, return_obj2 = _Next(self.this,
                                        flags
                                        )
        return_obj = winapi.convert_bstr_to_str(return_obj)
        return return_obj, return_obj2

    def EndEnumeration(self):
        prototype = ctypes.WINFUNCTYPE(HRESULT)

        paramflags = ()

        _EndEnumeration = prototype(IWbemContext_EndEnumeration_Idx,
                                    'EndEnumeration',
                                    paramflags)
        _EndEnumeration.errcheck = winapi.RAISE_NON_ZERO_ERR
        _EndEnumeration(self.this
                        )

    def SetValue(self, name, flags, value):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(winapi.VARIANT))

        paramflags = ((_In_, 'wszName'),
                      (_In_, 'lFlags'),
                      (_In_, 'pValue'),
                      )

        _SetValue = prototype(IWbemContext_SetValue_Idx,
                              'SetValue',
                              paramflags)
        _SetValue.errcheck = winapi.RAISE_NON_ZERO_ERR
        _SetValue(self.this,
                  name,
                  flags,
                  ctypes.byref(value) if value else None
                  )

    def GetValue(self, name, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(winapi.VARIANT))

        paramflags = ((_In_, 'wszName'),
                      (_In_, 'lFlags'),
                      (_Out_, 'pValue'),
                      )

        _GetValue = prototype(IWbemContext_GetValue_Idx,
                              'GetValue',
                              paramflags)
        _GetValue.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _GetValue(self.this,
                               name,
                               flags
                               )
        return return_obj

    def DeleteValue(self, name, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       wintypes.LPCWSTR,
                                       ctypes.c_long)

        paramflags = ((_In_, 'wszName'),
                      (_In_, 'lFlags'),
                      )

        _DeleteValue = prototype(IWbemContext_DeleteValue_Idx,
                                 'DeleteValue',
                                 paramflags)
        _DeleteValue.errcheck = winapi.RAISE_NON_ZERO_ERR
        _DeleteValue(self.this,
                     name,
                     flags
                     )

    def DeleteAll(self):
        prototype = ctypes.WINFUNCTYPE(HRESULT)

        paramflags = ()

        _DeleteAll = prototype(IWbemContext_DeleteAll_Idx,
                               'DeleteAll',
                               paramflags)
        _DeleteAll.errcheck = winapi.RAISE_NON_ZERO_ERR
        _DeleteAll(self.this
                   )


IWbemServices_OpenNamespace_Idx = 3
IWbemServices_CancelAsyncCall_Idx = 4
IWbemServices_QueryObjectSink_Idx = 5
IWbemServices_GetObject_Idx = 6
IWbemServices_GetObjectAsync_Idx = 7
IWbemServices_PutClass_Idx = 8
IWbemServices_PutClassAsync_Idx = 9
IWbemServices_DeleteClass_Idx = 10
IWbemServices_DeleteClassAsync_Idx = 11
IWbemServices_CreateClassEnum_Idx = 12
IWbemServices_CreateClassEnumAsync_Idx = 13
IWbemServices_PutInstance_Idx = 14
IWbemServices_PutInstanceAsync_Idx = 15
IWbemServices_DeleteInstance_Idx = 16
IWbemServices_DeleteInstanceAsync_Idx = 17
IWbemServices_CreateInstanceEnum_Idx = 18
IWbemServices_CreateInstanceEnumAsync_Idx = 19
IWbemServices_ExecQuery_Idx = 20
IWbemServices_ExecQueryAsync_Idx = 21
IWbemServices_ExecNotificationQuery_Idx = 22
IWbemServices_ExecNotificationQueryAsync_Idx = 23
IWbemServices_ExecMethod_Idx = 24
IWbemServices_ExecMethodAsync_Idx = 25


class IWbemServices(com.IUnknown):

    def OpenNamespaceWithResult(self, namespace, flags, ctx):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(wintypes.LPVOID),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'strNamespace'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_Out_, 'ppWorkingNamespace'),
                      (_Out_, 'ppResult'),
                      )

        _OpenNamespace = prototype(IWbemServices_OpenNamespace_Idx,
                                   'OpenNamespace',
                                   paramflags)
        _OpenNamespace.errcheck = winapi.RAISE_NON_ZERO_ERR
        namespace_bstr = winapi.SysAllocString(namespace) if namespace is not None else None
        try:
            return_obj, return_obj2 = _OpenNamespace(self.this,
                                                     namespace_bstr,
                                                     flags,
                                                     ctx.this if ctx else None
                                                     )
        finally:
            if namespace_bstr is not None:
                winapi.SysFreeString(namespace_bstr)
        try:
            return_obj = IWbemServices(return_obj)
        except WindowsError:
            return_obj = None
        try:
            return_obj2 = IWbemCallResult(return_obj2)
        except WindowsError:
            return_obj2 = None
        return return_obj, return_obj2

    def OpenNamespace(self, namespace, flags, ctx):
        return_obj, return_obj2 = self.OpenNamespaceWithResult(namespace, flags, ctx)
        if return_obj2:
            return_obj2.Release()
        return return_obj

    def CancelAsyncCall(self, sink):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.POINTER(IWbemObjectSink))

        paramflags = ((_In_, 'pSink'),
                      )

        _CancelAsyncCall = prototype(IWbemServices_CancelAsyncCall_Idx,
                                     'CancelAsyncCall',
                                     paramflags)
        _CancelAsyncCall.errcheck = winapi.RAISE_NON_ZERO_ERR
        _CancelAsyncCall(self.this,
                         sink.this if sink else None
                         )

    def QueryObjectSink(self, flags):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.c_long,
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'lFlags'),
                      (_Out_, 'ppResponseHandler'),
                      )

        _QueryObjectSink = prototype(IWbemServices_QueryObjectSink_Idx,
                                     'QueryObjectSink',
                                     paramflags)
        _QueryObjectSink.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _QueryObjectSink(self.this,
                                      flags
                                      )
        try:
            return_obj = IWbemObjectSink(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def GetObjectWithResult(self, object_path, flags, ctx):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(wintypes.LPVOID),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'strObjectPath'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_Out_, 'ppObject'),
                      (_Out_, 'ppCallResult'),
                      )

        _GetObject = prototype(IWbemServices_GetObject_Idx,
                               'GetObject',
                               paramflags)
        _GetObject.errcheck = winapi.RAISE_NON_ZERO_ERR
        object_path_bstr = winapi.SysAllocString(object_path) if object_path is not None else None
        try:
            return_obj, return_obj2 = _GetObject(self.this,
                                                 object_path_bstr,
                                                 flags,
                                                 ctx.this if ctx else None
                                                 )
        finally:
            if object_path_bstr is not None:
                winapi.SysFreeString(object_path_bstr)
        try:
            return_obj = IWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        try:
            return_obj2 = IWbemCallResult(return_obj2)
        except WindowsError:
            return_obj2 = None
        return return_obj, return_obj2

    def GetObject(self, object_path, flags, ctx):
        return_obj, return_obj2 = self.GetObjectWithResult(object_path, flags, ctx)
        if return_obj2:
            return_obj2.Release()
        return return_obj

    def GetObjectAsync(self, object_path, flags, ctx, response_handler):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(IWbemObjectSink))

        paramflags = ((_In_, 'strObjectPath'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_In_, 'pResponseHandler'),
                      )

        _GetObjectAsync = prototype(IWbemServices_GetObjectAsync_Idx,
                                    'GetObjectAsync',
                                    paramflags)
        _GetObjectAsync.errcheck = winapi.RAISE_NON_ZERO_ERR
        object_path_bstr = winapi.SysAllocString(object_path) if object_path is not None else None
        try:
            _GetObjectAsync(self.this,
                            object_path_bstr,
                            flags,
                            ctx.this if ctx else None,
                            response_handler.this if response_handler else None
                            )
        finally:
            if object_path_bstr is not None:
                winapi.SysFreeString(object_path_bstr)

    def PutClassWithResult(self, object_param, flags, ctx):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.POINTER(IWbemClassObject),
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'pObject'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_Out_, 'ppCallResult'),
                      )

        _PutClass = prototype(IWbemServices_PutClass_Idx,
                              'PutClass',
                              paramflags)
        _PutClass.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _PutClass(self.this,
                               object_param.this if object_param else None,
                               flags,
                               ctx.this if ctx else None
                               )
        try:
            return_obj = IWbemCallResult(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def PutClass(self, object_param, flags, ctx):
        return_obj = self.PutClassWithResult(object_param, flags, ctx)
        if return_obj:
            return_obj.Release()

    def PutClassAsync(self, object_param, flags, ctx, response_handler):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.POINTER(IWbemClassObject),
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(IWbemObjectSink))

        paramflags = ((_In_, 'pObject'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_In_, 'pResponseHandler'),
                      )

        _PutClassAsync = prototype(IWbemServices_PutClassAsync_Idx,
                                   'PutClassAsync',
                                   paramflags)
        _PutClassAsync.errcheck = winapi.RAISE_NON_ZERO_ERR
        _PutClassAsync(self.this,
                       object_param.this if object_param else None,
                       flags,
                       ctx.this if ctx else None,
                       response_handler.this if response_handler else None
                       )

    def DeleteClassWithResult(self, class_param, flags, ctx):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'strClass'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_Out_, 'ppCallResult'),
                      )

        _DeleteClass = prototype(IWbemServices_DeleteClass_Idx,
                                 'DeleteClass',
                                 paramflags)
        _DeleteClass.errcheck = winapi.RAISE_NON_ZERO_ERR
        class_param_bstr = winapi.SysAllocString(class_param) if class_param is not None else None
        try:
            return_obj = _DeleteClass(self.this,
                                      class_param_bstr,
                                      flags,
                                      ctx.this if ctx else None
                                      )
        finally:
            if class_param_bstr is not None:
                winapi.SysFreeString(class_param_bstr)
        try:
            return_obj = IWbemCallResult(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def DeleteClass(self, class_param, flags, ctx):
        return_obj = self.DeleteClassWithResult(class_param, flags, ctx)
        if return_obj:
            return_obj.Release()

    def DeleteClassAsync(self, class_param, flags, ctx, response_handler):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(IWbemObjectSink))

        paramflags = ((_In_, 'strClass'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_In_, 'pResponseHandler'),
                      )

        _DeleteClassAsync = prototype(IWbemServices_DeleteClassAsync_Idx,
                                      'DeleteClassAsync',
                                      paramflags)
        _DeleteClassAsync.errcheck = winapi.RAISE_NON_ZERO_ERR
        class_param_bstr = winapi.SysAllocString(class_param) if class_param is not None else None
        try:
            _DeleteClassAsync(self.this,
                              class_param_bstr,
                              flags,
                              ctx.this if ctx else None,
                              response_handler.this if response_handler else None
                              )
        finally:
            if class_param_bstr is not None:
                winapi.SysFreeString(class_param_bstr)

    def CreateClassEnum(self, superclass, flags, ctx):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'strSuperclass'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_Out_, 'ppEnum'),
                      )

        _CreateClassEnum = prototype(IWbemServices_CreateClassEnum_Idx,
                                     'CreateClassEnum',
                                     paramflags)
        _CreateClassEnum.errcheck = winapi.RAISE_NON_ZERO_ERR
        superclass_bstr = winapi.SysAllocString(superclass) if superclass is not None else None
        try:
            return_obj = _CreateClassEnum(self.this,
                                          superclass_bstr,
                                          flags,
                                          ctx.this if ctx else None
                                          )
        finally:
            if superclass_bstr is not None:
                winapi.SysFreeString(superclass_bstr)
        try:
            return_obj = IEnumWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def CreateClassEnumAsync(self, superclass, flags, ctx, response_handler):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(IWbemObjectSink))

        paramflags = ((_In_, 'strSuperclass'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_In_, 'pResponseHandler'),
                      )

        _CreateClassEnumAsync = prototype(IWbemServices_CreateClassEnumAsync_Idx,
                                          'CreateClassEnumAsync',
                                          paramflags)
        _CreateClassEnumAsync.errcheck = winapi.RAISE_NON_ZERO_ERR
        superclass_bstr = winapi.SysAllocString(superclass) if superclass is not None else None
        try:
            _CreateClassEnumAsync(self.this,
                                  superclass_bstr,
                                  flags,
                                  ctx.this if ctx else None,
                                  response_handler.this if response_handler else None
                                  )
        finally:
            if superclass_bstr is not None:
                winapi.SysFreeString(superclass_bstr)

    def PutInstanceWithResult(self, inst, flags, ctx):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.POINTER(IWbemClassObject),
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'pInst'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_Out_, 'ppCallResult'),
                      )

        _PutInstance = prototype(IWbemServices_PutInstance_Idx,
                                 'PutInstance',
                                 paramflags)
        _PutInstance.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_obj = _PutInstance(self.this,
                                  inst.this if inst else None,
                                  flags,
                                  ctx.this if ctx else None
                                  )
        try:
            return_obj = IWbemCallResult(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def PutInstance(self, inst, flags, ctx):
        return_obj = self.PutInstanceWithResult(inst, flags, ctx)
        if return_obj:
            return_obj.Release()

    def PutInstanceAsync(self, inst, flags, ctx, response_handler):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.POINTER(IWbemClassObject),
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(IWbemObjectSink))

        paramflags = ((_In_, 'pInst'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_In_, 'pResponseHandler'),
                      )

        _PutInstanceAsync = prototype(IWbemServices_PutInstanceAsync_Idx,
                                      'PutInstanceAsync',
                                      paramflags)
        _PutInstanceAsync.errcheck = winapi.RAISE_NON_ZERO_ERR
        _PutInstanceAsync(self.this,
                          inst.this if inst else None,
                          flags,
                          ctx.this if ctx else None,
                          response_handler.this if response_handler else None
                          )

    def DeleteInstanceWithResult(self, object_path, flags, ctx):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'strObjectPath'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_Out_, 'ppCallResult'),
                      )

        _DeleteInstance = prototype(IWbemServices_DeleteInstance_Idx,
                                    'DeleteInstance',
                                    paramflags)
        _DeleteInstance.errcheck = winapi.RAISE_NON_ZERO_ERR
        object_path_bstr = winapi.SysAllocString(object_path) if object_path is not None else None
        try:
            return_obj = _DeleteInstance(self.this,
                                         object_path_bstr,
                                         flags,
                                         ctx.this if ctx else None
                                         )
        finally:
            if object_path_bstr is not None:
                winapi.SysFreeString(object_path_bstr)
        try:
            return_obj = IWbemCallResult(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def DeleteInstance(self, object_path, flags, ctx):
        return_obj = self.DeleteInstanceWithResult(object_path, flags, ctx)
        if return_obj:
            return_obj.Release()

    def DeleteInstanceAsync(self, object_path, flags, ctx, response_handler):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(IWbemObjectSink))

        paramflags = ((_In_, 'strObjectPath'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_In_, 'pResponseHandler'),
                      )

        _DeleteInstanceAsync = prototype(IWbemServices_DeleteInstanceAsync_Idx,
                                         'DeleteInstanceAsync',
                                         paramflags)
        _DeleteInstanceAsync.errcheck = winapi.RAISE_NON_ZERO_ERR
        object_path_bstr = winapi.SysAllocString(object_path) if object_path is not None else None
        try:
            _DeleteInstanceAsync(self.this,
                                 object_path_bstr,
                                 flags,
                                 ctx.this if ctx else None,
                                 response_handler.this if response_handler else None
                                 )
        finally:
            if object_path_bstr is not None:
                winapi.SysFreeString(object_path_bstr)

    def CreateInstanceEnum(self, filter_param, flags, ctx):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'strFilter'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_Out_, 'ppEnum'),
                      )

        _CreateInstanceEnum = prototype(IWbemServices_CreateInstanceEnum_Idx,
                                        'CreateInstanceEnum',
                                        paramflags)
        _CreateInstanceEnum.errcheck = winapi.RAISE_NON_ZERO_ERR
        filter_param_bstr = winapi.SysAllocString(filter_param) if filter_param is not None else None
        try:
            return_obj = _CreateInstanceEnum(self.this,
                                             filter_param_bstr,
                                             flags,
                                             ctx.this if ctx else None
                                             )
        finally:
            if filter_param_bstr is not None:
                winapi.SysFreeString(filter_param_bstr)
        try:
            return_obj = IEnumWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def CreateInstanceEnumAsync(self, filter_param, flags, ctx, response_handler):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(IWbemObjectSink))

        paramflags = ((_In_, 'strFilter'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_In_, 'pResponseHandler'),
                      )

        _CreateInstanceEnumAsync = prototype(IWbemServices_CreateInstanceEnumAsync_Idx,
                                             'CreateInstanceEnumAsync',
                                             paramflags)
        _CreateInstanceEnumAsync.errcheck = winapi.RAISE_NON_ZERO_ERR
        filter_param_bstr = winapi.SysAllocString(filter_param) if filter_param is not None else None
        try:
            _CreateInstanceEnumAsync(self.this,
                                     filter_param_bstr,
                                     flags,
                                     ctx.this if ctx else None,
                                     response_handler.this if response_handler else None
                                     )
        finally:
            if filter_param_bstr is not None:
                winapi.SysFreeString(filter_param_bstr)

    def ExecQuery(self, query_language, query, flags, ctx):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'strQueryLanguage'),
                      (_In_, 'strQuery'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_Out_, 'ppEnum'),
                      )

        _ExecQuery = prototype(IWbemServices_ExecQuery_Idx,
                               'ExecQuery',
                               paramflags)
        _ExecQuery.errcheck = winapi.RAISE_NON_ZERO_ERR
        query_language_bstr = winapi.SysAllocString(query_language) if query_language is not None else None
        query_bstr = winapi.SysAllocString(query) if query is not None else None
        try:
            return_obj = _ExecQuery(self.this,
                                    query_language_bstr,
                                    query_bstr,
                                    flags,
                                    ctx.this if ctx else None
                                    )
        finally:
            if query_language_bstr is not None:
                winapi.SysFreeString(query_language_bstr)
            if query_bstr is not None:
                winapi.SysFreeString(query_bstr)
        try:
            return_obj = IEnumWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def ExecQueryAsync(self, query_language, query, flags, ctx, response_handler):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(IWbemObjectSink))

        paramflags = ((_In_, 'strQueryLanguage'),
                      (_In_, 'strQuery'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_In_, 'pResponseHandler'),
                      )

        _ExecQueryAsync = prototype(IWbemServices_ExecQueryAsync_Idx,
                                    'ExecQueryAsync',
                                    paramflags)
        _ExecQueryAsync.errcheck = winapi.RAISE_NON_ZERO_ERR
        query_language_bstr = winapi.SysAllocString(query_language) if query_language is not None else None
        query_bstr = winapi.SysAllocString(query) if query is not None else None
        try:
            _ExecQueryAsync(self.this,
                            query_language_bstr,
                            query_bstr,
                            flags,
                            ctx.this if ctx else None,
                            response_handler.this if response_handler else None
                            )
        finally:
            if query_language_bstr is not None:
                winapi.SysFreeString(query_language_bstr)
            if query_bstr is not None:
                winapi.SysFreeString(query_bstr)

    def ExecNotificationQuery(self, query_language, query, flags, ctx):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'strQueryLanguage'),
                      (_In_, 'strQuery'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_Out_, 'ppEnum'),
                      )

        _ExecNotificationQuery = prototype(IWbemServices_ExecNotificationQuery_Idx,
                                           'ExecNotificationQuery',
                                           paramflags)
        _ExecNotificationQuery.errcheck = winapi.RAISE_NON_ZERO_ERR
        query_language_bstr = winapi.SysAllocString(query_language) if query_language is not None else None
        query_bstr = winapi.SysAllocString(query) if query is not None else None
        try:
            return_obj = _ExecNotificationQuery(self.this,
                                                query_language_bstr,
                                                query_bstr,
                                                flags,
                                                ctx.this if ctx else None
                                                )
        finally:
            if query_language_bstr is not None:
                winapi.SysFreeString(query_language_bstr)
            if query_bstr is not None:
                winapi.SysFreeString(query_bstr)
        try:
            return_obj = IEnumWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj

    def ExecNotificationQueryAsync(self, query_language, query, flags, ctx, response_handler):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(IWbemObjectSink))

        paramflags = ((_In_, 'strQueryLanguage'),
                      (_In_, 'strQuery'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_In_, 'pResponseHandler'),
                      )

        _ExecNotificationQueryAsync = prototype(IWbemServices_ExecNotificationQueryAsync_Idx,
                                                'ExecNotificationQueryAsync',
                                                paramflags)
        _ExecNotificationQueryAsync.errcheck = winapi.RAISE_NON_ZERO_ERR
        query_language_bstr = winapi.SysAllocString(query_language) if query_language is not None else None
        query_bstr = winapi.SysAllocString(query) if query is not None else None
        try:
            _ExecNotificationQueryAsync(self.this,
                                        query_language_bstr,
                                        query_bstr,
                                        flags,
                                        ctx.this if ctx else None,
                                        response_handler.this if response_handler else None
                                        )
        finally:
            if query_language_bstr is not None:
                winapi.SysFreeString(query_language_bstr)
            if query_bstr is not None:
                winapi.SysFreeString(query_bstr)

    def ExecMethodWithResult(self, object_path, method_name, flags, ctx, in_params):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(IWbemClassObject),
                                       ctypes.POINTER(wintypes.LPVOID),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'strObjectPath'),
                      (_In_, 'strMethodName'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_In_, 'pInParams'),
                      (_Out_, 'ppOutParams'),
                      (_Out_, 'ppCallResult'),
                      )

        _ExecMethod = prototype(IWbemServices_ExecMethod_Idx,
                                'ExecMethod',
                                paramflags)
        _ExecMethod.errcheck = winapi.RAISE_NON_ZERO_ERR
        object_path_bstr = winapi.SysAllocString(object_path) if object_path is not None else None
        method_name_bstr = winapi.SysAllocString(method_name) if method_name is not None else None
        try:
            return_obj, return_obj2 = _ExecMethod(self.this,
                                                  object_path_bstr,
                                                  method_name_bstr,
                                                  flags,
                                                  ctx.this if ctx else None,
                                                  in_params.this if in_params else None
                                                  )
        finally:
            if object_path_bstr is not None:
                winapi.SysFreeString(object_path_bstr)
            if method_name_bstr is not None:
                winapi.SysFreeString(method_name_bstr)
        try:
            return_obj = IWbemClassObject(return_obj)
        except WindowsError:
            return_obj = None
        try:
            return_obj2 = IWbemCallResult(return_obj2)
        except WindowsError:
            return_obj2 = None
        return return_obj, return_obj2

    def ExecMethod(self, object_path, method_name, flags, ctx, in_params):
        return_obj, return_obj2 = self.ExecMethodWithResult(object_path, method_name, flags, ctx, in_params)
        if return_obj2:
            return_obj2.Release()
        return return_obj

    def ExecMethodAsync(self, object_path, method_name, flags, ctx, in_params, response_handler):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       BSTR,
                                       ctypes.c_long,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(IWbemClassObject),
                                       ctypes.POINTER(IWbemObjectSink))

        paramflags = ((_In_, 'strObjectPath'),
                      (_In_, 'strMethodName'),
                      (_In_, 'lFlags'),
                      (_In_, 'pCtx'),
                      (_In_, 'pInParams'),
                      (_In_, 'pResponseHandler'),
                      )

        _ExecMethodAsync = prototype(IWbemServices_ExecMethodAsync_Idx,
                                     'ExecMethodAsync',
                                     paramflags)
        _ExecMethodAsync.errcheck = winapi.RAISE_NON_ZERO_ERR
        object_path_bstr = winapi.SysAllocString(object_path) if object_path is not None else None
        method_name_bstr = winapi.SysAllocString(method_name) if method_name is not None else None
        try:
            _ExecMethodAsync(self.this,
                             object_path_bstr,
                             method_name_bstr,
                             flags,
                             ctx.this if ctx else None,
                             in_params.this if in_params else None,
                             response_handler.this if response_handler else None
                             )
        finally:
            if object_path_bstr is not None:
                winapi.SysFreeString(object_path_bstr)
            if method_name_bstr is not None:
                winapi.SysFreeString(method_name_bstr)


IWbemLocator_ConnectServer_Idx = 3


class IWbemLocator(com.IUnknown):

    def ConnectServer(self, network_resource, user, password, locale, security_flags, authority, ctx):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       BSTR,
                                       BSTR,
                                       BSTR,
                                       BSTR,
                                       ctypes.c_long,
                                       BSTR,
                                       ctypes.POINTER(IWbemContext),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'strNetworkResource'),
                      (_In_, 'strUser'),
                      (_In_, 'strPassword'),
                      (_In_, 'strLocale'),
                      (_In_, 'lSecurityFlags'),
                      (_In_, 'strAuthority'),
                      (_In_, 'pCtx'),
                      (_Out_, 'ppNamespace'),
                      )

        _ConnectServer = prototype(IWbemLocator_ConnectServer_Idx,
                                   'ConnectServer',
                                   paramflags)
        _ConnectServer.errcheck = winapi.RAISE_NON_ZERO_ERR
        network_resource_bstr = winapi.SysAllocString(network_resource) if network_resource is not None else None
        user_bstr = winapi.SysAllocString(user) if user is not None else None
        password_bstr = winapi.SysAllocString(password) if password is not None else None
        locale_bstr = winapi.SysAllocString(locale) if locale is not None else None
        authority_bstr = winapi.SysAllocString(authority) if authority is not None else None
        try:
            return_obj = _ConnectServer(self.this,
                                        network_resource_bstr,
                                        user_bstr,
                                        password_bstr,
                                        locale_bstr,
                                        security_flags,
                                        authority_bstr,
                                        ctx.this if ctx else None
                                        )
        finally:
            if network_resource_bstr is not None:
                winapi.SysFreeString(network_resource_bstr)
            if user_bstr is not None:
                winapi.SysFreeString(user_bstr)
            if password_bstr is not None:
                winapi.SysFreeString(password_bstr)
            if locale_bstr is not None:
                winapi.SysFreeString(locale_bstr)
            if authority_bstr is not None:
                winapi.SysFreeString(authority_bstr)
        try:
            return_obj = IWbemServices(return_obj)
        except WindowsError:
            return_obj = None
        return return_obj


class WMI(com.COM):
    '''
    Wrapper class for WMI interactions. WMI initialization / uninitialization are done via ctxmgr.

    N.B. If using this class, do not call init() and fini() directly. Only use through via ctxmgr
    '''
    def __init__(self,
                 net_resource,
                 user=None,
                 password=None,
                 locale=None,
                 sec_flags=0,
                 auth=None,
                 ctx=None,
                 cls_ctx=com.CLSCTX_INPROC_SERVER,
                 authn_svc=winapi.RPC_C_AUTHN_WINNT,
                 authz_svc=winapi.RPC_C_AUTHZ_NONE,
                 spn=None,
                 auth_info=None,
                 coinit=com.COINIT_MULTITHREADED,
                 authn_level=winapi.RPC_C_AUTHN_LEVEL_DEFAULT,
                 imp_level=winapi.RPC_C_IMP_LEVEL_IMPERSONATE,
                 auth_list=None,
                 capabilities=winapi.EOAC_NONE):
        self._net_resource = net_resource
        self._user = user
        self._password = password
        self._locale = locale
        self._sec_flags = sec_flags
        self._auth = auth
        self._ctx = ctx
        self._cls_ctx = cls_ctx
        self._authn_svc = authn_svc
        self._authz_svc = authz_svc
        self._spn = spn
        self._auth_info = auth_info
        self._authn_level = authn_level
        self._imp_level = imp_level
        self._auth_list = auth_list
        self._capabilities = capabilities
        self.service = None
        super(WMI, self).__init__(coinit)

    def __enter__(self):
        self.init()
        return self.service

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.fini()

    def init(self):
        super(WMI, self).init()
        self.initialize_security(None,
                                 -1,
                                 None,
                                 self._authn_level,
                                 self._imp_level,
                                 self._auth_list,
                                 self._capabilities)

        locator = self.create_instance(winapi.CLSID_WbemLocator,
                                       None,
                                       self._cls_ctx,
                                       winapi.IID_IWbemLocator,
                                       IWbemLocator)
        try:
            self.service = locator.ConnectServer(self._net_resource,
                                                 self._user,
                                                 self._password,
                                                 self._locale,
                                                 self._sec_flags,
                                                 self._auth,
                                                 self._ctx)
        finally:
            locator.Release()

        self.set_proxy_blanket(self.service.this,
                               self._authn_svc,
                               self._authz_svc,
                               self._spn,
                               self._authn_level,
                               self._imp_level,
                               self._auth_info,
                               self._capabilities)

    def fini(self):
        while self.service.Release():
            pass
        super(WMI, self).fini()
