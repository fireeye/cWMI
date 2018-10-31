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

from . import winapi
from .wintype import HRESULT


_In_ = 1
_Out_ = 2

CLSCTX_INPROC_SERVER = 0x1
CLSCTX_INPROC_HANDLER = 0x2
CLSCTX_LOCAL_SERVER = 0x4
CLSCTX_INPROC_SERVER16 = 0x8
CLSCTX_REMOTE_SERVER = 0x10
CLSCTX_INPROC_HANDLER16 = 0x20
CLSCTX_RESERVED1 = 0x40
CLSCTX_RESERVED2 = 0x80
CLSCTX_RESERVED3 = 0x100
CLSCTX_RESERVED4 = 0x200
CLSCTX_NO_CODE_DOWNLOAD = 0x400
CLSCTX_RESERVED5 = 0x800
CLSCTX_NO_CUSTOM_MARSHAL = 0x1000
CLSCTX_ENABLE_CODE_DOWNLOAD = 0x2000
CLSCTX_NO_FAILURE_LOG = 0x4000
CLSCTX_DISABLE_AAA = 0x8000
CLSCTX_ENABLE_AAA = 0x10000
CLSCTX_FROM_DEFAULT_CONTEXT = 0x20000
CLSCTX_ACTIVATE_32_BIT_SERVER = 0x40000
CLSCTX_ACTIVATE_64_BIT_SERVER = 0x80000
CLSCTX_ENABLE_CLOAKING = 0x100000
CLSCTX_APPCONTAINER = 0x400000
CLSCTX_ACTIVATE_AAA_AS_IU = 0x800000
CLSCTX_PS_DLL = 0x80000000

CLSCTX_SERVER = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER

COINIT_MULTITHREADED = 0
COINIT_APARTMENTTHREADED = 2

IUnknown_QueryInterface_Idx = 0
IUnknown_AddRef_Idx = 1
IUnknown_Release_Idx = 2


class ComClassInstance(ctypes.Structure):
    '''
    COM instance class. This class represents an instantiated class. Class construction will fail a null pointer is
    passed into constructor.
    '''
    def __init__(self, this):
        if not this:
            raise WindowsError('Could not construct {}'.format(self.__class__))
        self.this = ctypes.cast(this, ctypes.POINTER(self.__class__))
        ctypes.Structure.__init__(self)

    def __getattr__(self, item):
        if item == 'this':
            return self.this
        else:
            raise NameError(item)

    def __eq__(self, other):
        if isinstance(other, ComClassInstance):
            ret = self.this == other.this
        else:
            ret = self.this == other
        return ret

    def __ne__(self, other):
        if isinstance(other, ComClassInstance):
            ret = self.this != other.this
        else:
            ret = self.this != other
        return ret

    def __lt__(self, other):
        if isinstance(other, ComClassInstance):
            ret = self.this < other.this
        else:
            ret = self.this < other
        return ret

    def __le__(self, other):
        if isinstance(other, ComClassInstance):
            ret = self.this <= other.this
        else:
            ret = self.this <= other
        return ret

    def __gt__(self, other):
        if isinstance(other, ComClassInstance):
            ret = self.this > other.this
        else:
            ret = self.this > other
        return ret

    def __ge__(self, other):
        if isinstance(other, ComClassInstance):
            ret = self.this >= other.this
        else:
            ret = self.this >= other
        return ret


class IUnknown(ComClassInstance):

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.AddRef() > 0:
            while self.Release():
                pass

    def QueryInterface(self, riid):
        prototype = ctypes.WINFUNCTYPE(HRESULT,
                                       ctypes.POINTER(winapi.GUID),
                                       ctypes.POINTER(wintypes.LPVOID))

        paramflags = ((_In_, 'riid'),
                      (_Out_, 'ppvObject', ctypes.pointer(wintypes.LPVOID(None)))
                      )

        _QueryInterface = prototype(IUnknown_QueryInterface_Idx,
                                    'QueryInterface',
                                    paramflags)
        _QueryInterface.errcheck = winapi.RAISE_NON_ZERO_ERR
        return_ptr = _QueryInterface(self.this,
                                     ctypes.byref(riid))
        return IUnknown(return_ptr.contents)

    def AddRef(self):
        prototype = ctypes.WINFUNCTYPE(wintypes.LONG)
        paramflags = ()
        _AddRef = prototype(IUnknown_AddRef_Idx,
                            'AddRef',
                            paramflags)
        return _AddRef(self.this)

    def Release(self):
        prototype = ctypes.WINFUNCTYPE(wintypes.LONG)
        paramflags = ()
        _Release = prototype(IUnknown_Release_Idx,
                             'Release',
                             paramflags)
        return _Release(self.this)


class COM(object):
    '''
    COM wrapper class. Wraps COM initialization / uninitialization via ctxmgr.

    N.B. If using this class, do not call init() and fini() directly. Only use through via ctxmgr
    '''
    def __init__(self, coinit=COINIT_MULTITHREADED):
        self._coinit = coinit
        self._initialized = False

    def __enter__(self):
        self.init()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.fini()

    def init(self):
        winapi.CoInitializeEx(None, self._coinit)
        self._initialized = True

    def fini(self):
        if self._initialized is True:
            winapi.CoUninitialize()
        self._initialized = False

    def create_instance(self, clsid, outer, ctx, iid, obj_type=None):
        obj = winapi.CoCreateInstance(clsid,
                                      outer,
                                      ctx,
                                      iid)
        if obj_type:
            obj = obj_type(obj)
        return obj

    def initialize_security(self,
                            desc=None,
                            auth_svc=-1,
                            as_auth_svc=None,
                            auth_level=winapi.RPC_C_AUTHN_LEVEL_DEFAULT,
                            imp_level=winapi.RPC_C_IMP_LEVEL_IMPERSONATE,
                            auth_list=None,
                            capabilities=winapi.EOAC_NONE):
        winapi.CoInitializeSecurity(desc,
                                    auth_svc,
                                    as_auth_svc,
                                    None,
                                    auth_level,
                                    imp_level,
                                    auth_list,
                                    capabilities,
                                    None)

    def set_proxy_blanket(self,
                          proxy,
                          auth_svc=winapi.RPC_C_AUTHN_WINNT,
                          authz_svc=winapi.RPC_C_AUTHZ_NONE,
                          name=None,
                          auth_level=winapi.RPC_C_AUTHN_LEVEL_CALL,
                          imp_level=winapi.RPC_C_IMP_LEVEL_IMPERSONATE,
                          auth_info=None,
                          capabilities=winapi.EOAC_NONE):
        winapi.CoSetProxyBlanket(proxy,
                                 auth_svc,
                                 authz_svc,
                                 name,
                                 auth_level,
                                 imp_level,
                                 auth_info,
                                 capabilities)
