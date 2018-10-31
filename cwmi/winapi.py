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

from .wintype import HRESULT, BSTR


PSHORT = ctypes.POINTER(wintypes.SHORT)
PUSHORT = ctypes.POINTER(wintypes.USHORT)
LPLONG = ctypes.POINTER(wintypes.LONG)
PULONG = ctypes.POINTER(wintypes.ULONG)
CHAR = ctypes.c_char
PCHAR = ctypes.POINTER(CHAR)


_In_ = 1
_Out_ = 2

WBEM_FLAG_RETURN_IMMEDIATELY = 0x10
WBEM_FLAG_RETURN_WBEM_COMPLETE = 0
WBEM_FLAG_BIDIRECTIONAL = 0
WBEM_FLAG_FORWARD_ONLY = 0x20
WBEM_FLAG_NO_ERROR_OBJECT = 0x40
WBEM_FLAG_RETURN_ERROR_OBJECT = 0
WBEM_FLAG_SEND_STATUS = 0x80
WBEM_FLAG_DONT_SEND_STATUS = 0
WBEM_FLAG_ENSURE_LOCATABLE = 0x100
WBEM_FLAG_DIRECT_READ = 0x200
WBEM_FLAG_SEND_ONLY_SELECTED = 0
WBEM_RETURN_WHEN_COMPLETE = 0
WBEM_RETURN_IMMEDIATELY = 0x10
WBEM_MASK_RESERVED_FLAGS = 0x1f000
WBEM_FLAG_USE_AMENDED_QUALIFIERS = 0x20000
WBEM_FLAG_STRONG_VALIDATION = 0x10000

WBEM_NO_WAIT = 0
WBEM_INFINITE = 0xFFFFFFFF

WBEM_FLAG_ALWAYS = 0
WBEM_FLAG_ONLY_IF_TRUE = 0x1
WBEM_FLAG_ONLY_IF_FALSE = 0x2
WBEM_FLAG_ONLY_IF_IDENTICAL = 0x3
WBEM_MASK_PRIMARY_CONDITION = 0x3
WBEM_FLAG_KEYS_ONLY = 0x4
WBEM_FLAG_REFS_ONLY = 0x8
WBEM_FLAG_LOCAL_ONLY = 0x10
WBEM_FLAG_PROPAGATED_ONLY = 0x20
WBEM_FLAG_SYSTEM_ONLY = 0x30
WBEM_FLAG_NONSYSTEM_ONLY = 0x40
WBEM_MASK_CONDITION_ORIGIN = 0x70
WBEM_FLAG_CLASS_OVERRIDES_ONLY = 0x100
WBEM_FLAG_CLASS_LOCAL_AND_OVERRIDES = 0x200
WBEM_MASK_CLASS_CONDITION = 0x300

WBEM_NO_ERROR = 0
WBEM_S_NO_ERROR = 0
WBEM_S_SAME = 0
WBEM_S_FALSE = 1
WBEM_S_ALREADY_EXISTS = 0x40001
WBEM_S_RESET_TO_DEFAULT = 0x40002
WBEM_S_DIFFERENT = 0x40003
WBEM_S_TIMEDOUT = 0x40004
WBEM_S_NO_MORE_DATA = 0x40005
WBEM_S_OPERATION_CANCELLED = 0x40006
WBEM_S_PENDING = 0x40007
WBEM_S_DUPLICATE_OBJECTS = 0x40008
WBEM_S_ACCESS_DENIED = 0x40009
WBEM_S_PARTIAL_RESULTS = 0x40010
WBEM_S_SOURCE_NOT_AVAILABLE = 0x40017
WBEM_E_FAILED = 0x80041001
WBEM_E_NOT_FOUND = 0x80041002
WBEM_E_ACCESS_DENIED = 0x80041003
WBEM_E_PROVIDER_FAILURE = 0x80041004
WBEM_E_TYPE_MISMATCH = 0x80041005
WBEM_E_OUT_OF_MEMORY = 0x80041006
WBEM_E_INVALID_CONTEXT = 0x80041007
WBEM_E_INVALID_PARAMETER = 0x80041008
WBEM_E_NOT_AVAILABLE = 0x80041009
WBEM_E_CRITICAL_ERROR = 0x8004100a
WBEM_E_INVALID_STREAM = 0x8004100b
WBEM_E_NOT_SUPPORTED = 0x8004100c
WBEM_E_INVALID_SUPERCLASS = 0x8004100d
WBEM_E_INVALID_NAMESPACE = 0x8004100e
WBEM_E_INVALID_OBJECT = 0x8004100f
WBEM_E_INVALID_CLASS = 0x80041010
WBEM_E_PROVIDER_NOT_FOUND = 0x80041011
WBEM_E_INVALID_PROVIDER_REGISTRATION = 0x80041012
WBEM_E_PROVIDER_LOAD_FAILURE = 0x80041013
WBEM_E_INITIALIZATION_FAILURE = 0x80041014
WBEM_E_TRANSPORT_FAILURE = 0x80041015
WBEM_E_INVALID_OPERATION = 0x80041016
WBEM_E_INVALID_QUERY = 0x80041017
WBEM_E_INVALID_QUERY_TYPE = 0x80041018
WBEM_E_ALREADY_EXISTS = 0x80041019
WBEM_E_OVERRIDE_NOT_ALLOWED = 0x8004101a
WBEM_E_PROPAGATED_QUALIFIER = 0x8004101b
WBEM_E_PROPAGATED_PROPERTY = 0x8004101c
WBEM_E_UNEXPECTED = 0x8004101d
WBEM_E_ILLEGAL_OPERATION = 0x8004101e
WBEM_E_CANNOT_BE_KEY = 0x8004101f
WBEM_E_INCOMPLETE_CLASS = 0x80041020
WBEM_E_INVALID_SYNTAX = 0x80041021
WBEM_E_NONDECORATED_OBJECT = 0x80041022
WBEM_E_READ_ONLY = 0x80041023
WBEM_E_PROVIDER_NOT_CAPABLE = 0x80041024
WBEM_E_CLASS_HAS_CHILDREN = 0x80041025
WBEM_E_CLASS_HAS_INSTANCES = 0x80041026
WBEM_E_QUERY_NOT_IMPLEMENTED = 0x80041027
WBEM_E_ILLEGAL_NULL = 0x80041028
WBEM_E_INVALID_QUALIFIER_TYPE = 0x80041029
WBEM_E_INVALID_PROPERTY_TYPE = 0x8004102a
WBEM_E_VALUE_OUT_OF_RANGE = 0x8004102b
WBEM_E_CANNOT_BE_SINGLETON = 0x8004102c
WBEM_E_INVALID_CIM_TYPE = 0x8004102d
WBEM_E_INVALID_METHOD = 0x8004102e
WBEM_E_INVALID_METHOD_PARAMETERS = 0x8004102f
WBEM_E_SYSTEM_PROPERTY = 0x80041030
WBEM_E_INVALID_PROPERTY = 0x80041031
WBEM_E_CALL_CANCELLED = 0x80041032
WBEM_E_SHUTTING_DOWN = 0x80041033
WBEM_E_PROPAGATED_METHOD = 0x80041034
WBEM_E_UNSUPPORTED_PARAMETER = 0x80041035
WBEM_E_MISSING_PARAMETER_ID = 0x80041036
WBEM_E_INVALID_PARAMETER_ID = 0x80041037
WBEM_E_NONCONSECUTIVE_PARAMETER_IDS = 0x80041038
WBEM_E_PARAMETER_ID_ON_RETVAL = 0x80041039
WBEM_E_INVALID_OBJECT_PATH = 0x8004103a
WBEM_E_OUT_OF_DISK_SPACE = 0x8004103b
WBEM_E_BUFFER_TOO_SMALL = 0x8004103c
WBEM_E_UNSUPPORTED_PUT_EXTENSION = 0x8004103d
WBEM_E_UNKNOWN_OBJECT_TYPE = 0x8004103e
WBEM_E_UNKNOWN_PACKET_TYPE = 0x8004103f
WBEM_E_MARSHAL_VERSION_MISMATCH = 0x80041040
WBEM_E_MARSHAL_INVALID_SIGNATURE = 0x80041041
WBEM_E_INVALID_QUALIFIER = 0x80041042
WBEM_E_INVALID_DUPLICATE_PARAMETER = 0x80041043
WBEM_E_TOO_MUCH_DATA = 0x80041044
WBEM_E_SERVER_TOO_BUSY = 0x80041045
WBEM_E_INVALID_FLAVOR = 0x80041046
WBEM_E_CIRCULAR_REFERENCE = 0x80041047
WBEM_E_UNSUPPORTED_CLASS_UPDATE = 0x80041048
WBEM_E_CANNOT_CHANGE_KEY_INHERITANCE = 0x80041049
WBEM_E_CANNOT_CHANGE_INDEX_INHERITANCE = 0x80041050
WBEM_E_TOO_MANY_PROPERTIES = 0x80041051
WBEM_E_UPDATE_TYPE_MISMATCH = 0x80041052
WBEM_E_UPDATE_OVERRIDE_NOT_ALLOWED = 0x80041053
WBEM_E_UPDATE_PROPAGATED_METHOD = 0x80041054
WBEM_E_METHOD_NOT_IMPLEMENTED = 0x80041055
WBEM_E_METHOD_DISABLED = 0x80041056
WBEM_E_REFRESHER_BUSY = 0x80041057
WBEM_E_UNPARSABLE_QUERY = 0x80041058
WBEM_E_NOT_EVENT_CLASS = 0x80041059
WBEM_E_MISSING_GROUP_WITHIN = 0x8004105a
WBEM_E_MISSING_AGGREGATION_LIST = 0x8004105b
WBEM_E_PROPERTY_NOT_AN_OBJECT = 0x8004105c
WBEM_E_AGGREGATING_BY_OBJECT = 0x8004105d
WBEM_E_UNINTERPRETABLE_PROVIDER_QUERY = 0x8004105f
WBEM_E_BACKUP_RESTORE_WINMGMT_RUNNING = 0x80041060
WBEM_E_QUEUE_OVERFLOW = 0x80041061
WBEM_E_PRIVILEGE_NOT_HELD = 0x80041062
WBEM_E_INVALID_OPERATOR = 0x80041063
WBEM_E_LOCAL_CREDENTIALS = 0x80041064
WBEM_E_CANNOT_BE_ABSTRACT = 0x80041065
WBEM_E_AMENDED_OBJECT = 0x80041066
WBEM_E_CLIENT_TOO_SLOW = 0x80041067
WBEM_E_NULL_SECURITY_DESCRIPTOR = 0x80041068
WBEM_E_TIMED_OUT = 0x80041069
WBEM_E_INVALID_ASSOCIATION = 0x8004106a
WBEM_E_AMBIGUOUS_OPERATION = 0x8004106b
WBEM_E_QUOTA_VIOLATION = 0x8004106c
WBEM_E_RESERVED_001 = 0x8004106d
WBEM_E_RESERVED_002 = 0x8004106e
WBEM_E_UNSUPPORTED_LOCALE = 0x8004106f
WBEM_E_HANDLE_OUT_OF_DATE = 0x80041070
WBEM_E_CONNECTION_FAILED = 0x80041071
WBEM_E_INVALID_HANDLE_REQUEST = 0x80041072
WBEM_E_PROPERTY_NAME_TOO_WIDE = 0x80041073
WBEM_E_CLASS_NAME_TOO_WIDE = 0x80041074
WBEM_E_METHOD_NAME_TOO_WIDE = 0x80041075
WBEM_E_QUALIFIER_NAME_TOO_WIDE = 0x80041076
WBEM_E_RERUN_COMMAND = 0x80041077
WBEM_E_DATABASE_VER_MISMATCH = 0x80041078
WBEM_E_VETO_DELETE = 0x80041079
WBEM_E_VETO_PUT = 0x8004107a
WBEM_E_INVALID_LOCALE = 0x80041080
WBEM_E_PROVIDER_SUSPENDED = 0x80041081
WBEM_E_SYNCHRONIZATION_REQUIRED = 0x80041082
WBEM_E_NO_SCHEMA = 0x80041083
WBEM_E_PROVIDER_ALREADY_REGISTERED = 0x80041084
WBEM_E_PROVIDER_NOT_REGISTERED = 0x80041085
WBEM_E_FATAL_TRANSPORT_ERROR = 0x80041086
WBEM_E_ENCRYPTED_CONNECTION_REQUIRED = 0x80041087
WBEM_E_PROVIDER_TIMED_OUT = 0x80041088
WBEM_E_NO_KEY = 0x80041089
WBEM_E_PROVIDER_DISABLED = 0x8004108a
WBEMESS_E_REGISTRATION_TOO_BROAD = 0x80042001
WBEMESS_E_REGISTRATION_TOO_PRECISE = 0x80042002
WBEMESS_E_AUTHZ_NOT_PRIVILEGED = 0x80042003
WBEMMOF_E_EXPECTED_QUALIFIER_NAME = 0x80044001
WBEMMOF_E_EXPECTED_SEMI = 0x80044002
WBEMMOF_E_EXPECTED_OPEN_BRACE = 0x80044003
WBEMMOF_E_EXPECTED_CLOSE_BRACE = 0x80044004
WBEMMOF_E_EXPECTED_CLOSE_BRACKET = 0x80044005
WBEMMOF_E_EXPECTED_CLOSE_PAREN = 0x80044006
WBEMMOF_E_ILLEGAL_CONSTANT_VALUE = 0x80044007
WBEMMOF_E_EXPECTED_TYPE_IDENTIFIER = 0x80044008
WBEMMOF_E_EXPECTED_OPEN_PAREN = 0x80044009
WBEMMOF_E_UNRECOGNIZED_TOKEN = 0x8004400a
WBEMMOF_E_UNRECOGNIZED_TYPE = 0x8004400b
WBEMMOF_E_EXPECTED_PROPERTY_NAME = 0x8004400c
WBEMMOF_E_TYPEDEF_NOT_SUPPORTED = 0x8004400d
WBEMMOF_E_UNEXPECTED_ALIAS = 0x8004400e
WBEMMOF_E_UNEXPECTED_ARRAY_INIT = 0x8004400f
WBEMMOF_E_INVALID_AMENDMENT_SYNTAX = 0x80044010
WBEMMOF_E_INVALID_DUPLICATE_AMENDMENT = 0x80044011
WBEMMOF_E_INVALID_PRAGMA = 0x80044012
WBEMMOF_E_INVALID_NAMESPACE_SYNTAX = 0x80044013
WBEMMOF_E_EXPECTED_CLASS_NAME = 0x80044014
WBEMMOF_E_TYPE_MISMATCH = 0x80044015
WBEMMOF_E_EXPECTED_ALIAS_NAME = 0x80044016
WBEMMOF_E_INVALID_CLASS_DECLARATION = 0x80044017
WBEMMOF_E_INVALID_INSTANCE_DECLARATION = 0x80044018
WBEMMOF_E_EXPECTED_DOLLAR = 0x80044019
WBEMMOF_E_CIMTYPE_QUALIFIER = 0x8004401a
WBEMMOF_E_DUPLICATE_PROPERTY = 0x8004401b
WBEMMOF_E_INVALID_NAMESPACE_SPECIFICATION = 0x8004401c
WBEMMOF_E_OUT_OF_RANGE = 0x8004401d
WBEMMOF_E_INVALID_FILE = 0x8004401e
WBEMMOF_E_ALIASES_IN_EMBEDDED = 0x8004401f
WBEMMOF_E_NULL_ARRAY_ELEM = 0x80044020
WBEMMOF_E_DUPLICATE_QUALIFIER = 0x80044021
WBEMMOF_E_EXPECTED_FLAVOR_TYPE = 0x80044022
WBEMMOF_E_INCOMPATIBLE_FLAVOR_TYPES = 0x80044023
WBEMMOF_E_MULTIPLE_ALIASES = 0x80044024
WBEMMOF_E_INCOMPATIBLE_FLAVOR_TYPES2 = 0x80044025
WBEMMOF_E_NO_ARRAYS_RETURNED = 0x80044026
WBEMMOF_E_MUST_BE_IN_OR_OUT = 0x80044027
WBEMMOF_E_INVALID_FLAGS_SYNTAX = 0x80044028
WBEMMOF_E_EXPECTED_BRACE_OR_BAD_TYPE = 0x80044029
WBEMMOF_E_UNSUPPORTED_CIMV22_QUAL_VALUE = 0x8004402a
WBEMMOF_E_UNSUPPORTED_CIMV22_DATA_TYPE = 0x8004402b
WBEMMOF_E_INVALID_DELETEINSTANCE_SYNTAX = 0x8004402c
WBEMMOF_E_INVALID_QUALIFIER_SYNTAX = 0x8004402d
WBEMMOF_E_QUALIFIER_USED_OUTSIDE_SCOPE = 0x8004402e
WBEMMOF_E_ERROR_CREATING_TEMP_FILE = 0x8004402f
WBEMMOF_E_ERROR_INVALID_INCLUDE_FILE = 0x80044030
WBEMMOF_E_INVALID_DELETECLASS_SYNTAX = 0x80044031

EOAC_NONE = 0
EOAC_MUTUAL_AUTH = 0x1
EOAC_STATIC_CLOAKING = 0x20
EOAC_DYNAMIC_CLOAKING = 0x40
EOAC_ANY_AUTHORITY = 0x80
EOAC_MAKE_FULLSIC = 0x100
EOAC_DEFAULT = 0x800
EOAC_SECURE_REFS = 0x2
EOAC_ACCESS_CONTROL = 0x4
EOAC_APPID = 0x8
EOAC_DYNAMIC = 0x10
EOAC_REQUIRE_FULLSIC = 0x200
EOAC_AUTO_IMPERSONATE = 0x400
EOAC_NO_CUSTOM_MARSHAL = 0x2000
EOAC_DISABLE_AAA = 0x1000


RPC_C_IMP_LEVEL_DEFAULT = 0
RPC_C_IMP_LEVEL_ANONYMOUS = 1
RPC_C_IMP_LEVEL_IDENTIFY = 2
RPC_C_IMP_LEVEL_IMPERSONATE = 3
RPC_C_IMP_LEVEL_DELEGATE = 4

RPC_C_AUTHN_LEVEL_DEFAULT = 0
RPC_C_AUTHN_LEVEL_NONE = 1
RPC_C_AUTHN_LEEL_CONNECT = 2
RPC_C_AUTHN_LEVEL_CALL = 3
RPC_C_AUTHN_LEVEL_PKT = 4
RPC_C_AUTHN_LEVEL_PKT_INTEGRITY = 5
RPC_C_AUTHN_LEVEL_PKT_PRIVACY = 6

RPC_AUTH_IDENTITY_HANDLE = wintypes.HANDLE
RPC_AUTHZ_HANDLE = wintypes.HANDLE

RPC_C_AUTHN_NONE = 0
RPC_C_AUTHN_DCE_PRIVATE = 1
RPC_C_AUTHN_DCE_PUBLIC = 2
RPC_C_AUTHN_DEC_PUBLIC = 4
RPC_C_AUTHN_GSS_NEGOTIATE = 9
RPC_C_AUTHN_WINNT = 10
RPC_C_AUTHN_GSS_SCHANNEL = 14
RPC_C_AUTHN_GSS_KERBEROS = 16
RPC_C_AUTHN_DPA = 17
RPC_C_AUTHN_MSN = 18
RPC_C_AUTHN_DIGEST = 21
RPC_C_AUTHN_KERNEL = 20

RPC_C_AUTHZ_NONE = 0
RPC_C_AUTHZ_NAME = 1
RPC_C_AUTHZ_DCE = 2
RPC_C_AUTHZ_DEFAULT = 0xffffffff

VT_EMPTY = 0
VT_NULL = 1
VT_I2 = 2
VT_I4 = 3
VT_R4 = 4
VT_R8 = 5
VT_CY = 6
VT_DATE = 7
VT_BSTR = 8
VT_DISPATCH = 9
VT_ERROR = 10
VT_BOOL = 11
VT_VARIANT = 12
VT_UNKNOWN = 13
VT_DECIMAL = 14
VT_I1 = 16
VT_UI1 = 17
VT_UI2 = 18
VT_UI4 = 19
VT_I8 = 20
VT_UI8 = 21
VT_INT = 22
VT_UINT = 23
VT_VOID = 24
VT_HRESULT = 25
VT_PTR = 26
VT_SAFEARRAY = 27
VT_CARRAY = 28
VT_USERDEFINED = 29
VT_LPSTR = 30
VT_LPWSTR = 31
VT_RECORD = 36
VT_INT_PTR = 37
VT_UINT_PTR = 38
VT_FILETIME = 64
VT_BLOB = 65
VT_STREAM = 66
VT_STORAGE = 67
VT_STREAMED_OBJECT = 68
VT_STORED_OBJECT = 69
VT_BLOB_OBJECT = 70
VT_CF = 71
VT_CLSID = 72
VT_VERSIONED_STREAM = 73
VT_BSTR_BLOB = 0xfff
VT_VECTOR = 0x1000
VT_ARRAY = 0x2000
VT_BYREF = 0x4000
VT_RESERVED = 0x8000
VT_ILLEGAL = 0xffff
VT_ILLEGALMASKED = 0xfff
VT_TYPEMASK = 0xfff


VT_GENERIC = -1


VARTYPE = ctypes.c_ushort


class SOLE_AUTHENTICATION_SERVICE(ctypes.Structure):
    _fields_ = [
        ('dwAuthnSvc', wintypes.DWORD),
        ('dwAuthzSvc', wintypes.DWORD),
        ('pPrincipalName', wintypes.OLESTR),
        ('hr', HRESULT)
    ]


class GUID(ctypes.Structure):
    _fields_ = [
        ('Data1', wintypes.DWORD),
        ('Data2', wintypes.WORD),
        ('Data3', wintypes.WORD),
        ('Data4', wintypes.BYTE * 8),
    ]


class SECURITY_DESCRIPTOR(ctypes.Structure):
    _fields_ = [('Revision', ctypes.c_byte),
                ('Sbz1', ctypes.c_byte),
                ('Control', ctypes.c_uint16),
                ('Owner', ctypes.c_void_p),
                ('Group', ctypes.c_void_p),
                ('Sacl', ctypes.c_void_p),
                ('Dacl', ctypes.c_void_p),
                ]


class SAFEARRAYBOUND(ctypes.Structure):
    _fields_ = [('cElements', wintypes.ULONG),
                ('lLbound', wintypes.LONG)]


class SAFEARRAY(ctypes.Structure):
    _fields_ = [('cDims', wintypes.USHORT),
                ('fFeatures', wintypes.USHORT),
                ('cbElements', wintypes.ULONG),
                ('cLocks', wintypes.ULONG),
                ('pvData', wintypes.LPVOID),
                ('rgsabound', ctypes.POINTER(SAFEARRAYBOUND))]


class DECIMAL_DUMMYSTRUCTNAME(ctypes.Structure):
    _fields_ = [('scale', wintypes.BYTE),
                ('sign', wintypes.BYTE)]


class DECIMAL_DUMMYSTRUCTNAME2(ctypes.Structure):
    _fields_ = [('Lo32', wintypes.ULONG),
                ('Mid32', wintypes.ULONG)]


class DECIMAL_DUMMYUNIONNAME(ctypes.Union):
    _fields_ = [('DUMMYSTRUCTNAME', DECIMAL_DUMMYSTRUCTNAME),
                ('signscale', wintypes.USHORT)]


class DECIMAL_DUMMYUNIONNAME2(ctypes.Union):
    _fields_ = [('DUMMYSTRUCTNAME2', DECIMAL_DUMMYSTRUCTNAME2),
                ('Lo64', ctypes.c_ulonglong)]


class DECIMAL(ctypes.Structure):
    _fields_ = [('DUMMYUNIONNAME', DECIMAL_DUMMYUNIONNAME),
                ('Hi32', wintypes.ULONG),
                ('DUMMYUNIONNAME2', DECIMAL_DUMMYUNIONNAME2)]


class VARIANT__tagBRECORD(ctypes.Structure):
    _fields_ = [("pvRecord", wintypes.LPVOID),
                ("pRecInfo", wintypes.LPVOID)]  # IRecordInfo * pRecInfo;


class VARIANT__VARIANT_NAME_3(ctypes.Union):
    _fields_ = [("llVal", wintypes.USHORT),
                ("lVal", wintypes.USHORT),
                ('llVal', ctypes.c_longlong),
                ('lVal', wintypes.LONG),
                ('bVal', wintypes.BYTE),
                ('iVal', wintypes.SHORT),
                ('fltVal', wintypes.FLOAT),
                ('dblVal', wintypes.DOUBLE),
                ('boolVal', wintypes.VARIANT_BOOL),
                ('bool', wintypes.VARIANT_BOOL),
                ('scode', wintypes.LONG),
                ('cyVal', wintypes.LPVOID),  # struct CY
                ('date', wintypes.DOUBLE),
                ('bstrVal', BSTR),
                ('punkVal', wintypes.LPVOID),  # IUnknown *
                ('pdispVal', wintypes.LPVOID),  # IDispatch *
                ('parray', ctypes.POINTER(SAFEARRAY)),
                ('pbVal', ctypes.POINTER(wintypes.BYTE)),
                ('piVal', PSHORT),
                ('plVal', ctypes.POINTER(wintypes.LONG)),
                ('pllVal', ctypes.POINTER(ctypes.c_longlong)),
                ('pfltVal', ctypes.POINTER(wintypes.FLOAT)),
                ('pdblVal', ctypes.POINTER(wintypes.DOUBLE)),
                ('pboolVal', ctypes.POINTER(wintypes.VARIANT_BOOL)),
                ('pbool', ctypes.POINTER(wintypes.VARIANT_BOOL)),
                ('pscode', LPLONG),
                ('pcyVal', wintypes.LPVOID),  # CY *
                ('pdate', ctypes.POINTER(wintypes.DOUBLE)),
                ('pbstrVal', ctypes.POINTER(BSTR)),
                ('ppunkVal', ctypes.POINTER(wintypes.LPVOID)),  # IUnknown **
                ('ppdispVal', ctypes.POINTER(wintypes.LPVOID)),  # IDispatch **
                ('pparray', ctypes.POINTER(ctypes.POINTER(SAFEARRAY))),
                ('pvarVal', wintypes.LPVOID),  # VARIANT*
                ('byref', wintypes.LPVOID),
                ('cVal', CHAR),
                ('uiVal', wintypes.USHORT),
                ('ulVal', wintypes.ULONG),
                ('ullVal', ctypes.c_ulonglong),
                ('intVal', wintypes.INT),
                ('uintVal', wintypes.UINT),
                ('pdecVal', ctypes.POINTER(DECIMAL)),
                ('pcVal', PCHAR),
                ('puiVal', PUSHORT),
                ('pulVal', PULONG),
                ('pullVal', ctypes.POINTER(ctypes.c_ulonglong)),
                ('pintVal', ctypes.POINTER(wintypes.INT)),
                ('puintVal', ctypes.POINTER(wintypes.UINT)),
                ("__VARIANT_NAME_4", VARIANT__tagBRECORD)]


class VARIANT__tagVARIANT(ctypes.Structure):
    _fields_ = [("vt", wintypes.USHORT),
                ("wReserved1", wintypes.WORD),
                ("wReserved2", wintypes.WORD),
                ("wReserved3", wintypes.WORD),
                ("__VARIANT_NAME_3", VARIANT__VARIANT_NAME_3)]


class VARIANT__VARIANT_NAME_1(ctypes.Union):
    _fields_ = [("__VARIANT_NAME_2", VARIANT__tagVARIANT),
                ("decVal", DECIMAL)]


class VARIANT(ctypes.Structure):
    _fields_ = [('__VARIANT_NAME_1', VARIANT__VARIANT_NAME_1)]


def V_VAR3(X):
    return X.__VARIANT_NAME_1.__VARIANT_NAME_2.__VARIANT_NAME_3


def V_VT(X):
    return X.__VARIANT_NAME_1.__VARIANT_NAME_2.vt


V_VAR = V_VAR3


def V_EMPTY():
    return ''


def V_NULL():
    return 'NULL'


def V_BSTR(var):
    return convert_bstr_to_str(V_VAR3(var).bstrVal)


def V_BOOL(var):
    return V_VAR(var).boolVal


def V_DATE(var):
    return V_VAR(var).date


def V_I2(var):
    return V_VAR(var).lVal


def V_I4(var):
    return V_VAR(var).lVal


def V_I8(var):
    return V_VAR(var).llVal


def V_R4(var):
    return V_VAR(var).lVal


def V_R8(var):
    return V_VAR(var).llVal


def V_GENERIC(var):
    return V_VAR(var).byref


def V_TO_STR(var):
    ret = ''
    vt = V_VT(var)
    if vt == VT_EMPTY:
        ret = V_EMPTY()
    elif vt == VT_NULL:
        ret = V_NULL()
    elif vt == VT_DATE:
        ret = str(V_DATE(var))
    elif vt == VT_BSTR:
        ret = V_BSTR(var)
    elif vt == VT_BOOL:
        ret = str(V_BOOL(var))
    elif vt == VT_I2:
        ret = str(V_I2(var))
    elif vt == VT_I4:
        ret = str(V_I4(var))
    elif vt == VT_I8:
        ret = str(V_I8(var))
    elif vt == VT_R4:
        ret = str(V_R4(var))
    elif vt == VT_R8:
        ret = str(V_R8(var))
    else:
        try:
            ret = str(V_GENERIC(var))
        except:
            pass
    return ret


def V_TO_VT_DICT(var):
    ret = {}
    vt = V_VT(var)
    if vt == VT_EMPTY:
        ret['value'] = V_EMPTY()
    elif vt == VT_NULL:
        ret['value'] = V_NULL()
    elif vt == VT_DATE:
        ret['value'] = V_DATE(var)
    elif vt == VT_BSTR:
        ret['value'] = V_BSTR(var)
    elif vt == VT_BOOL:
        ret['value'] = V_BOOL(var)
    elif vt == VT_I2:
        ret['value'] = V_I2(var)
    elif vt == VT_I4:
        ret['value'] = V_I4(var)
    elif vt == VT_I8:
        ret['value'] = V_I8(var)
    elif vt == VT_R4:
        ret['value'] = V_R4(var)
    elif vt == VT_R8:
        ret['value'] = V_R8(var)
    else:
        try:
            ret['value'] = V_GENERIC(var)
        except:
            pass
    ret['type'] = vt
    return ret


def V_TO_TYPE(var):
    ret = None
    vt = V_VT(var)
    if vt == VT_NULL:
        ret = V_NULL()
    elif vt == VT_DATE:
        ret = V_DATE(var)
    elif vt == VT_BSTR:
        ret = V_BSTR(var)
    elif vt == VT_BOOL:
        ret = V_BOOL(var)
    elif vt == VT_I2:
        ret = V_I2(var)
    elif vt == VT_I4:
        ret = V_I4(var)
    elif vt == VT_I8:
        ret = V_I8(var)
    elif vt == VT_R4:
        ret = V_R4(var)
    elif vt == VT_R8:
        ret = V_R8(var)
    else:
        try:
            ret = V_GENERIC(var)
        except:
            pass
    return ret


def SET_VT(var, vt):
    var.__VARIANT_NAME_1.__VARIANT_NAME_2.vt = vt


VARIANTARG = VARIANT


CLSID_WbemLocator = GUID(0x4590f811, 0x1d3a, 0x11d0, (0x89, 0x1f, 0x00, 0xaa, 0x00, 0x4b, 0x2e, 0x24))
IID_IWbemLocator = GUID(0xdc12a687, 0x737f, 0x11cf, (0x88, 0x4d, 0x00, 0xaa, 0x00, 0x4b, 0x2e, 0x24))


ole32 = ctypes.windll.ole32
oleaut32 = ctypes.windll.oleaut32


def RAISE_NON_ZERO_ERR(result, func, args):
    if result != 0:
        raise ctypes.WinError(result)
    return args


def convert_bstr_to_str(bstr):
    length = SysStringLen(bstr)
    converted_string = str(bstr[:length])
    SysFreeString(bstr)
    return converted_string


def CoInitializeEx(reserved, co_init):
    prototype = ctypes.WINFUNCTYPE(
        HRESULT,
        wintypes.LPVOID,
        wintypes.DWORD
    )

    paramflags = (
        (_In_, 'pvReserved'),
        (_In_, 'dwCoInit')
    )

    _CoInitializeEx = prototype(('CoInitializeEx', ole32), paramflags)
    _CoInitializeEx.errcheck = RAISE_NON_ZERO_ERR

    return _CoInitializeEx(reserved, co_init)


def CoUninitialize():
    prototype = ctypes.WINFUNCTYPE(
        None
    )

    _CoUninitialize = prototype(('CoUninitialize', ole32))
    _CoUninitialize()


def CoCreateInstance(clsid, unk, ctx, iid):
    prototype = ctypes.WINFUNCTYPE(
        HRESULT,
        ctypes.POINTER(GUID),
        wintypes.LPVOID,
        wintypes.DWORD,
        ctypes.POINTER(GUID),
        ctypes.POINTER(wintypes.LPVOID)
    )

    paramflags = (
        (_In_, 'rclsid'),
        (_In_, 'pUnkOuter'),
        (_In_, 'dwClsContext'),
        (_In_, 'riid'),
        (_Out_, 'ppv')
    )

    _CoCreateInstance = prototype(('CoCreateInstance', ole32), paramflags)
    _CoCreateInstance.errcheck = RAISE_NON_ZERO_ERR
    return _CoCreateInstance(clsid, unk, ctx, iid)


def CoInitializeSecurity(sec_desc,
                         c_auth_svc,
                         as_auth_svc,
                         reserved1,
                         auth_level,
                         imp_level,
                         auth_list,
                         capibilities,
                         reserved3):
    prototype = ctypes.WINFUNCTYPE(
        HRESULT,
        ctypes.POINTER(SECURITY_DESCRIPTOR),
        wintypes.LONG,
        ctypes.POINTER(SOLE_AUTHENTICATION_SERVICE),
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.LPVOID
    )

    paramflags = (
        (_In_, 'pSecDesc'),
        (_In_, 'cAuthSvc'),
        (_In_, 'asAuthSvc'),
        (_In_, 'pReserved1'),
        (_In_, 'dwAuthnLevel'),
        (_In_, 'dwImpLevel'),
        (_In_, 'pAuthList'),
        (_In_, 'dwCapabilities'),
        (_In_, 'pReserved3')
    )

    _CoInitializeSecurity = prototype(('CoInitializeSecurity', ole32), paramflags)
    _CoInitializeSecurity.errcheck = RAISE_NON_ZERO_ERR

    return _CoInitializeSecurity(sec_desc,
                                 c_auth_svc,
                                 as_auth_svc,
                                 reserved1,
                                 auth_level,
                                 imp_level,
                                 auth_list,
                                 capibilities,
                                 reserved3)


def CoSetProxyBlanket(proxy,
                      authn_svc,
                      authz_svc,
                      server_p_name,
                      authn_level,
                      imp_level,
                      auth_info,
                      capabilities):
    prototype = ctypes.WINFUNCTYPE(
        HRESULT,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.OLESTR,
        wintypes.DWORD,
        wintypes.DWORD,
        RPC_AUTH_IDENTITY_HANDLE,
        wintypes.DWORD
    )

    paramflags = (
        (_In_, 'pProxy'),
        (_In_, 'dwAuthnSvc'),
        (_In_, 'dwAuthzSvc'),
        (_In_, 'pServerPrincName'),
        (_In_, 'dwAuthnLevel'),
        (_In_, 'dwImpLevel'),
        (_In_, 'pAuthInfo'),
        (_In_, 'dwCapabilities')
    )

    _CoSetProxyBlanket = prototype(('CoSetProxyBlanket', ole32), paramflags)
    _CoSetProxyBlanket.errcheck = RAISE_NON_ZERO_ERR
    return _CoSetProxyBlanket(proxy,
                              authn_svc,
                              authz_svc,
                              server_p_name,
                              authn_level,
                              imp_level,
                              auth_info,
                              capabilities)


def SysAllocString(wstr):
    prototype = ctypes.WINFUNCTYPE(
        BSTR,
        wintypes.LPOLESTR
    )

    paramflags = (
        (_In_, 'psz'),
    )

    _SysAllocString = prototype(('SysAllocString', oleaut32), paramflags)
    return _SysAllocString(wstr)


def SysFreeString(bstr):
    prototype = ctypes.WINFUNCTYPE(
        None,
        BSTR
    )

    paramflags = (
        (_In_, 'bstrString'),
    )

    _SysFreeString = prototype(('SysFreeString', oleaut32), paramflags)
    _SysFreeString(bstr)


def SysStringLen(bstr):
    prototype = ctypes.WINFUNCTYPE(
        wintypes.UINT,
        BSTR
    )

    paramflags = (
        (_In_, 'pbstr'),
    )

    _SysStringLen = prototype(('SysStringLen', oleaut32), paramflags)
    return _SysStringLen(bstr)


def VariantInit(var):
    prototype = ctypes.WINFUNCTYPE(
        None,
        ctypes.POINTER(VARIANTARG)
    )

    paramflags = (
        (_In_, 'pvarg'),
    )

    _VariantInit = prototype(('VariantInit', oleaut32), paramflags)
    _VariantInit(var)


def VariantClear(var):
    prototype = ctypes.WINFUNCTYPE(
        HRESULT,
        ctypes.POINTER(VARIANTARG)
    )

    paramflags = (
        (_In_, 'pvarg'),
    )

    _VariantClear = prototype(('VariantClear', oleaut32), paramflags)
    _VariantClear.errcheck = RAISE_NON_ZERO_ERR
    return _VariantClear(var)


def SafeArrayCreate(vt, dims, bound):
    prototype = ctypes.WINFUNCTYPE(
        ctypes.POINTER(SAFEARRAY),
        VARTYPE,
        wintypes.UINT,
        ctypes.POINTER(SAFEARRAYBOUND)
    )

    paramflags = (
        (_In_, 'vt'),
        (_In_, 'cDims'),
        (_In_, 'rgsabound'),
    )

    _SafeArrayCreate = prototype(('SafeArrayCreate', oleaut32), paramflags)
    return _SafeArrayCreate(vt, dims, bound)


def SafeArrayDestroy(sa):
    prototype = ctypes.WINFUNCTYPE(
        HRESULT,
        ctypes.POINTER(SAFEARRAY)
    )

    paramflags = (
        (_In_, 'psa'),
    )

    _SafeArrayDestroy = prototype(('SafeArrayDestroy', oleaut32), paramflags)
    _SafeArrayDestroy.errcheck = RAISE_NON_ZERO_ERR
    return _SafeArrayDestroy(sa)


def SafeArrayLock(sa):
    prototype = ctypes.WINFUNCTYPE(
        HRESULT,
        ctypes.POINTER(SAFEARRAY)
    )

    paramflags = (
        (_In_, 'psa'),
    )

    _SafeArrayLock = prototype(('SafeArrayLock', oleaut32), paramflags)
    _SafeArrayLock.errcheck = RAISE_NON_ZERO_ERR
    return _SafeArrayLock(sa)


def SafeArrayUnlock(sa):
    prototype = ctypes.WINFUNCTYPE(
        HRESULT,
        ctypes.POINTER(SAFEARRAY)
    )

    paramflags = (
        (_In_, 'psa'),
    )

    _SafeArrayUnlock = prototype(('SafeArrayUnlock', oleaut32), paramflags)
    _SafeArrayUnlock.errcheck = RAISE_NON_ZERO_ERR
    return _SafeArrayUnlock(sa)


def SafeArrayAccessData(sa):
    prototype = ctypes.WINFUNCTYPE(
        HRESULT,
        ctypes.POINTER(SAFEARRAY),
        ctypes.POINTER(wintypes.LPVOID)
    )

    paramflags = (
        (_In_, 'psa'),
        (_Out_, 'ppvData', ctypes.pointer(wintypes.LPVOID(None))),
    )

    _SafeArrayAccessData = prototype(('SafeArrayAccessData', oleaut32), paramflags)
    _SafeArrayAccessData.errcheck = RAISE_NON_ZERO_ERR
    return_obj = _SafeArrayAccessData(sa)
    return return_obj.contents


def SafeArrayUnaccessData(sa):
    prototype = ctypes.WINFUNCTYPE(
        HRESULT,
        ctypes.POINTER(SAFEARRAY)
    )

    paramflags = (
        (_In_, 'psa'),
    )

    _SafeArrayUnaccessData = prototype(('SafeArrayUnaccessData', oleaut32), paramflags)
    _SafeArrayUnaccessData.errcheck = RAISE_NON_ZERO_ERR
    return _SafeArrayUnaccessData(sa)
