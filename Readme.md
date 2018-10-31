# cWMI
This project is a lightweight wrapper for interacting with WMI using python/ctypes without any external dependencies. It allows a lower level of interaction with WMI, but requires a greater knowledge of WMI internals. There are several helpers included that will come in handy for those who don't need to do anything low level, or just want a quick start.

This project was not intended to be a replacement for the existing (and very good) `wmi` module, but as a pure python alternative without additional dependencies.

## Usage

To use this module `import cwmi` and either use the helper functions, or the WMI APIs directly.

Easy way, using helper functions.

```
import cwmi


def list_processes():
    processes = cwmi.query('root\\cimv2', 'SELECT * FROM Win32_Process')
    print('NAME\t\tPROCESS ID')
    print('=' * 80)
    for process, values in processes.items():
        print('{:s}\t\t{:d}'.format(values['Name'], values['ProcessId']))
```
        
Not as easy way, directly interacting with WMI APIs.

```
import cwmi

def list_processes():
    with cwmi.WMI('root\\cimv2') as svc:
        with svc.ExecQuery('WQL', 'SELECT * FROM Win32_Process', 0, None) as enum:

            print('NAME\t\tPROCESS ID')
            while True:
                try:
                    with enum.Next(cwmi.WBEM_INFINITE) as obj:
                        obj.BeginEnumeration(0)
                        proc_data = {}
                        while True:
                            try:
                                prop_name, var, _, _ = obj.Next(0)
                                proc_data[prop_name] = cwmi.V_TO_TYPE(var)
                            except WindowsError:
                                break

                        obj.EndEnumeration()

                        print('{:s}\t\t{:d}'.format(proc_data['Name'], 
                                                    proc_data['ProcessId']))

                except WindowsError:
                    break
```

For more examples see [examples](examples).

Microsoft's [WMI documentation](https://docs.microsoft.com/en-us/windows/desktop/wmisdk/using-wmi) will be helpful if directly interacting with the APIs. 

Some things to note are that while the documentation specifies input pointer parameters to some methods, this will not be necessary when using the wrapper APIs. The wrappers also handle python `str` to `BSTR` conversion and resource management.
