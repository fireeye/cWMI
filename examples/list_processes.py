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

import cwmi


def list_processes():
    processes = cwmi.query('root\\cimv2', 'SELECT * FROM Win32_Process')
    print('NAME\t\tPROCESS ID')
    print('=' * 80)
    for process, values in processes.items():
        print('{:s}\t\t{:d}'.format(values['Name'], values['ProcessId']))


if __name__ == '__main__':
    list_processes()
