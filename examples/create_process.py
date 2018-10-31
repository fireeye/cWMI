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

import argparse

import cwmi


def create_process(path):
    out = cwmi.call_method('root\\cimv2', 'Win32_Process', 'Create', {'CommandLine': path})
    ret = out['ReturnValue']
    if not ret:
        print('Process created successfully with process id of {:d}'.format(out['ProcessId']))
    else:
        print('Process not created successfully, ERROR is {:d}'.format(ret))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--path', help='Full path to executable', required=True)
    parsed_args = parser.parse_args()
    create_process(parsed_args.path)
