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

import sys
import argparse

import cwmi


def control_service(service_name, action):
    svc_name = 'Win32_Service.Name="{:s}"'.format(service_name)
    if action == 'START':
        method_name = 'StartService'
    elif action == 'STOP':
        method_name = 'StopService'
    elif action == 'PAUSE':
        method_name = 'PauseService'
    else:
        method_name = 'ResumeService'

    print('Attempting to {:s} {:s} service'.format(action, service_name))
    out = cwmi.call_method('root\\cimv2', svc_name, method_name, None)
    ret = out['ReturnValue']
    if not ret:
        print('Service state changed successfully')
    else:
        print('Service state not changed successfully, ERROR is {:d}'.format(ret))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--service', help='Service short name', required=True)
    parser.add_argument('--action', help='Action [START|STOP|PAUSE|RESUME]', required=True)
    parsed_args = parser.parse_args()
    if parsed_args.action != 'START' and \
       parsed_args.action != 'STOP' and \
       parsed_args.action != 'PAUSE' and \
       parsed_args.action != 'RESUME':
        print('Action must be one of [START|STOP|PAUSE|RESUME]')
        sys.exit(-1)
    control_service(parsed_args.service, parsed_args.action)
