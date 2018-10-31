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
import unittest

import cwmi


class TestCWMI(unittest.TestCase):

    def test_get_object_info(self):
        """
        Tests calling get_object_info function

        :return: None
        """
        info = cwmi.get_object_info('root\\cimv2', 'Win32_BaseBoard.Tag="Base Board"', ['Caption'])
        assert info

    def test_get_all_object_info(self):
        """
        Tests calling get_all_object_info function

        :return: None
        """
        info = cwmi.get_all_object_info('root\\cimv2', 'Win32_BaseBoard.Tag="Base Board"')
        assert info

    def test_get_object_names(self):
        """
        Tests calling get_object_names function

        :return: None
        """
        names = cwmi.get_object_names('root\\cimv2', 'Win32_BaseBoard.Tag="Base Board"')
        assert names

    def test_query(self):
        """
        Tests calling query function

        :return: None
        """
        info = cwmi.query('root\\cimv2', 'SELECT * FROM Win32_BaseBoard')
        assert info
        assert len(info) == 1

    def test_does_object_exist(self):
        """
        Tests calling does_object_exist function

        :return: None
        """
        assert cwmi.does_object_exist('root\\cimv2', 'Win32_BaseBoard.Tag="Base Board"')
        assert not cwmi.does_object_exist('root\\cimv2', 'Win32_BaseBoard.Tag="Fake Base Board"')

    def test_get_method_names(self):
        """
        Tests calling get_method_names function

        :return: None
        """
        names = cwmi.get_method_names('root\\cimv2', 'Win32_Process')
        assert names

    def test_get_method_info(self):
        """
        Tests calling get_method_info function

        :return: None
        """
        info = cwmi.get_method_info('root\\cimv2', 'Win32_Process')
        assert info

    def test_call_method(self):
        """
        Tests calling call_method function

        :return: None
        """
        out = cwmi.call_method('root\\cimv2',
                               'Win32_Process',
                               'Create',
                               {'CommandLine': 'c:\\windows\\system32\\notepad.exe'})
        assert out
        assert not out['ReturnValue']
        pid = out['ProcessId']
        handle = ctypes.windll.kernel32.OpenProcess(1, False, pid)
        ctypes.windll.kernel32.TerminateProcess(handle, 0)
        ctypes.windll.kernel32.CloseHandle(handle)

if __name__ == '__main__':
    unittest.main()
