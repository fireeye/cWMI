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

import unittest

from cwmi import com
from cwmi import wmi
from cwmi import winapi


class TestCOM(unittest.TestCase):

    def setUp(self):
        self.instance = com.COM()
        self.instance.init()

        self.instance.initialize_security()

    def tearDown(self):
        self.instance.fini()

    def test_iunknown(self):
        """
        Tests the IUnknown interface

        :return: None
        """
        obj = self.instance.create_instance(winapi.CLSID_WbemLocator,
                                            None,
                                            com.CLSCTX_INPROC_SERVER,
                                            winapi.IID_IWbemLocator,
                                            com.IUnknown)
        obj.QueryInterface(winapi.IID_IWbemLocator)
        count = obj.AddRef()
        assert count == 3
        count = obj.Release()
        assert count == 2
        count = obj.Release()
        assert count == 1
        count = obj.Release()
        assert count == 0

    def test_com_create_instance(self):
        """
        Tests the ability to create a COM object instance

        :return: None
        """

        obj = self.instance.create_instance(winapi.CLSID_WbemLocator,
                                            None,
                                            com.CLSCTX_INPROC_SERVER,
                                            winapi.IID_IWbemLocator,
                                            wmi.IWbemLocator)
        assert obj.Release() == 0

    def test_com_set_proxy_blanket(self):
        """
        Tests the ability to set security on object

        :return: None
        """

        locator = self.instance.create_instance(winapi.CLSID_WbemLocator,
                                                None,
                                                com.CLSCTX_INPROC_SERVER,
                                                winapi.IID_IWbemLocator,
                                                wmi.IWbemLocator)
        assert locator.this

        svc = locator.ConnectServer('root\\cimv2', None, None, None, 0, None, None)
        assert svc.this
        self.instance.set_proxy_blanket(svc.this)
        assert locator.Release() == 0
        assert svc.Release() == 0


if __name__ == '__main__':
    unittest.main()
