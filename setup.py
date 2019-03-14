########################################################################
# Copyright 2017 FireEye Inc.
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

import platform
from setuptools import setup


system = platform.system()
if system != 'Windows':
    print('ERROR: Cannot install this module on {:s}. This package is compatible with Windows only'.format(system))
    exit(-1)


setup(name='cwmi',
      version='0.1.0',
      description='ctypes wrapper for WMI',
      author='Anthony Berglund',
      author_email='anthony.berglund@fireeye.com',
      url='https://github.com/fireeye/cwmi',
      download_url='https://github.com/fireeye/cwmi/archive/v0.0.1.tar.gz',
      platforms=['Windows'],
      license='Apache',
      packages=['cwmi'],
      scripts=['utils/i_to_m.py'],
      classifiers=['Environment :: Console',
                   'Operating System :: Microsoft :: Windows',
                   'License :: OSI Approved :: Apache Software License',
                   'Programming Language :: Python :: 3',
                   'Programming Language :: Python :: 2.7',
                   'Topic :: Software Development :: Libraries']
      )
