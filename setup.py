import os
import re
from distutils.core import setup

classifiers = [
    'Development Status :: 5 - Production/Stable',
    'Environment :: Win32 (MS Windows)',
    'Intended Audience :: Developers',
    'Intended Audience :: System Administrators',
    'License :: MIT',
    'Natural Language :: English',
    'Operating System :: Microsoft :: Windows :: Windows 95/98/2000',
    'Topic :: System :: Systems Administration'
]

HERE = os.path.dirname(__file__)
#
# Slight hack to avoid importing the module
# when setup is run
#
wmi_filepath = os.path.join(HERE, "wmi.py")
version = re.search(r'__version__ = "([a-z0-9\.]+)"', open(wmi_filepath).read()).group(1)

setup (
    name = "WMI",
    version = version,
    description = "Windows Management Instrumentation",
    author = "Tim Golden",
    author_email = "mail@timgolden.me.uk",
    url = "http://timgolden.me.uk/python/wmi.html",
    license = "http://www.opensource.org/licenses/mit-license.php",
    py_modules = ["wmi"],
    scripts = ["wmitest.py", "wmiweb.py", "wmitest.cmd", "wmitest.master.ini"],
    data_files = ["readme.rst"],
    long_description = open(os.path.join(HERE, "readme.rst")).read()
)

