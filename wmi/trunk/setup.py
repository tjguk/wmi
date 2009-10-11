from distutils.core import setup
import wmi

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

setup (
  name = "WMI",
  version = wmi.__version__,
  description = "Windows Management Instrumentation",
  author = "Tim Golden",
  author_email = "mail@timgolden.me.uk",
  url = "http://timgolden.me.uk/python/wmi.html",
  license = "http://www.opensource.org/licenses/mit-license.php",
  py_modules = ["wmi"],
  scripts = ["wmitest.py", "wmiweb.py", "wmitest.cmd", "wmitest.master.ini"],
  data_files = ["readme.txt", "changelog.txt"]
)

