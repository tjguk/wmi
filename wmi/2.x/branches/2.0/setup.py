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

setup (
  name = "WMI",
  version = "1.3.2",
  description = "Windows Management Instrumentation",
  author = "Tim Golden",
  author_email = "mail@timgolden.me.uk",
  url = "http://timgolden.me.uk/python/wmi.html",
  license = "http://www.opensource.org/licenses/mit-license.php",
  py_modules = ["wmi"]
)

