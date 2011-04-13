.. WinSys documentation master file, created by sphinx-quickstart on Fri Oct 31 15:35:06 2008.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

WMI - Windows Management Instrumentation
========================================

What is it?
-----------

Windows Management Instrumentation (WMI) is Microsoft's implementation of
Web-Based Enterprise Management (WBEM), an industry initiative to provide
a Common Information Model (CIM) for pretty much any information about a
computer system.

The Python WMI module is a lightweight wrapper on top of the pywin32
extensions, and hides some of the messy plumbing needed to get Python to
talk to the WMI API. It's pure Python and has been tested against all
versions of Python from 2.4 to 3.2. It should work with any recent
version of pywin32.


Where do I get it?
------------------

* **Subversion**: http://svn.timgolden.me.uk/wmi/trunk/
* **Windows installer** and **zip**: http://pypi.python.org/pypi/WMI/#downloads

* **Older Versions**: http://timgolden.me.uk/python/downloads


How do I install it?
--------------------

When all's said and done, it's just a module. But for those who like setup programs::

  python setup.py install

Or download the Windows installer and double-click.


How do I use it?
----------------

Have a look at the :doc:`tutorial` or the :doc:`cookbook`. As a quick
taster, try this, to find all Automatic services which are not running
and offer the option to restart each one::

  import wmi

  c = wmi.WMI ()
  for s in c.Win32_Service (StartMode="Auto", State="Stopped"):
    if raw_input ("Restart %s? " % s.Caption).upper () == "Y":
      s.StartService ()

What's Changed?
---------------

See the :doc:`changes` document

Copyright & License?
--------------------

* Copyright Tim Golden <mail@timgolden.me.uk> 2003 - 2010

* Licensed under the (GPL-compatible) MIT License:
  http://www.opensource.org/licenses/mit-license.php

Prerequisites
-------------

If you're running a recent Python (2.4+) on a recent Windows (2k, 2k3, XP)
and you have Mark Hammond's win32 extensions installed, you're probably
up-and-running already. Otherwise...

Windows
~~~~~~~
If you're running Win9x / NT4 you'll need to get WMI support
from Microsoft. Microsoft URLs change quite often, so I suggest you
do this: http://www.google.com/search?q=wmi+downloads

Python
~~~~~~
http://www.python.org/ (just in case you didn't know)

pywin32 (was win32all)
~~~~~~~~~~~~~~~~~~~~~~
http://sourceforge.net/projects/pywin32/files/

Specifically, builds 154/155 fixed a problem which affected the WMI
moniker construction. You can still work without this fix, but some
more complex monikers will fail. (The current build is 214 so you're
probably ok unless you have some very stringent backwards-compatible
requirement).

makepy
~~~~~~
(NB my own experience over several systems is that this
step isn't necessary. However, if you have problems...)
You may have to compile makepy support for some typelibs. The following
are reported to be significant:

* Microsoft WMI Scripting Library
* WMI ADSI Extension Type Library
* WMICntl Type Library

If you've not done this before, start the PythonWin environment, select
Tools > Com Makepy utility from the menu, select the library by name, and
click [OK].
