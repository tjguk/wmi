..  module:: wmi

Changes
=======

1.4
---

* WMI now needs Python 2.4+. It wouldn't be impossible to continue to support older
  versions but it's increasingly onerous and 2.4 is now 5 years old.

* Tests - WMI now comes with a unittest-based set of tests which have been run (and passed)
  against all versions of Python from 2.4 to 3.2 using pywin32 214.

* Exception changes - :exc:`x_wmi` and its subclasses now store the underlying COM error as
  :attr:`x_wmi.com_error` instead of returning
  an equivalent string. This should help code which needs to know the exact error which
  occurred and should also make it easier for non-ASCII systems which were experiencing
  difficulties when the code attempted to manipulate non-decoded byte strings.

* Specific trapping of uninitialised threads - like all COM-based code, WMI needs to be initialised if run
  inside a thread. If this isn't done, the error message is slightly obscure. The connection maker now
  traps this specific case and returns a specific exception: :exc:`x_wmi_uninitialised_thread`.

* More specific checks for invalid connection combinations - certain combinations of identification
  and authentication are invalid when connection. A specific exception is now raised for these:
  :exc:`x_wmi_authentication`.

* `keys` - each :class:`_wmi_object` now has a :attr:`_wmi_object.keys` attribute, inferred from the underlying
  WMI class definition, which is the list of attributes which uniquely define an instance of that class.

* `associated_classes` - each :class:`_wmi_object` has an :attr:`_wmi_object.associated_classes` attribute which is
  a dictionary mapping the names of the other WMI classes which can be associated to this one to their
  :class:`_wmi_class` objects. This is most immediately of use in the wmiweb browser (qv) but can
  be used by client code.

* By default, the :func:`WMI` connection function -- the one you call most often -- no longer looks to
  find the subclasses of a namespace. This makes for much faster startup times: altho' it was
  always possible to pass `find_classes=False` this was little known, and you now have to pass
  `find_classes=True` to get this functionality, or use the :attr:`_wmi_namespace.classes` attribute which
  is now calculated lazily, so things like IPython's attribute lookup still work.

* wmiweb.py - the installation now ships with a small but functional web-based WMI browser.
  It uses only the stdlib WSGI server and makes it easy to explore any of the namespaces
  on the local or a remote machine.

* Removed the rarely-used Win32 autoprefix: previously, if you tried for a class
  called `Process`, the module would try it again under `Win32_Process` if it failed
  first time round. This has now been removed to avoid the magic and because I certainly
  never use it, and I'm not aware of anyone who did.

* Impersonation & Authentication levels now supported when connecting to a remote
  server with specific credentials. Thanks to Matt Mercer for sample code.

* Documentation is now Sphinx-based.

* Association classes no longer try to treat all their attributes as WMI classes.
  (Thanks to Miroslav Ježek for the bug report)

* Setting a property's value now works again
  (Thanks to John Holcomb for the bug report and the detective work)

1.3
---

* Support for IPython's getAttribute protocol (patch supplied by Igor Dvorkin)

* Allow positional parameters for method calls. Previously, parameters had to
  be passed by keyword and failure to do so resulted in an obscure error message.

* Allow extrinsic events to use the same watcher API as intrinsic ones. Under the
  covers these behave slightly differently. Intrinsic events now default to modification
  rather than creation.

* Remove the restriction where an instantiated class didn't know its own namespace.
