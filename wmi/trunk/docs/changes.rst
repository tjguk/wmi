Changes
=======

1.4
---

* Tests - WMI now comes with a unittest-based set of tests which have been run (and passed) 
  against all versions of Python from 2.4 to 3.2 using pywin32 214.

* Exception changes - x_wmi and its subclasses now store the underlying COM error as .com_error instead of returning
  an equivalent string. This should help code which needs to know the exact error which
  occurred and should also make it easier for non-ASCII systems which were experiencing
  difficulties when the code attempted to manipulate non-decoded byte strings.

* Specific trapping of uninitialised threads - like all COM-based code, WMI needs to be initialised if run
  inside a thread. If this isn't done, the error message is slightly obscure. The connection maker now
  traps this specific case and returns a specific exception: :exc:`x_wmi_uninitialised_thread`.

* More specific checks for invalid connection combinations - certain combinations of identification
  and authentication are invalid when connection. A specific exception is now raised for these:
  :exc:`x_wmi_authentication`.

* keys - each :class:`_wmi_object` now has a keys attribute, inferred from the underlying
  WMI class definition, which is the list of attributes which uniquely define an instance of that class.

* associated_classes - each :class:`_wmi_object` has an associated_classes attribute which is
  a dictionary mapping the names of the other WMI classes which can be associated to this one to their
  :class:`_wmi_class` objects. This is most immediately of use in the wmiweb browser (qv) but can
  be used by client code.

