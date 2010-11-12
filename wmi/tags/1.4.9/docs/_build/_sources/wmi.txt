:mod:`wmi` -- Windows Management Instrumentation
================================================

..  automodule:: wmi

Exceptions
----------

All COM-related exceptions are wrapped in :exc:`x_wmi` or one of its
subclasses. Therefore you can safely trap :exc:`x_wmi` as a root
exception.

From v1.4 onwards exceptions now contain the underlying COM error
(if any) as their :attr:`x_wmi.com_error` attribute::

  import wmi

  try:
    c = wmi.WMI ("non-existent-machine")
  except wmi.x_wmi, x: # Py3+ except wmi.x_wmi as x:
    print "Exception number", x.com_error.hresult

..  autoexception:: x_wmi
..  autoexception:: x_wmi_invalid_query
..  autoexception:: x_wmi_timed_out
..  autoexception:: x_wmi_no_namespace
..  autoexception:: x_access_denied
..  autoexception:: x_wmi_authentication
..  autoexception:: x_wmi_uninitialised_thread

Support Classes & Functions
---------------------------

These classes and functions are provided internally to the
module to support various operations internally. Certain of
them, in particular :func:`from_time` and :func:`to_time`
could well be used externally.

..  autoclass:: SelfDeprecatingDict
..  autoclass:: ProvideConstants
..  autofunction:: handle_com_error
..  autofunction:: from_time
..  autofunction:: to_time
..  autofunction:: _set

Implementation
--------------

These functions & classes are part of the underlying implementation and you
are not expected to call upon them explicitly. That said, it can be useful
to know what's used and what you might use to assist in implementing some
functionality not currently exposed.

..  autoclass:: _wmi_method
    :members:

    ..  automethod:: __init__
    ..  automethod:: __call__

..  autoclass:: _wmi_object
    :members:

..  autoclass:: _wmi_class
    :members:

..  autoclass:: _wmi_namespace
    :members:

Main Entry Points
-----------------

..  autofunction:: WMI
..  autofunction:: connect_server
..  autofunction:: Registry
