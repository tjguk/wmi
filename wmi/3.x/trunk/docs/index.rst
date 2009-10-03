.. WMI documentation master file, created by sphinx-quickstart.py on Mon Mar 17 13:47:19 2008.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

:mod:`wmi` -- Windows Management Instrumentation
================================================

..  automodule:: wmi
    :synopsis: Windows Management Instrumentation
..  moduleauthor:: Tim Golden <mail@timgolden.me.uk>

Functions
---------
..  autofunction:: WMI
..  autofunction:: handle_com_error
..  autofunction:: to_time
..  autofunction:: from_time

Classes
-------
..  autoclass:: _wmi_namespace
    :members:
    
..  autoclass:: _wmi_object
    :members:
    
..  autoclass:: _wmi_class
    :members:

..  autoclass:: _wmi_method
    :members:

..  autoclass:: _wmi_event
    :members:
    
..  autoclass:: _wmi_result
    :members:

Exceptions
----------
..  autoexception:: x_wmi
..  autoexception:: x_wmi_invalid_query
..  autoexception:: x_wmi_timed_out
..  autoexception:: x_wmi_no_namespace

References
----------
..  seealso::

    `Python WMI Tutorial <http://timgolden.me.uk/python/wmi-tutorial.html>`_
      Tutorial covering basic and advanced usage
      
    `Python WMI Cookbook <http://timgolden.me.uk/python/wmi_cookbook.html>`_
      Cookbook with a range of examples

To Do
-----
* Asynchronous queries
* Rework for more recent Python versions
* Confirm that unicode error messages now work
