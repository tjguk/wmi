"""wmi - a lightweight Python wrapper around Microsoft's WMI interface

Windows Management Instrumentation (WMI) is Microsoft's answer to
 the DMTF's Common Information Model. It allows you to query just
 about any conceivable piece of information from any computer which
 is running the necessary agent and over which have you the
 necessary authority.

The implementation is by means of COM/DCOM and most of the examples
 assume you're running one of Microsoft's scripting technologies.
 Fortunately, Mark Hammond's pywin32 has pretty much all you need
 for a workable Python adaptation. I haven't tried any of the fancier
 stuff like Async calls and so on, so I don't know if they'd work.

Since the COM implementation doesn't give much away to Python
 programmers, I've wrapped it in some lightweight classes with
 some getattr / setattr magic to ease the way. In particular:

+ The _wmi_namespace object itself will determine its classes
   and allow you to return all instances of any of them by
   using its name as an attribute. As an additional shortcut,
   you needn't specify the Win32_; if the first lookup fails
   it will try again with a Win32_ on the front.

   eg
   disks = wmi.WMI ().Win32_LogicalDisk ()

   or just

   disks = wmi.WMI ().LogicalDisk ()

  In addition, you can specify what would become the WHERE clause
   as keyword parameters.

   eg
   fixed_disks = wmi.WMI ().Win32_LogicalDisk (DriveType = 3)

+ The objects returned by a WMI lookup are wrapped in a Python
   class which determines their methods and classes and allows
   you to access them as though they were Python classes. The
   methods only allow named parameters.

   eg,
   for p in wmi.WMI ().Win32_Process ():
     if p.Name.lower () == 'notepad.exe':
       p.Terminate (Result=1)

+ Doing a print on one of the WMI objects will result in its
   GetObjectText_ method being called, which usually produces
   a meaningful printout of current values.
  The repr of the object will include its full WMI path,
   which lets you get directly to it if you need to.

+ You can get the associators and references of an object as
   a list of python objects by calling the associators () and
   references () methods on a WMI Python object.
   NB Don't do this on a Win32_ComputerSystem object; it will
    take all day and kill your machine!

   eg,
   for p in wmi.WMI ().Win32_Process ():
     if p.Name.lower () == 'notepad.exe':
       for r in p.references ():
         print r.Name

+ WMI classes (as opposed to instances) are first-class
  objects, so you can get hold of a class, and call
  its methods or set up a watch against it.

  eg
    process = wmi.WMI ().Win32_Process
    process.Create (CommandLine="notepad.exe")

+ To make it easier to use in embedded systems and py2exe-style
   executable wrappers, the module will not force early Dispatch.
   To do this, it uses a handy hack by Thomas Heller for easy access
   to constants.

+ Typical usage will be:

import wmi

vodev1 = wmi.WMI ("vodev1")
for disk in vodev1.Win32_LogicalDisk ():
  if disk.DriveType == 3:
    space = 100 * long (disk.FreeSpace) / long (disk.Size)
    print "%s has %d%% free" % (disk.Name, space)

Many thanks, obviously to Mark Hammond for creating the win32all
 extensions, but also to Alex Martelli and Roger Upole, whose
 c.l.py postings pointed me in the right direction.

(c) Tim Golden <mail@timgolden.me.uk> 5th June 2003
Licensed under the (GPL-compatible) MIT License:
http://www.opensource.org/licenses/mit-license.php

15th Aug 2006 1.1.1  . Fixed a small bug reported and patched by Jonas Bjering
7th Apr 2006  1.1    . Removed redundant qualifiers method of _wmi_object (the
                       qualifiers are held as a dictionary member of the class).
                     . If a moniker is passed which doesn't start with winmgmts:
                       then add it automatically (this helps with associations).
                     . Special-cased associations, whose properties are the paths
                       of the associated classes: when the properties are requested,
                       automatically return the instantiated class.
2nd Mar 2006  1.0    . Final release for v1.0
                     . Corrected example in .new method of _wmi_namespace, 
                       deprecating the previous inappropriate example of
                       Win32_Process, and substituting Win32_ProcessStartup.
11th Feb 2006 1.0rc6 . Adjusted .set method so it won't try to .Put_ unless
                       the instance has a path (ie has not been spawned).
10th Feb 2006 1.0rc5 . Fixed small bug in .new method of _wmi_class
10th Feb 2006 1.0rc4 . Added from_time function to convert Python times to WMI
                     . Remove final Put_ from .new method as some classes are
                       not intended to be created (eg Win32_ProcessStartup).
                     . Add .put method to allow explicit instance creation.
                     . Allow user to prevent the namespace from searching for
                       valid classes at startup. This makes the startup
                       much faster, but means you don't get a list of classes.
29th Nov 2005 1.0rc3 . Small changes to allow array of output parameters
                     . Added qualifiers to list of private attributes
                     . Added details of required privs to method docstring
                     . Fixed long-standing bug in Usage example
                     . Added provenance to WMI methods from MappingStrings
26th Oct 2005 1.0rc2 . Corrected __repr__ in _wmi_namespace
18th Oct 2005 1.0rc1 . Refactored into namespace, class, object classes.
                     . This makes it easier to use certain things, such
                       as the StdRegProv's registry object, and the XP
                       SystemRestore functionality.
                     . Added CompareTo_ to support __eq__ functionality.
                     . Added username/password support to WMI connection.
                       Saves having to do separate connect_server call.
                     . Added optional debug flag to WMI connection.
                     . Switched to MIT license (GPL-compatible and
                       not Python-specific).
                     . Bumped version to 1.0
25th May 2005 0.6b   . Removed late-dispatch code (EnsureDispatch)
                       and replaced with dynamic dispatch, using
                       Thomas Heller's ProvideConstants class to
                       avoid hard-coding WMI constants. This is
                       to help people using py2exe who would
                       otherwise need to specify one or more
                       typelibs.
19th May 2004 0.6    . Added namespace support to wmi.__init__.
                       This means you can now do, eg:
                       wmi.WMI (namespace="MicrosoftIISv2")
                     . _wmi_method parameters now check for array
                       parameters, showing them on the __doc__ and
                       raising an exception if the value passed in
                       is not iterable.
17th Jan 2004 0.5    . Added support for the WMI Registry interface. The new
                       module-level Registry method returns a WMI registry
                       object whose methods include EnumKeys, CreateValue &c.
15th Dec 2003 0.4    . Added machines_in_domain (from a post to python-win32 by "Sean")
                     . Factored out moniker construction to make it easier to support
                       use of StdRegProv to update registry. (Coming soon).
                     . Added support for a timeout on the event watcher; timeout is
                       specified in milliseconds and raises x_wmi_timed_out on a
                       call to the watcher object. This allows for the possibility
                       of pumping for waiting messages to prevent eg, the PythonWin
                       IDE locking up. See the docstring for the watch_for method.
                     . Added connect_server function, making it slightly easier to
                       construct a WMI object, eg with username and password.
10th Jul 2003 0.3    . Changes by Paul Moore to allow a ready-made WMI Services
                       object to be passed in (WMI.__init__).
                     . This header and the __VERSION__ number added by Tim G.
 9th Jul 2003 0.2    . Sundry changes by Tim G, including but not limited to:
                     + support for moniker parts (WMI.__init__)
                     + creating new instances of WMI classes (WMI.new)
                     + passing return value back from wmi methods (_wmi_method.__call__)
                     + better COM error-handling (handle_com_error)
 5th Jun 2003 0.1    Initial release by Tim Golden
"""

__VERSION__ = "1.1.1"

_DEBUG = False

import re
from win32com.client import GetObject, Dispatch
import pywintypes

#
# This class provided by Thomas Heller on c.l.py
#
class ProvideConstants (object):
   """A class which, when called on a win32com.client.Dispatch object,
   provides lazy access to constants defined in the typelib.

   They can be accessed as attributes of the _constants property."""
   def __init__(self, comobj):
     comobj.__dict__["_constants"] = self
     # Get the typelibrary's typecomp interface
     self.__typecomp = \
      comobj._oleobj_.GetTypeInfo().GetContainingTypeLib()[0].GetTypeComp()

   def __getattr__(self, name):
     if name.startswith("__") and name.endswith("__"):
       raise AttributeError, name
     result = self.__typecomp.Bind(name)
     # Bind returns a 2-tuple, first item is TYPEKIND,
     # the second item has the value
     if not result[0]:
       raise AttributeError, name
     return result[1].value

obj = GetObject ("winmgmts:")
ProvideConstants (obj)

wbemErrInvalidQuery = obj._constants.wbemErrInvalidQuery
wbemErrTimedout = obj._constants.wbemErrTimedout
wbemFlagReturnImmediately = obj._constants.wbemFlagReturnImmediately
wbemFlagForwardOnly = obj._constants.wbemFlagForwardOnly

def handle_com_error (error_info):
  hresult_code, hresult_name, additional_info, parameter_in_error = error_info
  exception_string = ["%s - %s" % (hex (hresult_code), hresult_name)]
  if additional_info:
    wcode, source_of_error, error_description, whlp_file, whlp_context, scode = additional_info
    exception_string.append ("  Error in: %s" % source_of_error)
    exception_string.append ("  %s - %s" % (hex (scode), error_description.strip ()))
  raise x_wmi, "\n".join (exception_string)

def from_time (year=None, month=None, day=None, hours=None, minutes=None, seconds=None, microseconds=None, timezone=None):
  """Returns a WMI time string of the form yyyymmddHHMMSS.mmmmmm+UUU
  replacing each placeholder by its respective integer value, or
  stars if None is supplied
  """
  def str_or_stars (i, length):
    if i is None:
      return "*" * length
    else:
      return str (i).rjust (length, "0")

  wmi_time = ""
  wmi_time += str_or_stars (year, 4)
  wmi_time += str_or_stars (month, 2)
  wmi_time += str_or_stars (day, 2)
  wmi_time += str_or_stars (hours, 2)
  wmi_time += str_or_stars (minutes, 2)
  wmi_time += str_or_stars (seconds, 2)
  wmi_time += "."
  wmi_time += str_or_stars (microseconds, 6)
  wmi_time += str_or_stars (timezone, 4)

  return wmi_time

def to_time (wmi_time):
  """Expects a WMI time string of the form yyyymmddHHMMSS.mmmmmm+UUU
  and returns:

  year, month, day, hours, minutes, seconds, microseconds, timezone

  If any part of the string is "*", returns None
  """
  def int_or_none (s, start, end):
    try:
      return int (s[start:end])
    except ValueError:
      return None

  year = int_or_none (wmi_time, 0, 4)
  month = int_or_none (wmi_time, 4, 6)
  day = int_or_none (wmi_time, 6, 8)
  hours = int_or_none (wmi_time, 8, 10)
  minutes = int_or_none (wmi_time, 10, 12)
  seconds = int_or_none (wmi_time, 12, 14)
  microseconds = int_or_none (wmi_time, 15, 21)
  timezone = wmi_time[21:]

  return year, month, day, hours, minutes, seconds, microseconds, timezone

#
# Exceptions
#
class x_wmi (Exception):
  pass

class x_wmi_invalid_query (x_wmi):
  pass

class x_wmi_timed_out (x_wmi):
  pass

class x_wmi_no_namespace (x_wmi):
  pass

WMI_EXCEPTIONS = {
  wbemErrInvalidQuery : x_wmi_invalid_query,
  wbemErrTimedout : x_wmi_timed_out
}

def _set (obj, attribute, value):
  """Helper function to add an attribute directly into the instance
   dictionary, bypassing possible __getattr__ calls
  """
  obj.__dict__[attribute] = value

#
# class _wmi_method
#
class _wmi_method:
  """A wrapper round a WMI COM object method"""

  def __init__ (self, ole_object, method_name):
    try:
      self.ole_object = Dispatch (ole_object)
      self.method = ole_object.Methods_ (method_name)
      self.qualifiers = {}
      for q in self.method.Qualifiers_:
        self.qualifiers[q.Name] = q.Value
      self.provenance = "\n".join (self.qualifiers.get ("MappingStrings", []))

      self.in_parameters = self.method.InParameters
      self.out_parameters = self.method.OutParameters
      if self.in_parameters is None:
        self.in_parameter_names = []
      else:
        self.in_parameter_names = [(i.Name, i.IsArray) for i in self.in_parameters.Properties_]
      if self.out_parameters is None:
        self.out_parameter_names = []
      else:
        self.out_parameter_names = [(i.Name, i.IsArray) for i in self.out_parameters.Properties_]

      doc = "%s (%s) => (%s)" % (
        method_name,
        ", ".join ([name + ("", "[]")[is_array] for (name, is_array) in self.in_parameter_names]),
        ", ".join ([name + ("", "[]")[is_array] for (name, is_array) in self.out_parameter_names])
      )
      privileges = self.qualifiers.get ("Privileges", [])
      if privileges:
        doc += " | Needs: " + ", ".join (privileges)
      self.__doc__ = doc
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def __call__ (self, **kwargs):
    """Execute the call to a WMI method, returning
     a tuple (even if is of only one value) containing
     the out and return parameters.
    """
    try:
      if self.in_parameters:
        parameter_names = {}
        for name, is_array in self.in_parameter_names:
          parameter_names[name] = is_array

        parameters = self.in_parameters
        for k, v in kwargs.items ():
          is_array = parameter_names.get (k)
          if is_array is None:
            raise AttributeError, "%s is not a valid parameter for %s" % (k, self.__doc__)
          else:
            if is_array:
              try: list (v)
              except TypeError: raise TypeError, "%s must be iterable" % k

          parameters.Properties_ (k).Value = v

        result = self.ole_object.ExecMethod_ (self.method.Name, self.in_parameters)
      else:
        result = self.ole_object.ExecMethod_ (self.method.Name)

      results = []
      for name, is_array in self.out_parameter_names:
        value = result.Properties_ (name).Value
        if is_array:
          #
          # Thanks to Jonas Bjering for bug report and path
          #
          results.append (list (value or []))
        else:
          results.append (value)
      return tuple (results)
      # return tuple ([result.Properties_ (name) for name in [n for (n, is_array) in self.out_parameter_names]])

    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def __repr__ (self):
    return "<function %s>" % self.__doc__

#
# class _wmi_object
#
class _wmi_object:
  """A lightweight wrapper round an OLE WMI object"""

  def __init__ (self, ole_object):
    try:
      _set (self, "ole_object", ole_object)
      _set (self, "_properties", [p.Name for p in self.ole_object.Properties_])
      _set (self, "_methods", [p.Name for p in self.ole_object.Methods_])
      _set (self, "qualifiers", {})
      for q in self.ole_object.Qualifiers_:
        self.qualifiers[q.Name] = q.Value
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def __str__ (self):
    """For a call to print <object> return the OLE description
     of the properties / values of the object
    """
    try:
      return self.ole_object.GetObjectText_ ()
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def __repr__ (self):
    """Indicate both the fact that this is a wrapped WMI object
     and the WMI object's own identifying class.
    """
    try:
      return "<%s: %s>" % (self.__class__.__name__, str (self.Path_.Path))
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def __getattr__ (self, attribute):
    """Attempt to pass attribute calls to the proxied COM object.
     If the attribute is recognised as a property, return its value;
     if it is recognised as a method, return a method wrapper which
     can then be called with parameters; otherwise pass the lookup
     on to the underlying object.
    """
    try:
      if attribute in self._properties:
        value = self.ole_object.Properties_ (attribute).Value
        #
        # If this is an association, its properties are
        #  actually the paths to the two aspects of the
        #  association, so translate them automatically
        #  into WMI objects.
        #
        if self.qualifiers.get ("Association", False):
          return WMI (moniker=value)
        else:
          return value
      elif attribute in self._methods:
        return _wmi_method (self.ole_object, attribute)
      else:
        return getattr (self.ole_object, attribute)
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def __setattr__ (self, attribute, value):
    """If the attribute to be set is valid for the proxied
     COM object, set that objects's parameter value; if not,
     raise an exception.
    """
    try:
      if attribute in self._properties:
        self.ole_object.Properties_ (attribute).Value = value
        if self.ole_object.Path_.Path:
          self.ole_object.Put_ ()
      else:
        raise AttributeError, attribute
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def __eq__ (self, other):
    """Use WMI's CompareTo_ to compare this object with
     another. Don't try to do anything if the other
     object is not a wmi object. It might be possible
     to compare this object's unique key with a string
     or something, but this doesn't seem to be univeral
     enough to merit a special case.
    """
    if isinstance (other, self.__class__):
      return self.ole_object.CompareTo_ (other.ole_object)
    else:
      raise x_wmi, "Can't compare a WMI object with something else"

  def put (self):
    self.ole_object.Put_ ()

  def set (self, **kwargs):
    """Set several properties of the underlying object
     at one go. This is particularly useful in combination
     with the new () method below. However, an instance
     which has been spawned in this way won't have enough
     information to write pack, so only try if the
     instance has a path.
    """
    if kwargs:
      try:
        for attribute, value in kwargs.items ():
          if attribute in self._properties:
            self.ole_object.Properties_ (attribute).Value = value
          else:
            raise AttributeError, attribute
        #
        # Only try to write the attributes
        #  back if the object exists.
        #
        if self.ole_object.Path_.Path:
          self.ole_object.Put_ ()
      except pywintypes.com_error, error_info:
        handle_com_error (error_info)

  def path (self):
    """Return the WMI URI to this object. Can be used to
     determine the path relative to the parent namespace. eg,

    pp0 = wmi.WMI ().Win32_ParallelPort ()[0]
    print pp0.path ().RelPath
    """
    try:
      return self.ole_object.Path_
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def derivation (self):
    """Return a tuple representing the object derivation for
     this object, with the most specific object first. eg,

    pp0 = wmi.WMI ().Win32_ParallelPort ()[0]
    print ' <- '.join (pp0.derivation ())
    """
    try:
      return self.ole_object.Derivation_
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def associators (self, wmi_association_class="", wmi_result_class=""):
    """Return a list of objects related to this one, optionally limited
     either by association class (ie the name of the class which relates
     them) or by result class (ie the name of the class which would be
     retrieved)

    eg
      c = wmi.WMI ()
      pp = c.Win32_ParallelPort ()[0]

      for i in pp.associators (wmi_association_class="Win32_PortResource"):
        print i

      for i in pp.associators (wmi_result_class="Win32_PnPEntity"):
        print i
    """
    try:
      return [
        _wmi_object (i) for i in \
          self.ole_object.Associators_ (
           strAssocClass=wmi_association_class,
           strResultClass=wmi_result_class
         )
      ]
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def references (self, wmi_class=""):
    """Return a list of associations involving this object, optionally
     limited by the result class (the name of the association class).
     
     NB Associations are treated specially; although WMI only returns
     the string corresponding to the instance of each associated object,
     this module will automatically convert that to the object itself.

    eg,
      c =  wmi.WMI ()
      sp = c.Win32_SerialPort ()[0]

      for i in sp.references ():
        print i

      for i in sp.references (wmi_class="Win32_SerialPortSetting"):
        print i
    """
    try:
      return [_wmi_object (i) for i in self.ole_object.References_ (strResultClass=wmi_class)]
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

#
# class _wmi_class
#
class _wmi_class (_wmi_object):
  """Currying class to assist in issuing queries against
   a WMI namespace. The idea is that when someone issues
   an otherwise unknown method against the WMI object, if
   it matches a known WMI class a query object will be
   returned which may then be called with one or more params
   which will form the WHERE clause. eg,

   c = wmi.WMI ()
   c_drive = c.Win32_LogicalDisk (Name='C:')
  """
  def __init__ (self, namespace, wmi_class):
    _wmi_object.__init__ (self, wmi_class)
    _set (self, "_namespace", namespace)
    _set (self, "_class_name", wmi_class.Path_.Class)

  def query (self, **where_clause):
    """Make it slightly easier to query against the class,
     by calling the namespace's query with the class preset.
     Won't work if the class has been instantiated directly.
    """
    if self._namespace is None:
      raise x_wmi_no_namespace, "You cannot query directly from a WMI class"

    try:
      wql = "SELECT * FROM " + self._class_name
      if where_clause:
        wql += " WHERE " + " AND ". join (["%s = '%s'" % (k, v) for k, v in where_clause.items ()])
      return self._namespace.query (wql)
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  __call__ = query

  def watch_for (
    self,
    notification_type=None,
    delay_secs=1,
    **where_clause
  ):
    if self._namespace is None:
      raise x_wmi_no_namespace, "You cannot watch directly from a WMI class"

    return self._namespace.watch_for (
      notification_type=notification_type,
      wmi_class=self._class_name,
      delay_secs=delay_secs,
      **where_clause
    )

  def instances (self):
    """Return a list of instances of the WMI class
    """
    try:
      return [_wmi_object (instance) for instance in self.Instances_ ()]
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def new (self, **kwargs):
    """This is the equivalent to the raw-WMI SpawnInstance_
     method. Note that there are relatively few uses for
     this, certainly fewer than you might imagine. Most
     classes which need to create a new *real* instance
     of themselves, eg Win32_Process, offer a .Create 
     method. SpawnInstance_ is generally reserved for
     instances which are passed as parameters to such
     .Create methods, a common example being the
     Win32_SecurityDescriptor, passed to Win32_Share.Create
     and other instances which need security.
     
    The example here is Win32_ProcessStartup, which
    controls the shown/hidden state etc. of a new
    Win32_Process instance.
    
    import win32con
    import wmi
    c = wmi.WMI ()
    startup = c.Win32_ProcessStartup.new (ShowWindow=win32con.SW_SHOWMINIMIZED)
    pid, retval = c.Win32_Process.Create (
      CommandLine="notepad.exe",
      ProcessStartupInformation=startup
    )
    
    NB previous versions of this module, used this function
    to create new process. This is *not* a good example
    of its use; it is better handled with something like
    the example above.
    """
    try:
      obj = _wmi_object (self.SpawnInstance_ ())
      obj.set (**kwargs)
      return obj
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

#
# class WMI
#
class _wmi_namespace:
  """A WMI root of a computer system. The classes attribute holds a list
   of the classes on offer. This means you can explore a bit with
   things like this:

   c = wmi.WMI ()
   for i in c.classes:
     if "user" in i.lower ():
       print i
  """
  def __init__ (self, namespace, find_classes):
    _set (self, "_namespace", namespace)
    #
    # wmi attribute preserved for backwards compatibility
    #
    _set (self, "wmi", namespace)

    # Initialise the "classes" attribute, to avoid infinite recursion in the
    # __getattr__ method (which uses it).
    self.classes = []
    #
    # Pick up the list of classes under this namespace
    #  so that they can be queried, and used as though
    #  properties of the namespace by means of the __getattr__
    #  hook below.
    # If the namespace does not support SubclassesOf, carry on
    #  regardless
    #
    if find_classes:
      try:
        _set (self, "classes", [c.Path_.Class for c in namespace.SubclassesOf ()])
      except AttributeError:
        pass

  def __repr__ (self):
    return "<_wmi_namespace: %s>" % self.wmi

  def __str__ (self):
    return repr (self)

  def get (self, moniker):
    try:
      return _wmi_object (self.wmi.Get (moniker))
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def handle (self):
    """The raw OLE object representing the WMI namespace"""
    return self._namespace

  def instances (self, class_name):
    """Return a list of instances of the WMI class. This is
     (probably) equivalent to querying with no qualifiers.

    eg system.instances ("Win32_LogicalDisk")

    or system.Win32_LogicalDisk ()
    """
    try:
      return [_wmi_object (obj) for obj in self._namespace.InstancesOf (class_name)]
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def new (self, wmi_class, **kwargs):
    """This is now implemented by a call to _wmi_namespace.new (qv)"""
    return getattr (self, wmi_class).new (**kwargs)

  new_instance_of = new

  def query (self, wql):
    """Perform an arbitrary query against a WMI object. Use the flags
     recommended by Microsoft to achieve a read-only, semi-synchronous
     query where the time is taken while looping through. Should really
     be a generator, but ...
    NB Backslashes need to be doubled up.
    """
    flags = wbemFlagReturnImmediately | wbemFlagForwardOnly
    wql = wql.replace ("\\", "\\\\")
    if _DEBUG: print wql
    try:
      return [
        _wmi_object (obj) for obj in \
         self._namespace.ExecQuery (
           strQuery=wql,
           iFlags=flags
         )
      ]
    except pywintypes.com_error, (hresult, hresult_text, additional, param_in_error):
      raise WMI_EXCEPTIONS.get (hresult, x_wmi (hresult))

  def watch_for (
    self,
    raw_wql=None,
    notification_type=None,
    wmi_class=None,
    delay_secs=1,
    **where_clause
  ):
    """Set up an event tracker on a WMI event. This function
     returns an wmi_watcher which can be called to get the
     next event. eg,

     c = wmi.WMI ()

     watcher = c.watch_for (raw_wql="SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_Process'")
     while 1:
       process_created = watcher ()
       print process_created.Name

     or,
     watcher = c.watch_for (
       notification_type="Creation",
       wmi_class="Win32_Process",
       delay_secs=2,
       Name='calc.exe'
     )
     calc_created = watcher ()

    Now supports timeout on the call to watcher, eg:

    import pythoncom
    import wmi
    c = wmi.WMI (privileges=["Security"])
    watcher1 = c.watch_for (
      notification_type="Creation",
      wmi_class="Win32_NTLogEvent",
      Type="error"
    )
    watcher2 = c.watch_for (
      notification_type="Creation",
      wmi_class="Win32_NTLogEvent",
      Type="warning"
    )

    while 1:
      try:
        error_log = watcher1 (500)
      except wmi.x_wmi_timed_out:
        pythoncom.PumpWaitingMessages ()
      else:
        print error_log

      try:
        warning_log = watcher2 (500)
      except wmi.x_wmi_timed_out:
        pythoncom.PumpWaitingMessages ()
      else:
        print warning_log
    """
    class_name = wmi_class
    if raw_wql:
      wql = raw_wql
    else:
      if where_clause:
        where = " AND " + " AND ".join (["TargetInstance.%s = '%s'" % (k, v) for k, v in where_clause.items ()])
      else:
        where = ""
      wql = \
        "SELECT * FROM __Instance%sEvent WITHIN %d WHERE TargetInstance ISA '%s' %s" % \
        (notification_type, delay_secs, class_name, where)

      if _DEBUG: print wql

    try:
      return _wmi_watcher (self._namespace.ExecNotificationQuery (wql))
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def __getattr__ (self, attribute):
    """Don't try to match against known classes, as the list may
    not have been requested (find_classes=False). Attempt to get
    the attribute as a class; if that fails, try getting it with
    Win32_ prepended. Failing that, assume it's a normal attribute
    and pass through.
    """
    try:
      return _wmi_class (self, self._namespace.Get (attribute))
    except pywintypes.com_error, error_info:
      try:
        return _wmi_class (self, self._namespace.Get ("Win32_" + attribute))
      except pywintypes.com_error, error_info:
        return getattr (self._namespace, attribute)

#
# class _wmi_watcher
#
class _wmi_watcher:
  """Helper class for WMI.watch_for below (qv)"""

  def __init__ (self, wmi_event):
    self.wmi_event = wmi_event

  def __call__ (self, timeout_ms=-1):
    """When called, return the instance which caused the event. Supports
     timeout in milliseconds (defaulting to infinite). If the watcher
     times out, x_wmi_timed_out is raised. This makes it easy to support
     watching for multiple objects.
    """
    try:
      return _wmi_object (self.wmi_event.NextEvent (timeout_ms).Properties_ ("TargetInstance").Value)
    except pywintypes.com_error, error_info:
      hresult_code, hresult_name, additional_info, parameter_in_error = error_info
      if additional_info:
        wcode, source_of_error, error_description, whlp_file, whlp_context, scode = additional_info
        if scode == wbemErrTimedout:
          raise x_wmi_timed_out
      handle_com_error (error_info)

PROTOCOL = "winmgmts:"
IMPERSONATION_LEVEL = "impersonate"
AUTHENTICATION_LEVEL = "default"
NAMESPACE = "root/cimv2"
def connect (
  computer=".",
  impersonation_level="",
  authentication_level="",
  authority="",
  privileges="",
  moniker="",
  wmi=None,
  namespace="",
  suffix="",
  user="",
  password="",
  find_classes=True,
  debug=False
):
  """The WMI constructor can either take a ready-made moniker or as many
   parts of one as are necessary. Eg,

   c = wmi.WMI (moniker="winmgmts:{impersonationLevel=Delegate}//remote")

   or

   c = wmi.WMI (computer="remote", privileges=["!RemoteShutdown", "Security"])

   I daren't link to a Microsoft URL; they change so often. Try Googling for
   WMI construct moniker and see what it comes back with.

   For complete control, a named argument "wmi" can be supplied, which
   should be a SWbemServices object, which you create yourself. Eg,

   loc = win32com.client.Dispatch("WbemScripting.SWbemLocator")
   svc = loc.ConnectServer(...)
   c = wmi.WMI(wmi=svc)

   This is the only way of connecting to a remote computer with a different
   username, as the moniker syntax does not allow specification of a user
   name.

   If the "wmi" parameter is supplied, all other parameters are ignored.

  """
  global _DEBUG
  _DEBUG = debug

  #
  # If namespace is a blank string, leave
  # it unaltered as it might to trying to
  # access the root namespace
  #
  #if namespace is None:
  #  namespace = NAMESPACE
    
  try:
    if wmi:
      obj = wmi

    elif moniker:
      if not moniker.startswith (PROTOCOL):
        moniker = PROTOCOL + moniker
      if _DEBUG: print moniker
      obj = GetObject (moniker)

    else:
      if user:
        if impersonation_level or authentication_level or privileges or suffix:
          raise x_wmi, "You can't specify an impersonation, authentication or privilege as well as a username"
        else:
          obj = connect_server (
            server=computer,
            namespace=namespace,
            user=user,
            password=password,
            authority=authority
          )

      else:
        moniker = construct_moniker (
          computer=computer,
          impersonation_level=impersonation_level or IMPERSONATION_LEVEL,
          authentication_level=authentication_level or AUTHENTICATION_LEVEL,
          authority=authority,
          privileges=privileges,
          namespace=namespace,
          suffix=suffix
        )
        if _DEBUG: print moniker
        obj = GetObject (moniker)

    wmi_type = get_wmi_type (obj)

    if wmi_type == "namespace":
      return _wmi_namespace (obj, find_classes)
    elif wmi_type == "class":
      return _wmi_class (None, obj)
    elif wmi_type == "instance":
      return _wmi_object (obj)
    else:
      raise x_wmi, "Unknown moniker type"

  except pywintypes.com_error, error_info:
    handle_com_error (error_info)

WMI = connect

def construct_moniker (
    computer=None,
    impersonation_level="Impersonate",
    authentication_level="Default",
    authority=None,
    privileges=None,
    namespace=None,
    suffix=None
):
  security = []
  if impersonation_level: security.append ("impersonationLevel=%s" % impersonation_level)
  if authentication_level: security.append ("authenticationLevel=%s" % authentication_level)
  #
  # Use of the authority descriptor is invalid on the local machine
  #
  if authority and computer: security.append ("authority=%s" % authority)
  if privileges: security.append ("(%s)" % ", ".join (privileges))

  moniker = [PROTOCOL]
  if security: moniker.append ("{%s}/" % ",".join (security))
  if computer: moniker.append ("/%s/" % computer)
  if namespace:
    parts = re.split (r"[/\\]", namespace)
    if parts[0] != 'root':
      parts.insert (0, "root")
    moniker.append ("/".join (parts))
#  if namespace and not namespace.startswith ("root/"): moniker.append ("root/")
#  if namespace: moniker.append (namespace)
  if suffix: moniker.append (":%s" % suffix)
  return "".join (moniker)

def get_wmi_type (obj):
  try:
    path = obj.Path_
  except AttributeError:
    return "namespace"
  else:
    if path.IsClass:
      return "class"
    else:
      return "instance"

def connect_server (
  server,
  namespace = "",
  user = "",
  password = "",
  locale = "",
  authority = "",
  security_flags = 0,
  named_value_set = None
):
  """Return a remote server running WMI

  server - name of the server
  namespace - namespace to connect to: defaults to whatever's defined as default
  user - username to connect as, either local or domain (dom\name or user@domain for XP)
  password: leave blank to use current context
  locale: desired locale in form MS_XXXX (eg MS_409 for Am En)
  authority: either "Kerberos:" or an NT domain. Not needed if included in user
  security_flags: if 0, connect will wait forever; if 0x80, connect will timeout at 2 mins
  named_value_set: typically empty, otherwise a context-specific SWbemNamedValueSet

  eg
    c = wmi.WMI (wmi=wmi.connect_server (server="remote_machine", user="myname", password="mypassword"))
  """
  if _DEBUG:
    print server
    print namespace
    print user
    print password
    print locale
    print authority
    print security_flags
    print named_value_set

  return Dispatch ("WbemScripting.SWbemLocator").\
    ConnectServer (
      server,
      namespace,
      user,
      password,
      locale,
      authority,
      security_flags,
      named_value_set
    )

def Registry (
  computer=None,
  impersonation_level="Impersonate",
  authentication_level="Default",
  authority=None,
  privileges=None,
  moniker=None
):

  if not moniker:
    moniker = construct_moniker (
      computer=computer,
      impersonation_level=impersonation_level,
      authentication_level=authentication_level,
      authority=authority,
      privileges=privileges,
      namespace="default",
      suffix="StdRegProv"
    )

  try:
    return _wmi_object (GetObject (moniker))

  except pywintypes.com_error, error_info:
    handle_com_error (error_info)

#
# From a post to python-win32 by Sean
#
def machines_in_domain (domain_name):
  adsi = Dispatch ("ADsNameSpaces")
  nt = adsi.GetObject ("","WinNT:")
  result = nt.OpenDSObject ("WinNT://%s" % domain_name, "", "", 0)
  result.Filter = ["computer"]
  domain = []
  for machine in result:
    domain.append (machine.Name)
  return domain

#
# Typical use test
#
if __name__ == '__main__':
  system = WMI ()
  for my_computer in system.Win32_ComputerSystem ():
    print "Disks on", my_computer.Name
    for disk in system.Win32_LogicalDisk ():
      print disk.Caption, disk.Description, disk.ProviderName or ""

