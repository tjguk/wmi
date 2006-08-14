"""wmi - a lightweight Python wrapper around Microsoft's WMI interface

Windows Management Instrumentation (WMI) is Microsoft's answer to
 the DMTF's Common Information Model. It allows you to query just
 about any conceivable piece of information from any computer which
 is running the necessary agent and over which have you the
 necessary authority.

The implementation is by means of COM/DCOM and most of the examples
 assume you're running one of Microsoft's scripting technologies.
 Fortunately, Mark Hammond's win32all has pretty much all you need
 for a workable Python adaptation. I haven't tried any of the fancier
 stuff like Async calls and so on, so I don't know if they'd work.

Since the COM implementation doesn't give much away to Python
 programmers, I've wrapped it in some lightweight classes with
 some getattr / setattr magic to ease the way. In particular:

+ The wmi object itself will determine its classes and allow you
   to return all instances of any of them by using its name as
   a property. As an additional shortcut, you needn't specify
   the Win32_; if the first lookup fails it will try again with
   a Win32_ on the front.

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

+ You can create a new instance of a WMI class by using
   the new method of the WMI object and passing the WMI
   class name as text.

  eg
  remote_process = wmi.WMI ("remote_machine").new ("Win32_process")
  remote_process.Create (CommandLine="c:\\temp\\newprog\\setup.exe /s")

+ Typical usage will be:

import wmi

vodev1 = wmi.WMI ("vodev1")
for disk in vodev1.Win32_LogicalDisk ():
  if disk.DriveType == 3:
    print "%s has %d%% free" % \\
     (disk.Name, 100 * long (disk.FreeSpace) / long (disk.Size)

Many thanks, obviously to Mark Hammond for creating the win32all
 extensions, but also to Alex Martelli and Roger Upole, whose
 c.l.py postings pointed me in the right direction.

(c) Tim Golden <tim.golden@iname.com> 5th June 2003

25th May 2005 0.6b Removed late-dispatch code (EnsureDispatch)
                    and replaced with dynamic dispatch, using
                    Thomas Heller's ProvideConstants class to
                    avoid hard-coding WMI constants. This is
                    to help people using py2exe who would
                    otherwise need to specify one or more
                    typelibs.
19th May 2004 0.6  Added namespace support to wmi.__init__.
                   This means you can now do, eg: 
                    wmi.WMI (namespace="MicrosoftIISv2")
                   _wmi_method parameters now check for array
                    parameters, showing them on the __doc__ and
                    raising an exception if the value passed in
                    is not iterable.
17th Jan 2004 0.5  Added support for the WMI Registry interface. The new
                    module-level Registry method returns a WMI registry
                    object whose methods include EnumKeys, CreateValue &c.
15th Dec 2003 0.4  Added machines_in_domain (from a post to python-win32 by "Sean")
                   Factored out moniker construction to make it easier to support
                    use of StdRegProv to update registry. (Coming soon).
                   Added support for a timeout on the event watcher; timeout is
                    specified in milliseconds and raises x_wmi_timed_out on a
                    call to the watcher object. This allows for the possibility
                    of pumping for waiting messages to prevent eg, the PythonWin 
                    IDE locking up. See the docstring for the watch_for method.
                   Added connect_server function, making it slightly easier to
                    construct a WMI object, eg with username and password.
10th Jul 2003 0.3  Changes by Paul Moore to allow a ready-made WMI Services
                    object to be passed in (WMI.__init__).
                   This header and the __VERSION__ number added by Tim G.
 9th Jul 2003 0.2  Sundry changes by Tim G, including but not limited to:
                    + support for moniker parts (WMI.__init__)
                    + creating new instances of WMI classes (WMI.new)
                    + passing return value back from wmi methods (_wmi_method.__call__)
                    + better COM error-handling (handle_com_error)
 5th Jun 2003 0.1  Initial release by Tim Golden
"""

__VERSION__ = "0.6b"

_DEBUG = False

import win32com.client
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
         
obj = win32com.client.GetObject ("winmgmts:")
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

def to_time (wmi_time):
  """Expects a WMI time string of the form yyyymmddHHMMSS.mmmmmm+UUU
  and returns:

  year, month, day, hours, minutes, seconds, microseconds, timezone
  """
  year, month, day = [int (i) for i in wmi_time[:4], wmi_time[4:6], wmi_time[6:8]]
  hours, minutes, seconds = [int (i) for i in wmi_time[8:10], wmi_time[10:12], wmi_time[12:14]]
  microseconds = int (wmi_time[15:21])
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
      self.ole_object = ole_object
      self.method_name = method_name
      self.in_parameters = ole_object.Methods_ (method_name).InParameters
      self.out_parameters = ole_object.Methods_ (method_name).OutParameters

      if self.in_parameters is None:
        self.in_parameter_names = []
      else:
        self.in_parameter_names = [(i.Name, i.IsArray) for i in self.in_parameters.Properties_]
      if self.out_parameters is None:
        self.out_parameter_names = []
      else:
        self.out_parameter_names = [(i.Name, i.IsArray) for i in self.out_parameters.Properties_]

      self.__doc__ = "%s (%s) => (%s)" % (
        method_name,
        ", ".join ([name + ("", "[]")[is_array] for (name, is_array) in self.in_parameter_names]),
        ", ".join ([name + ("", "[]")[is_array] for (name, is_array) in self.out_parameter_names])
      )
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

        result = self.ole_object.ExecMethod_ (self.method_name, self.in_parameters)
      else:
        result = self.ole_object.ExecMethod_ (self.method_name)

      return tuple ([result.Properties_ (name) for name in [n for (n, is_array) in self.out_parameter_names]])

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

  def __getattr__ (self, attribute):
    """Attempt to pass attribute calls to the proxied COM object.
     If the attribute is recognised as a property, return its value;
     if it is recognised as a method, return a method wrapper which
     can then be called with parameters; otherwise raise an exception
    """
    try:
      if attribute in self._properties:
        return self.ole_object.Properties_ (attribute).Value
      elif attribute in self._methods:
        return _wmi_method (self.ole_object, attribute)
      else:
        raise AttributeError, attribute
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
        self.ole_object.Put_ ()
      else:
        raise AttributeError, attribute
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

  def qualifiers (self):
    try:
      return self.ole_object.Qualifiers_
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
     limited by the result class (the name of the association class)

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

#
# class _wmi_query
#
class _wmi_query:
  """Currying class to assist in issuing queries against
   a WMI namespace. The idea is that when someone issues
   an otherwise unknown method against the WMI object, if
   it matches a known WMI class a query object will be
   returned which may then be called with one or more params
   which will form the WHERE clause. eg,

   c = wmi.WMI ()
   c_drive = c.Win32_LogicalDisk (Name='C:')
  """
  def __init__ (self, wmi, wmi_class):
    self.wmi = wmi
    self.wmi_class = wmi_class

  def __call__ (self, **where_clause):
    wql = "SELECT * FROM " + self.wmi_class
    if where_clause:
      wql += " WHERE " + " AND ". join (["%s = '%s'" % (k, v) for k, v in where_clause.items ()])
    return self.wmi.query (wql)

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

  moniker = ["winmgmts:"]
  if security: moniker.append ("{%s}/" % ",".join (security))
  if computer: moniker.append ("/%s/" % computer)
  moniker.append ("root/")
  if namespace: moniker.append (namespace)
  if suffix: moniker.append (":%s" % suffix)
  return "".join (moniker)

#
# class WMI
#
class WMI:
  """A WMI root of a computer system"""

  def __init__ (
    self,
    computer=None,
    impersonation_level="Impersonate",
    authentication_level="Default",
    authority=None,
    privileges=None,
    moniker=None,
    wmi=None,
    namespace="cimv2",
    suffix=None
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
    # Initialise the "classes" attribute, to avoid infinite recursion in the
    # __getattr__ method (which uses it).
    self.classes = []
    self.computer = computer

    if wmi:
      pass

    elif moniker:
      _set (self, "moniker", moniker)

    else:
      _set (self, "moniker",
        construct_moniker (
          computer=computer,
          impersonation_level=impersonation_level,
          authentication_level=authentication_level,
          authority=authority,
          privileges=privileges,
          namespace=namespace,
          suffix=suffix
        )
      )

    if _DEBUG: print self.moniker
    try:
      if not wmi:
        wmi = win32com.client.GetObject (self.moniker)
      _set (self, "wmi", win32com.client.Dispatch (wmi._oleobj_))

      #
      # Pick up the list of classes under this namespace
      #  so that they can be queried, and used as though
      #  properties of the namespace by means of the __getattr__
      #  hook below.
      # If the namespace does not support SubclassesOf, carry on
      #  regardless
      #
      try:
        _set (self, "classes", [c.Path_.Class for c in wmi.SubclassesOf ()])
      except AttributeError:
        pass

    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def get (self, moniker):
    try:
      return _wmi_object (self.wmi.Get (moniker))
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def handle (self):
    """The raw OLE object representing the WMI system"""
    return self.wmi

  def instances (self, wmi_class):
    """Return a list of instances of the WMI class

    eg system.instances ("Win32_LogicalDisk")
    """
    try:
      return [_wmi_object (obj) for obj in self.wmi.InstancesOf (wmi_class)]
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def new (self, wmi_class):
    """Create a new <whatever>, typically something like
     Win32_Process, eg:

     c = wmi.WMI ("remote_machine")
     for p in c.Win32_Process (name="notepad.exe"): print p
     c.new ("Win32_Process").Create (CommandLine="notepad.exe")
     for p in c.Win32_Process (name="notepad.exe"): print p
     p.Terminate ()
     for p in c.Win32_Process (name="notepad.exe"): print p
    """
    if wmi_class not in self.classes:
      raise Exception, wmi_class + " is not a WMI class"

    # return _wmi_object (self.wmi.Get (wmi_class).SpawnInstance_ ())
    return _wmi_object (self.wmi.Get (wmi_class))

  new_instance_of = new

  def run (self, command_line):
    """Convenience function to run a program against this computer"""
    self.new ("Win32_Process").Create (CommandLine=command_line)

  def query (self, wql):
    """Perform an arbitrary query against a WMI object. Use the flags
     recommended by Microsoft to achieve a read-only, semi-synchronous
     query where the time is taken while looping through. Should really
     be a generator, but ...
    NB Backslashes need to be doubled up.
    """
    flags = wbemFlagReturnImmediately | wbemFlagForwardOnly
    wql = wql.replace ("\\", "\\\\")
    try:
      return [
        _wmi_object (obj) for obj in \
         self.wmi.ExecQuery (
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
    if raw_wql:
      wql = raw_wql
    else:
      if where_clause:
        where = " AND " + " AND ".join (["TargetInstance.%s = '%s'" % (k, v) for k, v in where_clause.items ()])
      else:
        where = ""
      wql = \
        "SELECT * FROM __Instance%sEvent WITHIN %d WHERE TargetInstance ISA '%s' %s" % \
        (notification_type, delay_secs, wmi_class, where)

    try:
      return _wmi_watcher (self.wmi.ExecNotificationQuery (wql))
    except pywintypes.com_error, error_info:
      handle_com_error (error_info)

  def __getattr__ (self, attribute):
    """Attempt to match the property requested against one of
     the classes picked up at the beginning. See the explanation
     against the _wmi_query above.
    """
    if attribute in self.classes:
      return _wmi_query (self, attribute)
    elif "Win32_" + attribute in self.classes:
      return _wmi_query (self, "Win32_" + attribute)
    else:
      raise AttributeError, attribute

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
  return win32com.client.Dispatch ("WbemScripting.SWbemLocator").\
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
    print moniker
    registry = win32com.client.GetObject (moniker)
    return _wmi_object (win32com.client.Dispatch (registry))

  except pywintypes.com_error, error_info:
    handle_com_error (error_info)

#
# From a post to python-win32 by Sean
#
def machines_in_domain (domain_name):
  adsi = win32com.client.Dispatch ("ADsNameSpaces")
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

