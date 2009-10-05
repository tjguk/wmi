import os, sys
import datetime
import ConfigParser
import Queue
import subprocess
import threading
import unittest
import warnings

import wmi

ini = ConfigParser.SafeConfigParser ()
ini.read ("wmitest.ini")
settings = {}
if ini.has_section ("settings"):
  settings.update (ini.items ("settings"))
excludes = [i.strip () for i in settings.get ("excludes", "").split (",")]

COMPUTERS = [None, "."]
if not "remote_connection" in excludes:
  COMPUTERS.append (settings['remote'])
IMPERSONATION_LEVELS = [None, "identify", "impersonate", "delegate"]
AUTHENTICATION_LEVELS = [None, "default", "none", "connect", "call", "pkt", "pktintegrity", "pktprivacy"]
AUTHORITIES = [None, "kerberos", "ntlmdomain"]
PRIVILEGES = [None, ['security', '!shutdown']]
NAMESPACES = [None, "root/cimv2", "default"]

class TestBasicConnections (unittest.TestCase):

  def test_basic_connection (self):
    "Check that a standard connection works"
    self.assert_ (wmi.WMI ())
    
  def test_remote_connection (self):
    "Check that a remote connection works, if specified"
    if "remote_connection" in excludes:
      warnings.warn ("Skipping test_remote_connection")
    else:
      if "remote" in settings:
        self.assert_ (wmi.WMI (settings['remote']))
      else:
        raise RuntimeError ("No value for settings.remote")
    
  def test_simple_moniker (self):
    "Check that a simple moniker works"
    self.assert_ (wmi.WMI (moniker="winmgmts:"))
    
  def test_moniker_with_class (self):
    "Check that specifying a class in moniker works"
    c0 = wmi.WMI ().Win32_ComputerSystem
    c1 = wmi.WMI (moniker="winmgmts:Win32_ComputerSystem")
    self.assert_ (c0 == c1)
    
  def test_moniker_with_instance (self):
    "Check that specifying an instance in the moniker works"
    for c0 in wmi.WMI ().Win32_ComputerSystem ():
      break
    c1 = wmi.WMI (moniker='winmgmts:Win32_ComputerSystem.Name="%s"' % c0.Name)
    self.assert_ (c0 == c1)
    
  def test_impersonation_levels (self):
    "Check that specifying an impersonation level works"
    for impersonation in IMPERSONATION_LEVELS:
      self.assert_ (wmi.WMI (impersonation_level=impersonation))
    
  def test_authentication_levels (self):
    "Check that specifying an authentication level works"
    for authentication in AUTHENTICATION_LEVELS:
      try:
        c = wmi.WMI (authentication_level=authentication)
      except wmi.x_access_denied:
        warnings.warn ("Access denied for authentication level %s" % authentication)
      else:
        self.assert_ (c)
    
  def test_authority (self):
    "Check that specifying an authority works"
    if "authority" in excludes:
      warnings.warn ("Skipping test_authority")
    else:
      for authority in AUTHORITIES:
        self.assert_ (wmi.WMI (authority=authority))
      
  def test_privileges (self):
    "Check that specifying privileges works"
    for privileges in PRIVILEGES:
      self.assert_ (wmi.WMI (privileges=privileges))

  def test_namespace (self):
    "Check that specifying a namespace works"
    for namespace in NAMESPACES:
      self.assert_ (wmi.WMI (namespace=namespace))
      
  def test_suffix (self):
    "Check that a suffix returns the class of that name"
    self.assert_ (wmi.WMI (namespace="DEFAULT", suffix="StdRegProv") == wmi.WMI (namespace="DEFAULT").StdRegProv)
    
  def test_user_password (self):
    "Check that username & password are passed through for a remote connection"
    if "remote_connection" in excludes:
      warnings.warn ("Skipping test_user_password because remote_connection excluded")
    else:
      self.assert_ (wmi.WMI (computer=settings['remote'], user=settings['user'], password=settings['password']))
      
  def test_too_much_authentication (self):
    "Check that user/password plus any other authentication raises exception"
    self.assertRaises (wmi.x_wmi_authentication, wmi.WMI, computer='***', user="***", password="***", impersonation_level="***")
    self.assertRaises (wmi.x_wmi_authentication, wmi.WMI, computer='***', user="***", password="***", authentication_level="***")
    self.assertRaises (wmi.x_wmi_authentication, wmi.WMI, computer='***', user="***", password="***", privileges=["***"])
    self.assertRaises (wmi.x_wmi_authentication, wmi.WMI, computer='***', user="***", password="***", suffix="***")
  
  def test_local_user_password (self):
    "Check that user/password for local connection raises exception"
    self.assertRaises (wmi.x_wmi_authentication, wmi.WMI, user="***", password="***")
  
  def test_find_classes (self):
    "Check ability to switch class scan on and off"
    self.assert_ (wmi.WMI (find_classes=True).classes)
    self.assertFalse (wmi.WMI (find_classes=False).classes)
    
  def test_find_classes_false (self):
    "By default, don't scan for classes"
    self.assertFalse (wmi.WMI ().classes)
      
class TestMoniker (unittest.TestCase):
  
  def test_moniker (self):
    """Look at all possible options for moniker construction and pass
    them through to a WMI connector
    """
    if "authority" in excludes:
      warnings.warn ("Skipping authorities in test_moniker")
      authorities = [None]
    else:
      authorities = AUTHORITIES
    
    for computer in COMPUTERS:
      for impersonation_level in IMPERSONATION_LEVELS:
        for authentication_level in AUTHENTICATION_LEVELS:
          for authority in authorities:
            for privileges in PRIVILEGES:
              for namespace in NAMESPACES:
                moniker = wmi.construct_moniker (
                  computer=computer,
                  impersonation_level=impersonation_level,
                  authority=authority,
                  privileges=privileges,
                  namespace=namespace
                )
                self.assert_ (wmi.WMI (moniker=moniker), "Moniker failed: %s" % moniker)
  
  def test_moniker_root_namespace (self):
    "Check that namespace is prefixed by root if needed"
    self.assertEquals (wmi.construct_moniker (namespace="default"), "winmgmts:root/default")
    self.assertEquals (wmi.construct_moniker (namespace="root/default"), "winmgmts:root/default")
                
class TestFunctions (unittest.TestCase):
  
  times = [
    ((2000, 1, 1), "20000101******.******+***"),
    ((2000, 1, 1, 10, 0, 0), "20000101100000.******+***"),
    ((2000, 1, 1, 10, 0, 0, 100), "20000101100000.000100+***"),
    ((2000, 1, 1, 10, 0, 0, 100, "GMT"), "20000101100000.000100+GMT")
  ]
  
  def test_signed_to_unsigned (self):
    tests = [
      (0, 0),
      (-1, 0xffffffff),
      (+1, 1),
      (sys.maxint, 0x7fffffff),
      (-sys.maxint, 0x80000001)
    ]
    for signed, unsigned in tests:
      self.assertEquals (wmi.signed_to_unsigned (signed), unsigned)
      
  def test_from_1601 (self):
    "Check conversion from 100-ns intervals since 1601 (!)"
    self.assertEquals (wmi.from_1601 (0), datetime.datetime (1601, 1, 1))
    self.assertEquals (wmi.from_1601 (24 * 60 * 60 * 10 * 1000 * 1000), datetime.datetime (1601, 1, 2))
    
  def test_from_time (self):
    "Check conversion from time-tuple to time-string"
    for t, s in self.times:
      self.assertEquals (wmi.from_time (*t), s)
    
  def test_to_time (self):
    "Check conversion from time-string to time-tuple"
    for t, s in self.times:
      t = tuple (list (t) + ([None] * 8))[:8]
      self.assertEquals (wmi.to_time (s), t)
  
  def test_get_wmi_type (self):
    "Check that namespace, class & instance are identified correctly"
    self.assertEquals (wmi.get_wmi_type (wmi.WMI ()), "namespace")
    self.assertEquals (wmi.get_wmi_type (wmi.WMI ().Win32_ComputerSystem), "class")
    for i in wmi.WMI ().Win32_ComputerSystem ():
      self.assertEquals (wmi.get_wmi_type (i), "instance")
    
  def test_registry (self):
    """Convenience Registry function is identical to picking 
    the StdRegProv class out of the DEFAULT namespace"""
    self.assertEquals (wmi.Registry (), wmi.WMI (namespace="DEFAULT").StdRegProv)

class TestNamespace (unittest.TestCase):
  
  def setUp (self):
    self.connection = wmi.WMI (namespace="root/cimv2", find_classes=False)
    self.logical_disks = set (self.connection.Win32_LogicalDisk ())
  
  def test_subclasses_of_simple (self):
    self.assert_ ("Win32_ComputerSystem" in self.connection.subclasses_of ())

  def test_subclasses_of_subtree (self):
    self.assert_ ("Win32_Desktop" in self.connection.subclasses_of ("CIM_Setting"))
  
  def test_subclasses_of_pattern (self):
    self.assert_ (set (["Win32_LogicalDisk", "Win32_MappedLogicalDisk"]) <= set (self.connection.subclasses_of ("CIM_LogicalDevice", "Win32_.*Disk")))

  def test_instances (self):
    self.assertEquals (self.logical_disks, set (self.connection.instances ("Win32_LogicalDisk")))
    
  def test_new (self):
    "Check this is an alias for the new method of the equivalent class"
    self.assertEquals (self.connection.new ("Win32_Process")._instance_of, self.connection.Win32_Process)
    
  def test_query (self):
    self.assertEquals (self.logical_disks, set (self.connection.query ("SELECT * FROM Win32_LogicalDisk")))
    
  def test_ipython_attributes (self):
    connection = wmi.WMI (find_classes=True)
    self.assertEquals (connection._getAttributeNames (), [i for i in connection.classes if not i.startswith ("__")])
    
  def test_getattr (self):
    "Check that WMI classes are returned by attribute access on their namespace"
    connection = wmi.WMI (find_classes=True)
    for c in list (connection.classes)[:5]:
      wmi_class = getattr (connection, c)
      self.assert_ (isinstance (wmi_class, wmi._wmi_class))
      self.assertEquals (wmi_class._class_name, c)
      
  def test_watch_for (self):
    """Check that the watch_for method returns a watcher. The watcher itself
    will be tested elsewhere.
    """
    watcher = self.connection.watch_for (
      wmi_class="Win32_Process"
    )
    self.assert_ (isinstance (watcher, wmi._wmi_watcher))
    
class TestClass (unittest.TestCase):
  
  def setUp (self):
    self.connection = wmi.WMI (namespace="root/cimv2", find_classes=False)
    
  def test_class_from_namespace (self):
    self.assert_ (self.connection.Win32_ComputerSystem._namespace is self.connection)
      
  def test_class_without_namespace (self):
    wmi_class = wmi.GetObject ("winmgmts:Win32_ComputerSystem")
    self.assert_ (wmi._wmi_class (None, wmi_class)._namespace)
    
  def test_query (self):
    self.assertEquals (
      set (self.connection.Win32_ComputerSystem.query ()), 
      set (self.connection.query ("SELECT * FROM Win32_ComputerSystem"))
    )
    
  def test_query_with_where (self):
    this_drive = os.getcwd ()[:2]
    for drive in self.connection.Win32_LogicalDisk (Name=this_drive):
      self.assertEquals (drive.Name, this_drive)

  def test_query_with_fields (self):
    this_drive = os.getcwd ()[:2]
    properties = set (["MediaType"])
    self.assert_ ("Name" not in properties)
    for drive in self.connection.Win32_LogicalDisk (properties, Name=this_drive):
      self.assertEquals (set (drive.properties), set (properties))
      self.assert_ (drive.MediaType)
      self.assertRaises (AttributeError, getattr, drive, "Name")
      
  def test_watch_for (self):
    """Check that the watch_for method returns a watcher. The watcher itself
    will be tested elsewhere.
    """
    watcher = self.connection.Win32_Process.watch_for ()
    self.assert_ (isinstance (watcher, wmi._wmi_watcher))
    
  def test_instances (self):
    self.assertEquals (
      set (self.connection.Win32_LogicalDisk ()), 
      set (self.connection.Win32_LogicalDisk.instances ())
    )
    
  def test_new (self):
    process = self.connection.Win32_Process.new ()
    self.assertEquals (wmi.get_wmi_type (process), "instance")
    self.assertEquals (process._instance_of, self.connection.Win32_process)


class TestWatcher (unittest.TestCase):
  
  def setUp (self):
    self.connection = wmi.WMI ()
    
  def test_creation (self):
    
    def _create (queue):
      queue.put (subprocess.Popen ([sys.executable, "-c", "pass"]))
    
    watcher = self.connection.Win32_Process.watch_for (
      notification_type="creation"
    )
    q = Queue.Queue ()
    t = threading.Timer (1, _create, (q,))
    try:
      t.start ()
      found_process = watcher (timeout_ms=2000.0)
      spawned_process = q.get_nowait ()
      self.assertEqual (int (found_process.Handle), spawned_process.pid)
    finally:
      t.cancel ()

  def test_extrinsic_event (self):
    
    def _create (queue):
      queue.put (subprocess.Popen ([sys.executable, "-c", "pass"]))
    
    watcher = self.connection.Win32_ProcessStartTrace.watch_for (
      fields=["*"], 
      ProcessName=os.path.basename (sys.executable)
    )
    q = Queue.Queue ()
    t = threading.Timer (0.5, _create, (q,))
    try:
      t.start ()
      found_process = watcher (timeout_ms=2000.0)
      spawned_process = q.get_nowait ()
      self.assert_ (isinstance (found_process, wmi._wmi_event))
      self.assertEqual (int (found_process.ProcessID), spawned_process.pid)
    finally:
      t.cancel ()
      

if __name__ == '__main__':
  unittest.main ()
