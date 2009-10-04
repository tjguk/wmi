import ConfigParser
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

if __name__ == '__main__':
  unittest.main ()
