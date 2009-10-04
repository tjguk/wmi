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

class TestBasicConnections (unittest.TestCase):

    def test_basic_connection (self):
      self.assert_ (wmi.WMI ())
      
    def test_remote_connection (self):
      if "remote_connection" in excludes:
        warnings.warn ("Skipping test_remote_connection")
      else:
        if "remote" in settings:
          self.assert_ (wmi.WMI (settings['remote']))
        else:
          raise RuntimeError ("No value for settings.remote")
      
    def test_simple_moniker (self):
      self.assert_ (wmi.WMI (moniker="winmgmts:"))
      
    def test_moniker_with_class (self):
      c0 = wmi.WMI ().Win32_ComputerSystem
      c1 = wmi.WMI (moniker="winmgmts:Win32_ComputerSystem")
      self.assert_ (c0 == c1)
      
    def test_moniker_with_instance (self):
      for c0 in wmi.WMI ().Win32_ComputerSystem ():
        break
      c1 = wmi.WMI (moniker='winmgmts:Win32_ComputerSystem.Name="%s"' % c0.Name)
      self.assert_ (c0 == c1)
      
    def test_impersonation_levels (self):
      for impersonation in ["identify", "impersonate", "delegate"]:
        self.assert_ (wmi.WMI (impersonation_level=impersonation))
      
    def test_authentication_levels (self):
      for authentication in ["default", "none", "connect", "call", "pkt", "pktintegrity", "pktprivacy"]:
        try:
          c = wmi.WMI (authentication_level=authentication)
        except wmi.x_access_denied:
          warnings.warn ("Access denied for authentication level %s" % authentication)
        else:
          self.assert_ (c)
      
    def test_authority (self):
      if "authority" in excludes:
        warnings.warn ("Skipping test_authority")
      else:
        for authority in ["kerberos", "ntlmdomain"]:
          self.assert_ (wmi.WMI (authority=authority))
        
    def test_privileges (self):
      self.assert_ (wmi.WMI (privileges=['security', '!shutdown']))

    def test_namespace (self):
      self.assert_ (wmi.WMI (namespace="DEFAULT"))
        
    def test_suffix (self):
      self.assert_ (wmi.WMI (namespace="DEFAULT", suffix="StdRegProv"))
      
    def test_user_password (self):
      if "remote_connection" in excludes:
        warnings.warn ("Skipping test_user_password because remote_connection excluded")
      else:
        self.assert_ (wmi.WMI (computer=settings['remote'], user=settings['user'], password=settings['password']))
    
    def test_find_classes (self):
      self.assert_ (wmi.WMI (find_classes=True).classes)
      self.assertFalse (wmi.WMI (find_classes=False).classes)

if __name__ == '__main__':
  unittest.main ()
