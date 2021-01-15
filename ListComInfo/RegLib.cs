using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ListComInfo
{
    internal class RegLib
    {
        internal static IEnumerable<string> GetSubKeyNames(string keyName, RegistryView registryView = RegistryView.Registry64)
        {
            RegistryHive hive = GetHive(keyName);

            var slashpos = keyName.IndexOf('\\');
            var extra = keyName.Last() == '\\' ? "" : "\\";

            using (var key = RegistryKey.OpenBaseKey(hive, registryView))
            using (var subkey = key.OpenSubKey(keyName.Substring(slashpos + 1)))
            {
                var subKeyNames = subkey.GetSubKeyNames();
                foreach (var skname in subKeyNames)
                {
                    yield return $"{keyName}{extra}{skname}";
                }
            }
        }

        public static string GetDefaultValue(string keyName, RegistryView registryView = RegistryView.Registry64)
        {
            var hive = GetHive(keyName);
            var slashpos = keyName.IndexOf('\\');
            using (var key = RegistryKey.OpenBaseKey(hive, registryView))
            using (var subkey = key.OpenSubKey(keyName.Substring(slashpos + 1)))
            {
                return subkey?.GetValue(string.Empty)?.ToString();
            }
        }

        private static RegistryHive GetHive(string keyName)
        {
            if (keyName.StartsWith(@"HKCR\"))
            {
                return RegistryHive.ClassesRoot;
            }
            else if (keyName.StartsWith(@"HKLM\"))
            {
                return RegistryHive.LocalMachine;
            }
            else if (keyName.StartsWith(@"HKCU\"))
            {
                return RegistryHive.CurrentUser;
            }
            else if (keyName.StartsWith(@"HKU\"))
            {
                return RegistryHive.Users;
            }
            else if (keyName.StartsWith(@"HKCC\"))
            {
                return RegistryHive.CurrentConfig;
            }
            else if (keyName.StartsWith(@"HKDD\"))
            {
                return RegistryHive.DynData;
            }
            else
            {
                throw new ArgumentException($"You must specify the three or four letter abbreviation for the registry hive (e.g. HKLM) and then a backslash.");
            }
        }

        internal static IEnumerable<(string key, string value)> GetValues(string keyName, RegistryView registryView = RegistryView.Registry64)
        {
            var hive = GetHive(keyName);
            var slashpos = keyName.IndexOf('\\');
            using (var key = RegistryKey.OpenBaseKey(hive, registryView))
            using (var subkey = key.OpenSubKey(keyName.Substring(slashpos + 1)))
            {
                foreach (var valueName in subkey.GetValueNames())
                {
                    var value = subkey.GetValue(valueName).ToString();
                    yield return (valueName, value);
                }
            }
        }

        public static string GetLastSegment(string key)
        {
            var slashpos = key.LastIndexOf('\\');
            return key.Substring(slashpos + 1);
        }
    }
}
