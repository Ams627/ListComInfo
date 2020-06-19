using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Threading.Tasks;

namespace ListTypeLibs
{
    class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    PrintTypeLibs();
                }
                else if (args.Length == 1 && args[0] == "-c")
                {
                    PrintClsIds();
                }
            }
            catch (Exception ex)
            {
                var fullname = System.Reflection.Assembly.GetEntryAssembly().Location;
                var progname = Path.GetFileNameWithoutExtension(fullname);
                Console.Error.WriteLine($"{progname} Error: {ex.Message}");
            }

        }

        private static void PrintClsIds()
        {

            foreach (var name in RegLib.GetSubKeyNames(@"HKCR\clsid"))
            {
                Console.WriteLine($"{name}");
                foreach (var k2 in RegLib.GetSubKeyNames(name))
                {
                    if (k2.EndsWith("InprocServer32", StringComparison.InvariantCultureIgnoreCase))
                    {
                        var dll = RegLib.GetDefaultValue(k2);
                        Console.WriteLine($"{k2} {dll}");
                    }
                }
            }
        }

        private static void PrintTypeLibs()
        {
            var versionPattern = @"\\[a-z0-9]{1,3}\.[a-z0-9]{1,3}$";

            foreach (var name in RegLib.GetSubKeyNames(@"HKCR\typelib"))
            {
                Console.WriteLine($"{name}");
                foreach (var k2 in RegLib.GetSubKeyNames(name))
                {
                    var match = Regex.Match(k2, versionPattern, RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        Console.WriteLine($"    {k2}");
                        foreach (var k3 in RegLib.GetSubKeyNames(k2))
                        {
                            var flagsStr = "";
                            if (k3.EndsWith("flags", StringComparison.OrdinalIgnoreCase))
                            {
                                flagsStr = $":{RegLib.GetDefaultValue(k3)}";
                                Console.WriteLine($"        {k3}{flagsStr}");
                            }
                            else
                            {
                                var possibleLCID = RegLib.GetLastSegment(k3);
                                if (possibleLCID.All(char.IsDigit)) // the LCID
                                {
                                    Console.WriteLine($"        {k3}");
                                    foreach (var platform in RegLib.GetSubKeyNames(k3))
                                    {
                                        Console.WriteLine($"            platform:{RegLib.GetLastSegment(platform)}");
                                        foreach ((string key, string value) in RegLib.GetValues(platform))
                                        {
                                            Console.WriteLine($"                {key} {value}");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
