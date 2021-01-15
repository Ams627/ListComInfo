using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace ListComInfo
{
    class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                var opts = args.Where(x => x[0] == '-');
                var allOpts = args.SelectMany(x => x).ToHashSet();
                var is32Bit = allOpts.Contains('3');
                var help = allOpts.Contains('h') || allOpts.Contains('?') || (args.Any() && args[0] == "--help");

                if (help)
                {
                    PrintUsageAndExit();
                }

                PrintTypeLibs(is32Bit);
                PrintClsIds(is32Bit);
            }
            catch (Exception ex)
            {
                var fullname = System.Reflection.Assembly.GetEntryAssembly().Location;
                var progname = Path.GetFileNameWithoutExtension(fullname);
                Console.Error.WriteLine($"{progname} Error: {ex.Message}");
            }

        }

        private static void PrintUsageAndExit()
        {
            Console.Error.WriteLine("Usage:");
            Console.Error.WriteLine("    listcominfo");
            Console.Error.WriteLine("       lists the 64 bit COM entries in the registry.");
            Console.Error.WriteLine();
            Console.Error.WriteLine("    listcominfo -3");
            Console.Error.WriteLine("       lists the 32 bit COM entries in the registry.");
        }

        private static void PrintClsIds(bool is32bit = false)
        {
            var view = is32bit ? RegistryView.Registry32 : RegistryView.Registry64;
            foreach (var name in RegLib.GetSubKeyNames(@"HKCR\clsid", view))
            {
                Console.WriteLine($"{name}");
                foreach (var k2 in RegLib.GetSubKeyNames(name, view))
                {
                    if (k2.EndsWith("InprocServer32", StringComparison.InvariantCultureIgnoreCase))
                    {
                        var dll = RegLib.GetDefaultValue(k2, view);
                        Console.WriteLine($"{k2} {dll}");
                    }
                }
            }
        }

        private static void PrintTypeLibs(bool is32bit = false)
        {
            var versionPattern = @"\\[a-z0-9]{1,3}\.[a-z0-9]{1,3}$";

            var view = is32bit ? RegistryView.Registry32 : RegistryView.Registry64;
            foreach (var name in RegLib.GetSubKeyNames(@"HKCR\typelib", view))
            {
                Console.WriteLine($"{name}");
                foreach (var k2 in RegLib.GetSubKeyNames(name, view))
                {
                    var match = Regex.Match(k2, versionPattern, RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        Console.WriteLine($"    {k2}");
                        foreach (var k3 in RegLib.GetSubKeyNames(k2, view))
                        {
                            var flagsStr = "";
                            if (k3.EndsWith("flags", StringComparison.OrdinalIgnoreCase))
                            {
                                flagsStr = $":{RegLib.GetDefaultValue(k3, view)}";
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
                                        foreach ((string key, string value) in RegLib.GetValues(platform, view))
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
