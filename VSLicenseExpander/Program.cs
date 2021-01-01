using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;
using Microsoft.Win32;

namespace VSLicenseExpander
{
    class Program
    {
        private const string VS2017 = @"Licenses\5C505A59-E312-4B89-9508-E162F8150517\";
        private const string VS2019 = @"Licenses\41717607-F34E-432C-A138-A3CFD7E25CDA\";
        private static readonly string[] subPath2017 = { "08860", "08862", "08863", "08864", "08865", "08866", "08867", "08868", "08869", "08871", "08873", "08875", "08876", "08877", "08878", "0bcad" };
        private static readonly string[] subPath2019 = { "09278" };

        private static int VSEdtion;

        static void Main(string[] args)
        {
            WriteOptions();
            var option = ConsoleKey.A;

            while (option != ConsoleKey.Z)
            {
                option = Console.ReadKey().Key;
                StartAction(option);
            }
        }

        private static void WriteOptions()
        {
            Console.WriteLine("What do you want to do?");
            Console.WriteLine("1. Clear console.");
            Console.WriteLine("2. Read all subkeys to find one for your VS.");
            Console.WriteLine("3. Set new VS expiration date.");
            Console.WriteLine("Z. Exit.");
            Console.WriteLine();
        }

        private static void StartAction(ConsoleKey option)
        {
            if (option == ConsoleKey.D1)
            {
                Console.Clear();
                WriteOptions();
            }
            else if (option == ConsoleKey.D2)
            {
                ReadAllSubKeysAndFindDates();
            }
            else if (option == ConsoleKey.D3)
            {
                SetNewExpirationDate();
            }
        }

        private static void ReadAllSubKeysAndFindDates()
        {
            Console.WriteLine();
            Console.WriteLine("2. Read all subkeys to find one for your VS.");

            var def = new DateTime(1000, 1, 1);
            var result = new Dictionary<string, List<string>>();

            ChooseVSEdition();
            var subPath = VSEdtion == 2017 ? subPath2017 : subPath2019;
            var path = VSEdtion == 2017 ? VS2017 : VS2019;

            foreach (var sp in subPath)
            {
                result.Add(sp, new List<string>());

                var lb = UnprotectSubKey(path + sp);

                if (lb != null)
                {
                    for (int i = 0; i < lb.Length - 5; i++)
                    {
                        var a = GetSubArray(lb, i);
                        var date = ConvertFromBinaryDate(a);

                        if (date != def)
                        {
                            result[sp].Add(i + ", Date:" + date.ToString("yyyy MM dd"));
                        }
                    }
                }
            }

            Console.WriteLine("Keys that contain dates:");
            foreach (var kv in result)
            {
                if (kv.Value.Count != 0)
                {
                    Console.WriteLine(kv.Key);
                    foreach (var item in kv.Value)
                    {
                        Console.WriteLine("DSP:" + item);
                    }
                }
            }
        }

        private static void ChooseVSEdition()
        {
            Console.WriteLine();
            Console.WriteLine("Which VS edition?");
            Console.WriteLine(" a) VS 2017");
            Console.WriteLine(" b) VS 2019");
            var key = ConsoleKey.Z;

            while (key != ConsoleKey.A && key != ConsoleKey.B)
            {
                key = Console.ReadKey().Key;

                if (key == ConsoleKey.A)
                {
                    VSEdtion = 2017;
                    Console.WriteLine(" You choose " + VSEdtion);
                }
                else if (key == ConsoleKey.B)
                {
                    VSEdtion = 2019;
                    Console.WriteLine(" You choose " + VSEdtion);
                }
            }
        }

        private static void SetNewExpirationDate()
        {
            Console.WriteLine();
            Console.WriteLine("3. Set new VS expiration date.");

            ChooseVSEdition();

            Console.WriteLine(" [-] insert subKey");
            var subKey = Console.ReadLine();
            if (subKey.Length != 5)
            {
                Console.WriteLine(" ! subKey is invalid !");
                return;
            }

            var subPath = VSEdtion == 2017 ? subPath2017 : subPath2019;
            if (!subPath.Contains(subKey))
            {
                Console.WriteLine(" ! This VS " + VSEdtion + " does not contain that key!");
                return;
            }

            Console.WriteLine(" [-] insert date starting position (DSP)");
            var start = Console.ReadLine();
            var sOk = int.TryParse(start, out var s);
            if (!sOk)
            {
                Console.WriteLine(" ! Its not a number !");
                return;
            }

            Console.WriteLine(" [-] insert days to add starting from today");
            var days = Console.ReadLine();
            var dOk = int.TryParse(days, out var d);
            if (!dOk)
            {
                Console.WriteLine(" ! Its not a number !");
                return;
            }

            //ChangeDate(208, "08862", 10);
            ChangeDate(s, subKey, d);
        }

        private static bool ChangeDate(int s, string subKey, int e)
        {
            var newDate = DateTime.UtcNow.AddDays(e);
            var binaryDate = ConvertToBinaryDate(newDate);

            var path = VSEdtion == 2017 ? VS2017 : VS2019;
            using (var rk = GetRegistryKey(path + subKey))
            {
                if (rk == null)
                {
                    Console.WriteLine("This key does not exist: " + subKey);
                }

                var lb = UnprotectSubKey(rk);
                lb = SubtituteBinaryData(lb, binaryDate, s);
                ValidateConvertion(lb, s, newDate);

                var success = ProtectSubKey(rk, lb);

                if (success)
                {
                    Console.WriteLine("Done: extended by " + e + " days, result: " + newDate.ToShortDateString());
                }
            }            

            return true;
        }

        private static byte[] SubtituteBinaryData(byte[] lb, byte[] data, int s)
        {
            for (int i = 0; i < 6; i++)
            {
                System.Diagnostics.Trace.WriteLine(lb[s + i] + "," + data[i]);
                lb[s + i] = data[i];
            }

            return lb;
        }

        private static void ValidateConvertion(byte[] lb, int s, DateTime newDate)
        {
            var a = GetSubArray(lb, s);
            var date = ConvertFromBinaryDate(a);

            if (newDate.Year != date.Year || newDate.Month != date.Month || newDate.Day != date.Day)
            {
                throw new Exception("Binary data substitution failed.");
            }
        }

        private static RegistryKey GetRegistryKey(string subKey)
        {
            using (var subReg = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Default))
            {
                return subReg.OpenSubKey(subKey, true);
            }
        }

        private static byte[] UnprotectSubKey(string subKey)
        {
            //Console.WriteLine(subKey);
            using (var registryKey = GetRegistryKey(subKey))
            {
                var val = registryKey.GetValue(null) as byte[];

                if (val != null)
                {
                    return ProtectedData.Unprotect(val, null, DataProtectionScope.LocalMachine);
                }
                else
                {
                    return null;
                }
            }
        }

        private static byte[] UnprotectSubKey(RegistryKey registryKey)
        {
            var val = registryKey.GetValue(null) as byte[];

            if (val != null)
            {
                return ProtectedData.Unprotect(val, null, DataProtectionScope.LocalMachine);
            }
            else
            {
                return null;
            }
        }

        private static bool ProtectSubKey(RegistryKey registryKey, byte[] lb)
        {
#if DEBUG
			Console.WriteLine("You're in debug mode. Keys won't be changed.");
            return true;
#endif
            try
            {
                var value = ProtectedData.Protect(lb, null, DataProtectionScope.LocalMachine);
                registryKey.SetValue(null, value, RegistryValueKind.Binary);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return false;
            }

            return true;
        }

        private static DateTime ConvertFromBinaryDate(byte[] array)
        {
            var y = System.BitConverter.ToInt16(array, 0);
            var m = System.BitConverter.ToInt16(array, 2);
            var d = System.BitConverter.ToInt16(array, 4);

            if ((y <= 2000 || y > 2020) || (m <= 0 || m > 12) || (d <= 0 || d > 31))
            {
                return new DateTime(1000, 1, 1);
            }

            return new DateTime(y, m, d);
        }

        private static byte[] ConvertToBinaryDate(DateTime newDate)
        {
            var y = System.BitConverter.GetBytes((UInt16)newDate.Year);
            var m = System.BitConverter.GetBytes((UInt16)newDate.Month);
            var d = System.BitConverter.GetBytes((UInt16)newDate.Day);

            var tmp = new List<byte>();
            tmp.AddRange(y);
            tmp.AddRange(m);
            tmp.AddRange(d);

            var result = tmp.ToArray();
            var date = ConvertFromBinaryDate(result);

            if (newDate.Year != date.Year || newDate.Month != date.Month || newDate.Day != date.Day)
            {
                Console.WriteLine("Expected: " + newDate.ToShortDateString() + ", actualy: " + date.ToShortDateString());
                throw new InvalidOperationException("Convertion of date to binary format failed.");
            }

            if (result.Length != 6)
            {
                throw new Exception("Binary date to long or to short.");
            }

            return result;
        }

        private static byte[] GetSubArray(byte[] a, int start)
        {
            if (start + 5 < a.Length)
            {
                return new byte[] { a[start], a[start + 1], a[start + 2], a[start + 3], a[start + 4], a[start + 5] };
            }

            throw new Exception("End of array");
        }
    }
}
