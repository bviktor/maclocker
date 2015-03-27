using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.Configuration;

namespace maclocker
{
    class Program
    {
        public static String GenerateDateString()
        {
            DateTime date = DateTime.Now;
            return date.ToString("yyyy-MM-ddTHH:mm:ss\\\\K");
        }

        public static String GenerateGuid(int mode)
        {
            Guid guid = Guid.NewGuid();

            switch (mode)
            {
                case 1:
                    return guid.ToString("D");

                case 2:
                    string guidString = guid.ToString("N");
                    char[] guidChar = guidString.ToCharArray();
                    String buf = "";

                    for (int i = 0; i < 32; i++)
                    {
                        buf += "0x" + guidChar[i] + guidChar[++i] + " ";
                    }

                    return buf;

                case 3:
                    return guid.ToString("N");

                default:
                    return "";
            }
        }

        public static String GenerateKeyCN()
        {
            return GenerateDateString() + "{" + GenerateGuid(1) + "}";
        }

        public static bool ComputerExists(String ouPath, String objName)
        {
            if (DirectoryEntry.Exists("LDAP://CN=" + objName + "," + ouPath))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool CreateComputer(String ouPath, String objName, bool useOsAuth, String adminUser, String adminPass)
        {
            try
            {
                DirectoryEntry comp;

                if (useOsAuth)
                {
                    comp = new DirectoryEntry("LDAP://" + ouPath);
                }
                else
                {
                    comp = new DirectoryEntry("LDAP://" + ouPath, adminUser, adminPass);
                }
                
                DirectoryEntry group = comp.Children.Add("CN=" + objName, "computer");
                group.Properties["sAmAccountName"].Value = objName;
                group.CommitChanges();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return false;
            }
        }

        public static bool AddRecoveryKey(String ouPath, String objName, String recoveryPassword, bool useOsAuth, String adminUser, String adminPass)
        {
            try
            {
                DirectoryEntry comp;

                if (useOsAuth)
                {
                    comp = new DirectoryEntry("LDAP://CN=" + objName + "," + ouPath);
                }
                else
                {
                    comp = new DirectoryEntry("LDAP://CN=" + objName + "," + ouPath, adminUser, adminPass);
                }

                DirectoryEntry recKey = comp.Children.Add("CN=" + GenerateKeyCN(), "msFVE-RecoveryInformation");
                recKey.Properties["msFVE-RecoveryPassword"].Value = recoveryPassword;
                recKey.Properties["msFVE-RecoveryGuid"].Value = GenerateGuid(2);
                recKey.CommitChanges();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return false;
            }
        }

        public static void WriteDashes()
        {
            Console.Write("--------------------------------------------------------------------------------");
        }

        public static void FinishUp()
        {
            WriteDashes();
            Console.WriteLine("End of execution, press any key to close...");
            Console.ReadKey(true);
        }

        static void PrintHelp()
        {
            Console.WriteLine("MacLocker - utility to save FileVault recovery passwords as BitLocker keys in Active Directory. Requires your AD schema to be prepared to store BitLocker info.");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("\tmaclocker <comp> <key>");
            Console.WriteLine("\tmaclocker <comp> <key> <OU DN>");
            Console.WriteLine("\tmaclocker <comp> <key> <OU DN> <user> <pass>");
            Console.WriteLine();
            Console.WriteLine("\tcomp:\tCN of the computer object");
            Console.WriteLine("\tkey:\tFileVault recovery key to store as BitLocker key");
            Console.WriteLine("\tOU DN:\tdistinguishedName of the OU containing the computer");
            Console.WriteLine("\tuser:\tsAMAccountName of the account to write the changes with");
            Console.WriteLine("\tpass:\tpassword of this account");
            Console.WriteLine();
            Console.WriteLine("Example:");
            Console.WriteLine();
            Console.WriteLine("maclocker mac01 UIER-IASD-U2DM-78DG-NM23-91DT OU=Mac,DC=xy,DC=lan admin1 P@ss");
            Console.WriteLine();
            Console.WriteLine("If OU DN or user/pass are omitted, the App.config values are used.");
            Console.WriteLine("If you set useOsAuth to true in App.config, user and pass are ignored.");
            Console.ReadLine();
        }

        static void Main(string[] args)
        {
            String computerName;
            String recoveryKey;
            String adminUser;
            String adminPass;
            String ouPath;
            bool useOsAuth;
            ConsoleKeyInfo key;

            adminUser = ConfigurationManager.AppSettings["adminUser"];
            adminPass = ConfigurationManager.AppSettings["adminPass"];
            ouPath = ConfigurationManager.AppSettings["ouPath"];

            if (ConfigurationManager.AppSettings["useOsAuth"].Equals("true"))
            {
                useOsAuth = true;
            }
            else
            {
                useOsAuth = false;
            }

            switch (args.Length)
            {
                case 2:
                    computerName = args[0];
                    recoveryKey = args[1];
                    break;
                case 3:
                    computerName = args[0];
                    recoveryKey = args[1];
                    ouPath = args[2];
                    break;
                case 5:
                    computerName = args[0];
                    recoveryKey = args[1];
                    ouPath = args[2];
                    adminUser = args[3];
                    adminPass = args[4];
                    break;
                default:
                    PrintHelp();
                    return;
            }

            Console.WriteLine("DN of the computer object:\tCN=" + computerName + "," + ouPath);
            Console.WriteLine("FileVault recovery password:\t" + recoveryKey);
            Console.WriteLine("Using OS authentication:\t" + useOsAuth.ToString());
            Console.Write("Store this key under this object? [y/n]\t");
            key = Console.ReadKey();
            Console.WriteLine();

            if (!key.Key.Equals(ConsoleKey.Y))
            {
                FinishUp();
                return;
            }

            WriteDashes();

            Console.Write("Checking for computer object: \t\t");

            if (ComputerExists(ouPath, computerName))
            {
                Console.WriteLine("exists");
            }
            else
            {
                Console.WriteLine("does not exist");
                Console.Write("Create computer object? [y/n]\t\t");

                key = Console.ReadKey();
                Console.WriteLine();

                if (key.Key.Equals(ConsoleKey.Y))
                {
                    Console.Write("Creating computer object:\t\t");

                    if (!CreateComputer(ouPath, computerName, useOsAuth, adminUser, adminPass))
                    {
                        Console.WriteLine("failure");
                        FinishUp();
                        return;
                    }
                    else
                    {
                        Console.WriteLine("success");
                    }
                }
                else
	            {
                    FinishUp();
                    return;
	            }
            }

            Console.Write("Adding recovery key:\t\t\t");

            if (AddRecoveryKey(ouPath, computerName, recoveryKey, useOsAuth, adminUser, adminPass))
            {
                Console.WriteLine("success");
            }
            else
            {
                Console.WriteLine("failure");
                FinishUp();
                return;
            }

            FinishUp();
        }
    }
}
