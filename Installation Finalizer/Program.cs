using System;
using System.Diagnostics;

namespace Installation_Finalizer
{
    internal class Program
    {
        static void Main(string[] args)
        {
            StartMenu();
        }

        static void StartMenu()
        {
            Console.WriteLine("Welcome to Windows 11 Beta Edition Wizard!\n");
            Console.WriteLine(">> Please select your operation!\n");
            Console.WriteLine("========================================");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("|| Windows Operations                 ||");
            Console.WriteLine("|| [1] Activate Windows 11 version!   ||");
            Console.WriteLine("|| [2] Install Microsoft Office!      ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("|| Personalization                    ||");
            Console.WriteLine("|| [3] Restore old context menu       ||");
            Console.WriteLine("|| [4] Restore Modern context menu    ||");
            Console.WriteLine("|| [5] Power sleep state Standby (S3) ||");
            Console.WriteLine("|| [6] System info summary            ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("|| [9] Exit                           ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("========================================");
            Console.Write(">> ");

            string selectionInput = Console.ReadLine();
            int selection = ParseString(selectionInput);

            switch (selection)
            {
                // Opens windows activation script
                case 1:
                    Console.WriteLine("\nRedirecting...");
                    UseShell("irm https://get.activated.win | iex");
                    ResetConsole();
                    break;
                // Redirects to office install page
                case 2:
                    Console.WriteLine("\nRedirecting...");
                    RedirectOfficeInstallation();
                    ResetConsole();
                    break;
                // Restores old context menu
                case 3:
                    RegOperation("add \"HKCU\\Software\\Classes\\CLSID\\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\\InprocServer32\" /f /ve");
                    Console.WriteLine("(Restart Required)");
                    ResetConsole();
                    break;
                // Restores modern context menu
                case 4:
                    RegOperation("delete \"HKCU\\Software\\Classes\\CLSID\\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\" /f");
                    ResetConsole();
                    break;
                // Changes power sleep state to (S3)
                case 5:
                    RunRegWithElevation();
                    ResetConsole();
                    break;
                // Shows system info
                case 6:
                    Console.Clear();
                    SystemInfo();
                    ResetConsole();
                    break;
                // Exits app
                case 9:
                    Environment.Exit(0);
                    break;
                // Incorrect Entry
                default:
                    Console.Clear();
                    StartMenu();
                    break;
            }
        }

        static void RunRegWithElevation()
        {
            var args = "add \"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Power\" /v PlatformAoAcOverride /t REG_DWORD /d 0 /f";

            var psi = new ProcessStartInfo
            {
                FileName = "reg.exe",
                Arguments = args,
                UseShellExecute = true,   // elevation için true olmalı
                Verb = "runas"            // UAC prompt
            };

            Process.Start(psi);
        }

        static void SystemInfo()
        {
            string output = "";
            var proc = new ProcessStartInfo("cmd.exe", "/c systeminfo")
            {
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                WorkingDirectory = @"C:\Windows\System32\"
            };
            Process p = Process.Start(proc);
            p.OutputDataReceived += (sender, args1) => { output += args1.Data + Environment.NewLine; };
            p.BeginOutputReadLine();
            p.WaitForExit();
            Console.WriteLine(output);
        }

        static void RegOperation(string command)
        {
            var psi = new ProcessStartInfo
            {
                FileName = "reg.exe",
                Arguments = command,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            var process = Process.Start(psi);
            process.WaitForExit();

            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            Console.WriteLine("Output: " + output);
            Console.WriteLine("Error: " + error);
        }

        static void RedirectOfficeInstallation()
        {
            var psi = new ProcessStartInfo
            {
                FileName = "https://gravesoft.dev/office_c2r_links",
                UseShellExecute = true
            };
            Process.Start(psi);
        }

        static void UseShell(string command)
        {

            var psCommand = command;

            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -Command \"{psCommand}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var proc = Process.Start(psi);
            string stdout = proc.StandardOutput.ReadToEnd();
            string stderr = proc.StandardError.ReadToEnd();
            proc.WaitForExit();

            Console.WriteLine("Out:\n" + stdout);
            if (!string.IsNullOrWhiteSpace(stderr))
                Console.Error.WriteLine("Error:\n" + stderr);

        }

        static int ParseString(string input)
        {
            int number = 0;
            if(Int32.TryParse(input, out number))
            {
                return number;
            }
            else
            {
                Console.WriteLine("\n'{0}' is not valid! Please try again!", input);
                Console.WriteLine("Press any key to try again...");
                Console.ReadKey();
                number = 0;
                return number;
            }
        }

        static void ResetConsole()
        {
            Console.WriteLine("Operation Done!");
            Console.ReadKey();
            Console.Clear();
            StartMenu();
        }
    }
}
