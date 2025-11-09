using System.Diagnostics;

namespace Installation_Finalizer
{
    internal class Program
    {
        static int page = 0;

        static void Main(string[] args)
        {
            Console.SetWindowSize(50, 20);
            StartMenu();
        }

        static void StartMenu()
        {
            page = 0; // Set the page "Start Menu"

            Console.Clear();

            Console.WriteLine(">> Please select your operation!\n");
            Console.WriteLine("========================================");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("||     Windows 11 Toolbox Utility     ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("|| [1] System Tools                   ||");
            Console.WriteLine("|| [2] Interface & Behavior           ||");
            Console.WriteLine("|| [3] Power & System Tweaks          ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("|| [0] Exit                           ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("========================================");
            Console.Write(">> ");

            ConsoleKeyInfo keyInfo = Console.ReadKey();
            char key = keyInfo.KeyChar;
            
            switch(key)
            {
                // Open the "System Tools Menu"
                case '1':
                    SystemTools();
                    break;
                // Open the "Interface And Behavior"
                case '2':
                    InterfaceAndBehavior();
                    break;
                // Open the "Power And System"
                case '3':
                    PowerAndSystem();
                    break;
                // Exit the application
                case '0':
                    Environment.Exit(0);
                    break;
                // Check for invalid input
                default:
                    InvalidInput(key);
                    break;
            }
        }

        static void SystemTools()
        {
            page = 1; // Set the page "Windows Operations"
            Console.Clear();

            Console.WriteLine(">> Please select your operation!\n");
            Console.WriteLine("========================================");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("||            System Tools            ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("|| [1] Activate Windows! (MassGrave)  ||");
            Console.WriteLine("|| [2] Install Microsoft Office!      ||");
            Console.WriteLine("|| [3] Show System Info               ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("|| [0] Back                           ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("========================================");
            Console.Write(">> ");

            ConsoleKeyInfo keyInfo = Console.ReadKey();
            char key = keyInfo.KeyChar;

            switch (key)
            {
                // Opens windows activation script
                case '1':
                    Console.Clear();
                    Console.WriteLine("\nRedirecting...");
                    UseShell("irm https://get.activated.win | iex");
                    ResetConsole();
                    break;
                // Redirects to office install page
                case '2':
                    Console.Clear();
                    Console.WriteLine("\nRedirecting...");
                    OpenInWeb("https://gravesoft.dev/office_c2r_links");
                    ResetConsole();
                    break;
                // Shows system info
                case '3':
                    Console.Clear();
                    SystemInfo();
                    ResetConsole();
                    break;
                // Return to main menu
                case '0':
                    StartMenu();
                    break;
                // Check for invalid input
                default:
                    InvalidInput(key);
                    break;
            }
        }

        static void InterfaceAndBehavior()
        {
            page = 2; // Set the page "Interface And Behavior"

            Console.Clear();

            Console.WriteLine(">> Please select your operation!\n");
            Console.WriteLine("========================================");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("||        Interface & Behavior        ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("|| [1] Restore Old Context Menu       ||");
            Console.WriteLine("|| [2] Restore Modern Context Menu    ||");
            Console.WriteLine("|| [3] Prevent ms-gamebar pop-up      ||");
            Console.WriteLine("|| [4] Prevent ms-gamingoverlay pop-up||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("|| [0] Back                           ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("========================================");
            Console.Write(">> ");

            ConsoleKeyInfo keyInfo = Console.ReadKey();
            char key = keyInfo.KeyChar;

            switch (key)
            {
                // Restores old context menu
                case '1':
                    Console.Clear();
                    RegOperation("add \"HKCU\\Software\\Classes\\CLSID\\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\\InprocServer32\" /f /ve");
                    RestartExplorer();
                    ResetConsole();
                    break;
                // Restores modern context menu
                case '2':
                    Console.Clear();
                    RegOperation("delete \"HKCU\\Software\\Classes\\CLSID\\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\" /f");
                    RestartExplorer();
                    ResetConsole();
                    break;
                // Prevent ms-gamebar pop-up
                case '3':
                    Console.Clear();
                    RegOperation("add \"HKCU\\Software\\Classes\\ms-gamebar\" /f /ve /d \"URL:ms-gamebar\"");
                    RegOperation("add \"HKCU\\Software\\Classes\\ms-gamebar\" /f /v \"URL Protocol\" /d \"\"");
                    RegOperation("add \"HKCU\\Software\\Classes\\ms-gamebar\" /f /v \"NoOpenWith\" /d \"\"");
                    RegOperation("add \"HKCU\\Software\\Classes\\ms-gamebar\\shell\\open\\command\" /f /ve /d \"%SystemRoot%\\System32\\systray.exe\"");
                    RegOperation("add \"HKCU\\Software\\Classes\\ms-gamebarservices\" /f /ve /d \"URL:ms-gamebarservices\"");
                    RegOperation("add \"HKCU\\Software\\Classes\\ms-gamebarservices\" /f /v \"URL Protocol\" /d \"\"");
                    RegOperation("add \"HKCU\\Software\\Classes\\ms-gamebarservices\" /f /v \"NoOpenWith\" /d \"\"");
                    RegOperation("add \"HKCU\\Software\\Classes\\ms-gamebarservices\\shell\\open\\command\" /f /ve /d \"%SystemRoot%\\System32\\systray.exe\"");
                    ResetConsole();
                    break;
                // Prevent ms-gameingoverlay pop-up
                case '4':
                    Console.Clear();
                    RegOperation("add \"HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\GameDVR\" /f /t REG_DWORD /v \"AppCaptureEnabled\" /d 0");
                    RegOperation("add \"HKEY_CURRENT_USER\\System\\GameConfigStore\" /f /t REG_DWORD /v \"GameDVR_Enabled\" /d 0");
                    ResetConsole();
                    break;
                // Return to main menu
                case '0':
                    StartMenu();
                    break;
                // Check for invalid input
                default:
                    InvalidInput(key);
                    break;
            }
        }

        static void PowerAndSystem()
        {
            page = 3; // Set the page "Power And System"

            Console.Clear();

            Console.WriteLine(">> Please select your operation!\n");
            Console.WriteLine("========================================");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("||       Power & System Tweaks        ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("|| [1] Enable Power Sleep State (S3)  ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("|| [0] Exit                           ||");
            Console.WriteLine("||                                    ||");
            Console.WriteLine("========================================");
            Console.Write(">> ");

            ConsoleKeyInfo keyInfo = Console.ReadKey();
            char key = keyInfo.KeyChar;

            switch (key)
            {
                // Changes power sleep state to (S3)
                case '1':
                    Console.Clear();
                    RunRegWithElevation();
                    ResetConsole();
                    break;
                // Return to main menu
                case '0':
                    StartMenu();
                    break;
                // Check for invalid input
                default:
                    InvalidInput(key);
                    break;
            }
        }

        static void RestartExplorer()
        {
            Process.Start("cmd.exe", "/c taskkill /f /im explorer.exe & start explorer.exe");
        }

        // It works... I think
        static void RunRegWithElevation()
        {
            var args = "add \"HKLM\\SYSTEM\\CurrentControlSet\\Control\\Power\" /v PlatformAoAcOverride /t REG_DWORD /d 0 /f";

            var psi = new ProcessStartInfo
            {
                FileName = "reg.exe",
                Arguments = args,
                UseShellExecute = true,
                Verb = "runas"
            };

            Process.Start(psi);
        }

        static void SystemInfo()
        {
            Console.SetWindowSize(200, 60);

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
                CreateNoWindow = false
            };

            var process = Process.Start(psi);
            process.WaitForExit();

            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            Console.WriteLine("Output: " + output);
            Console.WriteLine("Error: " + error);
        }

        static void OpenInWeb(string link)
        {
            var psi = new ProcessStartInfo
            {
                FileName = link,
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

        static void InvalidInput(char input)
        {
            Console.Clear();
            Console.WriteLine($"\n'{input}' is not valid! Please try again!");
            Console.WriteLine("Press any key to return...");
            Console.ReadKey();
            ReturnToCurrentPage();
        }

        static void ResetConsole()
        {
            Console.WriteLine("Operation Done!");
            Console.Write("\nPress any key to return...");
            Console.ReadKey();
            ReturnToCurrentPage();
        }

        static void ReturnToCurrentPage()
        {
            Console.SetWindowSize(50, 20);

            switch (page)
            {
                case 0:
                    StartMenu();
                    break;
                case 1:
                    SystemTools();
                    break;
                case 2:
                    InterfaceAndBehavior();
                    break;
            }
        }

        /*
        static int ParseString(string input)
        {
            int number = 0;
            if(!Int32.TryParse(input, out number))
            {
                return number;
            }
            else
            {
                return number;
            }
        } 
        */
    }
}
