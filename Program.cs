using System;
using System.Data;
using System.Diagnostics;
using System.Data.SqlClient;
using System.Security.Principal;
using System.Threading;
using System.Collections.Generic;

namespace EnableRemoting
{
    internal class Command
    {
        private string _value;

        public string Value => _value;

        public Command(string value)
        {
            _value = value;
        }

        public Process Execute()
        {
            ProcessStartInfo processInfo = new ProcessStartInfo();
            processInfo.FileName = @"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe";
            processInfo.Verb = "runas";
            processInfo.Arguments = _value;
            Process process = Process.Start(processInfo);
            while (process.HasExited == false) Thread.Sleep(60);
            return process;
        }
    }
    internal class Program
    {
        static void Main(string[] args)
        {
            SqlConnection connection;
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            List<Command> commands = new List<Command>();
            if (principal.IsInRole(WindowsBuiltInRole.Administrator) == false)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ты низок");
                Console.ResetColor();
                return;
            }
            commands.Add(new Command("Enable-PSRemoting -Force -SkipNetworkProfileCheck"));
            commands.Add(new Command("Set-NetFirewallRule -Name 'WINRM-HTTP-In-TCP' -RemoteAddress Any"));
            commands.Add(new Command("Set-ItemProperty -Path 'HKLM:\\System\\CurrentControlSet\\Control\\Terminal Server' -name fDenyTSConnections -Value 0"));
            foreach (Command command in commands)
            {
                Process process = command.Execute();
                if (process.ExitCode != 0)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"{command.Value} послал нахуй со словами: {process.ExitCode}");
                    Console.ResetColor();
                    return;
                }
            }

            connection = new SqlConnection();
            connection.ConnectionString = @"Server=DB-COL-SMIRNOVTORTURELETMEOUT.mssql.somee.com; Database=DB-COL-SMIRNOVTORTURELETMEOUT; User Id=ah308; Password=3221488228";
            connection.Open();
            if (connection.State != ConnectionState.Open)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ДА БЛЯДЬ Я НЕ МОГУ");
                Console.ResetColor();
                return;
            }
            SqlCommand cmd = connection.CreateCommand();
            cmd.CommandText = @"INSERT INTO [dbo].[valhalla.computer] VALUES (@name)";
            cmd.Parameters.Add("@name", SqlDbType.VarChar, 50).Value = $"{Environment.MachineName}";
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("такой хуй уже есть");
                Console.ResetColor();
                return;
            }
        }
    }
}
