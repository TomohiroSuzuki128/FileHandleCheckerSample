using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Principal;
using System.Threading;

namespace HandleTest
{
    class Program
    {
        static void Main(string[] args)
        {
#if (!DEBUG)
            Thread.GetDomain().SetPrincipalPolicy(PrincipalPolicy.WindowsPrincipal);
            var principal = (WindowsPrincipal)Thread.CurrentPrincipal;

            //管理者権限以外での起動なら, 別プロセスで本アプリを起動する
            if (!principal.IsInRole(WindowsBuiltInRole.Administrator))
            {
                var processStartInfo = new ProcessStartInfo
                {
                    WorkingDirectory = Environment.CurrentDirectory,
                    FileName = Process.GetCurrentProcess().MainModule.FileName,
                    UseShellExecute = true,
                    Verb = "RunAs"
                };

                Process.Start(processStartInfo);

                return;
            }
#endif
            // ファイルハンドル握られたファイルのパス
            // この場合、test.xlsx を開いたままにしておく
            var handledFile = @"C:\temp\test.xlsx";

            try
            {
                File.Move(handledFile, @"C:\temp\test2.xlsx", true);
            }
            catch (IOException)
            {
                var processStartInfo = new ProcessStartInfo
                {
                    WorkingDirectory = @$"%HOMEDRIVE%%HOMEPATH%",
                    FileName = "cmd.exe",
                    UseShellExecute = true,
                    Arguments = @$"/k {AppDomain.CurrentDomain.BaseDirectory}handle\handle.exe {handledFile}",
                    Verb = "RunAs"
                };

                Process.Start(processStartInfo);
            }
            return;

        }
    }
}
