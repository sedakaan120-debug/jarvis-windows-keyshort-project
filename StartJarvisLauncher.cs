using System;
using System.Diagnostics;
using System.IO;

class StartJarvisLauncher
{
    static int Main()
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string python = Path.Combine(baseDir, ".venv", "Scripts", "python.exe");
        if (!File.Exists(python))
        {
            python = "python";
        }

        string mainPy = Path.Combine(baseDir, "main.py");
        if (!File.Exists(mainPy))
        {
            Console.Error.WriteLine("main.py bulunamadi: " + mainPy);
            Console.ReadKey();
            return 1;
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = python,
            Arguments = "\"" + mainPy + "\"",
            WorkingDirectory = baseDir,
            UseShellExecute = false
        };

        try
        {
            Process process = Process.Start(startInfo);
            process.WaitForExit();
            return process.ExitCode;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("JARVIS baslatilamadi: " + ex.Message);
            Console.ReadKey();
            return 1;
        }
    }
}
