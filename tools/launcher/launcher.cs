using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

[assembly: AssemblyTitle("识文寻表")]
[assembly: AssemblyDescription("框选屏幕文字，在 Excel 中自动查找")]
[assembly: AssemblyProduct("识文寻表")]
[assembly: AssemblyVersion("1.1.0.0")]
[assembly: AssemblyFileVersion("1.1.0.0")]

internal static class Launcher
{
    [STAThread]
    private static void Main()
    {
        try
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string script = Path.Combine(baseDir, "app.py");
            if (!File.Exists(script))
            {
                MessageBox.Show(
                    "找不到程序文件 app.py。\n请将 识文寻表.exe 与 app.py 放在同一个文件夹中。",
                    "识文寻表",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            bool isPyLauncher;
            string python = FindPython(out isPyLauncher);
            if (python == null)
            {
                MessageBox.Show(
                    "未找到 Python 运行环境。\n\n请安装 Python 3.10 或更高版本，并在安装时勾选“Add Python to PATH”。",
                    "识文寻表 · 缺少运行环境",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = python,
                Arguments = (isPyLauncher ? "-3 " : "") + "\"" + script + "\"",
                WorkingDirectory = baseDir,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "启动识文寻表失败：\n" + ex.Message,
                "识文寻表",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private static string FindPython(out bool isPyLauncher)
    {
        isPyLauncher = false;

        string pathVariable = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        String[] pathParts = pathVariable.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string part in pathParts)
        {
            string dir = part.Trim().Trim('"');
            if (dir.Length == 0) continue;
            string pythonw = Path.Combine(dir, "pythonw.exe");
            if (File.Exists(pythonw)) return pythonw;
        }

        foreach (string part in pathParts)
        {
            string dir = part.Trim().Trim('"');
            if (dir.Length == 0) continue;
            string pyw = Path.Combine(dir, "pyw.exe");
            if (File.Exists(pyw))
            {
                isPyLauncher = true;
                return pyw;
            }
        }

        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string pythonRoot = Path.Combine(localAppData, "Programs", "Python");
        if (Directory.Exists(pythonRoot))
        {
            var candidates = new List<string>();
            foreach (string dir in Directory.GetDirectories(pythonRoot, "Python3*"))
            {
                candidates.Add(dir);
            }
            candidates.Sort(StringComparer.OrdinalIgnoreCase);
            candidates.Reverse();
            foreach (string dir in candidates)
            {
                string pythonw = Path.Combine(dir, "pythonw.exe");
                if (File.Exists(pythonw)) return pythonw;
            }
        }

        for (int minor = 20; minor >= 10; minor--)
        {
            string pythonw = Path.Combine(@"C:\", "Python3" + minor, "pythonw.exe");
            if (File.Exists(pythonw)) return pythonw;
        }

        foreach (string part in pathParts)
        {
            string dir = part.Trim().Trim('"');
            if (dir.Length == 0) continue;
            string python = Path.Combine(dir, "python.exe");
            if (File.Exists(python)) return python;
        }

        foreach (string part in pathParts)
        {
            string dir = part.Trim().Trim('"');
            if (dir.Length == 0) continue;
            string py = Path.Combine(dir, "py.exe");
            if (File.Exists(py))
            {
                isPyLauncher = true;
                return py;
            }
        }

        return null;
    }
}
