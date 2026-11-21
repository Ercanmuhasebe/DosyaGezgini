using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

internal static class UpdateInstaller
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool MoveFileEx(string existingFileName, string newFileName, int flags);

    private const int MoveFileReplaceExisting = 0x1;
    private const int MoveFileWriteThrough = 0x8;

    [STAThread]
    private static int Main(string[] args)
    {
        string stagedPath = null;
        try
        {
            if (args == null || args.Length != 1) throw new ArgumentException("Güncelleme parametreleri eksik.");
            string payload = Encoding.UTF8.GetString(Convert.FromBase64String(args[0]));
            string[] values = payload.Split(new char[] { '|' }, 3);
            int processId;
            if (values.Length != 3 || !Int32.TryParse(values[0], out processId))
                throw new ArgumentException("Güncelleme parametreleri geçersiz.");
            string targetPath = Path.GetFullPath(values[1]);
            stagedPath = Path.GetFullPath(values[2]);
            if (!File.Exists(targetPath) || !File.Exists(stagedPath))
                throw new FileNotFoundException("Program veya indirilen güncelleme dosyası bulunamadı.");

            try
            {
                using (Process application = Process.GetProcessById(processId))
                    if (!application.WaitForExit(120000)) throw new TimeoutException("Program kapanmadığı için güncelleme uygulanamadı.");
            }
            catch (ArgumentException) { }

            string replacementPath = targetPath + "." + Guid.NewGuid().ToString("N") + ".new";
            File.Copy(stagedPath, replacementPath, false);
            if (!MoveFileEx(replacementPath, targetPath, MoveFileReplaceExisting | MoveFileWriteThrough))
            {
                int error = Marshal.GetLastWin32Error();
                try { File.Delete(replacementPath); } catch { }
                throw new System.ComponentModel.Win32Exception(error, "Güncelleme dosyası programın üzerine alınamadı.");
            }
            try { File.Delete(stagedPath); } catch { }
            Process.Start(new ProcessStartInfo
            {
                FileName = targetPath,
                WorkingDirectory = Path.GetDirectoryName(targetPath),
                UseShellExecute = true
            });
            return 0;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Güncelleme uygulanamadı. Mevcut program dosyası korunmuştur.\r\n\r\n" + ex.Message,
                "Dosya Gezgini güncellemesi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }
}
