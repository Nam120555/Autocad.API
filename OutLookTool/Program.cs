using System;
using System.Linq;
using System.Windows.Forms;

namespace OutLookTool
{
    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            // Chạy ngầm cho lịch tự động: OutLookTool.exe --export
            if (args.Any(a => a.Equals("--export", StringComparison.OrdinalIgnoreCase)))
                return Headless.Run();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            return 0;
        }
    }
}
