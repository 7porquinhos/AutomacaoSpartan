using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spartan.BLL.Util.Extensao
{
    public static class ProcessExtension
    {
        public static void ProcessKill(string ProcessName)
        {
            foreach (var process in Process.GetProcessesByName(ProcessName))
            {
                process.Kill();
            }
        }

        public static void ProcessKill(List<string> ListProcessName)
        {
            foreach (var ProcessName in ListProcessName)
            {
                foreach (var process in Process.GetProcessesByName(ProcessName))
                {
                    process.Kill();
                }
            }
            
        }
    }
}
