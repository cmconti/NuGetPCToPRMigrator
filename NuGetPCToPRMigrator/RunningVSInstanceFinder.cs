using EnvDTE80;
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Ceridian
{
    public static class RunningVSInstanceFinder
    {
        public static DTE2 Find(string solutionFilePath)
        {
            IntPtr pceltFetched = new IntPtr();
            IMoniker[] rgelt = new IMoniker[1];
            GetRunningObjectTable(0, out IRunningObjectTable table);
            table.EnumRunning(out IEnumMoniker moniker);
            moniker.Reset();
            while (moniker.Next(1, rgelt, pceltFetched) == 0)
            {
                CreateBindCtx(0, out _);
                table.GetObject(rgelt[0], out object obj2);
                var dte = obj2 as DTE2;
                try
                {
                    if (dte?.Name != "Microsoft Visual Studio")
                    {
                        continue;
                    }
                    if (dte.Solution == null)
                    {
                        if (solutionFilePath == null)
                        {
                            return dte;
                        }
                        continue;
                    }
                    if (!string.Equals(dte.Solution.FullName, solutionFilePath, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                }
                catch (Exception exception)
                {
                    Console.Error.WriteLine(exception.Message);
                    continue;
                }
                return dte;
            }
            return null;
        }

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);
    }
}