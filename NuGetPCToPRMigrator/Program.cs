using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

using HWND = System.IntPtr;

namespace Ceridian
{
    using EnvDTE;

    using Process = System.Diagnostics.Process;
    using Thread = System.Threading.Thread;

    public static class Program
    {
        //private const string DEVENV = @"C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\devenv.exe";
        private const string DEVENV = @"C:\Program Files\Microsoft Visual Studio\2022\Professional\Common7\IDE\devenv.exe";
        private const string DEFAULT_BROWSER = "chrome.exe";

        private delegate bool EnumWindowsProc(HWND hWnd, int lParam);

        [DllImport("USER32.DLL")]
        private static extern bool EnumWindows(EnumWindowsProc enumFunc, int lParam);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowText(HWND hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowTextLength(HWND hWnd);

        [DllImport("USER32.DLL")]
        private static extern bool IsWindowVisible(HWND hWnd);

        [DllImport("USER32.DLL")]
        private static extern HWND GetShellWindow();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(HWND hWnd);

        public static HWND GetDialogHandle()
        {
            HWND shellWindow = GetShellWindow();
            HWND handle = HWND.Zero;

            EnumWindows(delegate (HWND hWnd, int lParam)
            {
                const string TITLE_PREFIX = "Migrate NuGet format to PackageReference - ";

                if (hWnd == shellWindow || !IsWindowVisible(hWnd))
                {
                    return true;
                }

                int length = GetWindowTextLength(hWnd);
                if (length < TITLE_PREFIX.Length + 1)
                {
                    return true;
                }

                var builder = new StringBuilder(length);
                GetWindowText(hWnd, builder, length + 1);

                var title = builder.ToString();
                if (title.StartsWith(TITLE_PREFIX))
                {
                    handle = hWnd;
                    return false;
                }
                return true;
            }, 0);

            return handle;
        }

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine($"Usage: {Process.GetCurrentProcess().ProcessName} SlnFile1 SlnFile2 ...");
                return;
            }

            foreach (var slnFilePath in args.Select(Path.GetFullPath))
            {
                WriteConsole(slnFilePath, ConsoleColor.Yellow);

                Process devenv = null;
                var dte = RunningVSInstanceFinder.Find(slnFilePath);
                if (dte == null)
                {
                    devenv = Process.Start(DEVENV, slnFilePath);
                    do
                    {
                        Thread.Sleep(1000);
                        dte = RunningVSInstanceFinder.Find(slnFilePath);
                    }
                    while (dte == null);
                }
                var ok = MigrateSolution(dte, slnFilePath);
                if (devenv != null && ok)
                {
                    devenv.Kill();
                }
            }
        }

        private static T ExecuteWithRetry<T>(Func<T> func)
        {
            int retryCount = 10;
            for (; ; )
            {
                try
                {
                    return func();
                }
                catch
                {
                    --retryCount;
                    if (retryCount == 0)
                    {
                        throw;
                    }
                    Thread.Sleep(1000);
                }
            }
        }

        private static bool StartProjectMigration(EnvDTE80.DTE2 dte, int i, int total, string solutionName, string projectRelativeFilePath, string projectRelativeSlnPath)
        {
            Console.WriteLine($"[{i}/{total}] {projectRelativeFilePath}");
            var projectItem = ExecuteWithRetry(() => GetItemWorkaround(dte.ToolWindows.SolutionExplorer, $@"{solutionName}\{projectRelativeSlnPath}"));
            if (ProjectUsesPackagesConfig(projectItem))
            {
                projectItem.Select(EnvDTE.vsUISelectionType.vsUISelectionTypeSelect);
                dte.ExecuteCommand("ClassViewContextMenus.ClassViewProject.Migratepackages.configtoPackageReference");
            }
            else
            {
                throw new ProjectDoesNotUsePackagesConfigException();
            }
            return true; //dummy return value just to satisfy ExecuteWithRetry
        }

        //select each element of the project item and its parents in the UI
        private static UIHierarchyItem GetItemWorkaround(UIHierarchy uh, string projectFullSlnPath)
        {
            var items = projectFullSlnPath.Split('\\');
            UIHierarchyItem uii = null;
            foreach (var s in items)
            {
                uii = uii == null ? uh.GetItem(s) : uii.UIHierarchyItems.Item(s);
                uii.Select(vsUISelectionType.vsUISelectionTypeSelect);
                if (uii.UIHierarchyItems.Count > 0)
                {
                    uii.UIHierarchyItems.Expanded = true;
                }
                if ((uii.Object is Project project) && (project.Kind == Constants.vsProjectKindSolutionItems))
                {
                    uii.UIHierarchyItems.Expanded = true;
                }
            }

            return uii;
        }

        private static bool ProjectUsesPackagesConfig(UIHierarchyItem uii)
        {
            if (uii.Object is Project project)
            {
                return ProjectUsesPackagesConfig(project, true);
            }
            return false;
        }

        //Check if project uses packages.config
        //if waitForFile is true, and the files does not exist, this may be a case where 2 or more project files are in
        //the same directory and a previous project has already been migrated, which would have deleted the file.
        // in that case, give the user the opportunity to restore the file or quit.
        private static bool ProjectUsesPackagesConfig(Project project, bool waitForFile = false)
        {
            var files = GetProjectItemsRecursively(project.ProjectItems);
            bool usesPackagesConfig = files.Any(f => f.EndsWith("/packages.config", StringComparison.InvariantCultureIgnoreCase));

            if (waitForFile)
            {
                var configFile = files.FirstOrDefault(f => f.EndsWith("/packages.config", StringComparison.InvariantCultureIgnoreCase));
                if (!File.Exists(configFile))
                {
                    var dlgResult = MessageBox.Show($"{configFile} does not exist.`n`nRestore the file and press OK to continue or press Cancel to abort",
                        "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                    if (dlgResult == DialogResult.Cancel)
                        throw new OperationCanceledException();
                }
            }

            return usesPackagesConfig;
        }

        //https://gist.github.com/niaher/bfa87f0aeda1204091fe
        private static List<string> GetProjectItemsRecursively(EnvDTE.ProjectItems items)
        {
            var ret = new List<string>();
            if (items == null) return ret;

            foreach (EnvDTE.ProjectItem item in items)
            {
                string result = item.FileNames[1].Replace("\\", "/");

                // If not folder.
                if (result[result.Length - 1] != '\\')
                {
                    ret.Add(result);
                }

                ret.AddRange(GetProjectItemsRecursively(item.ProjectItems));
            }

            return ret;
        }

        private static T WaitForResult<T>(Func<T> func, Predicate<T> isReady)
        {
            T res;
            while (!isReady(res = func()))
            {
                Thread.Sleep(1000);
            }
            return res;
        }

        private static object InitializeNuGetPackageManager(EnvDTE80.DTE2 dte)
        {
            dte.ExecuteCommand("View.PackageManagerConsole");
            return null;
        }

        private static string TryReadAllText(string filePath)
        {
            try
            {
                return File.ReadAllText(filePath);
            }
            catch
            {
                return null;
            }
        }

        private static string Save(EnvDTE.Project project)
        {
            project.Save();
            return null;
        }

        private static Func<string, string> GetHtmlReportFilePathFunc(string slnFilePath)
        {
            var migrationBackupFolderPath = $@"{slnFilePath}\..\MigrationBackup";
            HashSet<string> done;
            if (Directory.Exists(migrationBackupFolderPath))
            {
                done = Directory
                    .EnumerateDirectories(migrationBackupFolderPath)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
            }
            else
            {
                done = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            return projectRelativeSlnPath => Directory.Exists(migrationBackupFolderPath) ? Directory
                .EnumerateDirectories(migrationBackupFolderPath)
                .Where(path => !done.Contains(path) && File.Exists($@"{path}\{projectRelativeSlnPath}\NuGetUpgradeLog.html"))
                .Select(path => $@"{path}\{projectRelativeSlnPath}\NuGetUpgradeLog.html")
                .FirstOrDefault() : null;
        }

        private static bool MigrateSolution(EnvDTE80.DTE2 dte, string slnFilePath)
        {
            bool ok = true;
            var getHTMLReportFilePath = GetHtmlReportFilePathFunc(slnFilePath);

            var solutionName = Path.GetFileNameWithoutExtension(slnFilePath);

            // enumerate the projects in the solution, and get the relative solution path (which may not be the same as the the relative file path)
            IEnumerable<Tuple<string, EnvDTE.Project>> FindProjectsRecursive(EnvDTE.Project p, string prefix)
            {
                if (IsClassLibraryWithPackagesConfig(p))
                    yield return new Tuple<string, EnvDTE.Project>(prefix + p.Name, p);
                if (p.Kind == FolderKind)
                {
                    foreach (EnvDTE.ProjectItem pProjectItem in p.ProjectItems)
                    {
                        if (pProjectItem.SubProject != null && (pProjectItem.SubProject.Kind == ProjectKind || pProjectItem.SubProject.Kind == FolderKind))
                        {
                            var recursePrefix = string.IsNullOrEmpty(prefix) ? "" : prefix;
                            recursePrefix += $"{p.Name}\\";
                            var recurseResult = FindProjectsRecursive(pProjectItem.SubProject, recursePrefix);
                            foreach (var item in recurseResult)
                            {
                                yield return item;
                            }
                        }
                    }
                }
            }

            ExecuteWithRetry(() => InitializeNuGetPackageManager(dte));
            WriteConsole("Enumerating projects in solution.", ConsoleColor.Green);
            var projects = ExecuteWithRetry(() => dte.Solution.Projects.Cast<EnvDTE.Project>().SelectMany(x => FindProjectsRecursive(x, null)).ToList());
            int i = 0;
            foreach (var projectInfo in projects)
            {
                var projectSlnPath = projectInfo.Item1;
                var project = projectInfo.Item2;
                try
                {
                    ++i;
                    var projectFullName = ExecuteWithRetry(() => project.UniqueName);
                    var projectName = project.Name;
                    try
                    {
                        ExecuteWithRetry(() => StartProjectMigration(dte, i, projects.Count, solutionName, projectName, projectSlnPath));
                        var handle = WaitForResult(GetDialogHandle, h => h != HWND.Zero);

                        SetForegroundWindow(handle);
                        SendKeys.SendWait("{ENTER}");

                        int j = 0;
                        var packagesConfigFilePath = $@"{projectFullName}\..\packages.config";
                        string htmlReportFilePath = WaitForResult(() => getHTMLReportFilePath(projectSlnPath),
                            path => path != null || ++j >= 60 && !File.Exists(packagesConfigFilePath));

                        if (htmlReportFilePath != null)
                        {
                            string htmlReport = WaitForResult(() => TryReadAllText(htmlReportFilePath), content => content != null);
                            //Process.Start("taskkill", "/f /im " + DEFAULT_BROWSER);
                            if (htmlReport.Contains("No issues were found."))
                            {
                                Thread.Sleep(1500); // give the browser a moment to show the page before we delete it
                                Directory.Delete($@"{htmlReportFilePath}\..\..", true);
                                WriteConsole("\tNo issues were found.", ConsoleColor.Green);
                            }
                            else
                            {
                                WriteConsole("\tOne or more issues were found.", ConsoleColor.Red);
                            }
                        }
                        ExecuteWithRetry(() => Save(project));
                    }
                    catch (ProjectDoesNotUsePackagesConfigException)
                    {
                        WriteConsole("\tDoes not use packages.config.", ConsoleColor.Green);
                    }
                }
                catch (Exception e)
                {
                    WriteConsole(e.ToString(), ConsoleColor.Red);
                    ok = false;
                }
            }

            return ok;
        }

        private static void WriteConsole(string text, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(text);
            Console.ResetColor();
        }

        private static bool IsWebApplication(EnvDTE.Project project) =>
            project.Properties.Cast<EnvDTE.Property>().Any(p => p.Name == "WebApplication.UseIISExpress");

        const string FolderKind = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";
        const string ProjectKind = "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}";

        private static bool IsClassLibraryWithPackagesConfig(EnvDTE.Project p) =>
            p.Kind == "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}" &&
            ProjectUsesPackagesConfig(p) &&
            !IsWebApplication(p);
    }

    public class ProjectDoesNotUsePackagesConfigException : Exception
    {
    }
}