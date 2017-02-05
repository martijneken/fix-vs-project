using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;

namespace FixProject
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Fix
    {
        public const int ComponentCommandId = 0x0100;
        public const int LibraryCommandId = 0x0101;
        public const int NormalCommandId = 0x0102;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("8cb843ec-7aa4-435f-bc30-3c03385091aa");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Fix"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private Fix(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                // Fix as component
                var componentCommand = new CommandID(CommandSet, ComponentCommandId);
                commandService.AddCommand(new MenuCommand((sender, e) => this.FixProject(ProjectType.GHCOMPONENT), componentCommand));

                // Fix as shared library
                var libraryCommand = new CommandID(CommandSet, LibraryCommandId);
                commandService.AddCommand(new MenuCommand((sender, e) => this.FixProject(ProjectType.GHLIBRARY), libraryCommand));
                
                // Fix as regular executable
                var normalCommand = new CommandID(CommandSet, NormalCommandId);
                commandService.AddCommand(new MenuCommand((sender, e) => this.FixProject(ProjectType.NONGHEXE), normalCommand));
            }
        }
        
        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Fix Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new Fix(package);
        }

        private static EnvDTE80.DTE2 GetDTE2()
        {
            return Package.GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE80.DTE2;
        }

        public static string FullPath(string path)
        {
            try
            {
                return Path.GetFullPath(new Uri(path).LocalPath)
                       .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
            catch
            {
                return path;
            }
        }

        enum ProjectType
        {
            GHCOMPONENT,
            GHLIBRARY,
            NONGHEXE,
        };

        private void FixProject(ProjectType type)
        {
            var app = GetDTE2();

            // Find the currently selected project in the solution.
            EnvDTE.UIHierarchy uih = app.ToolWindows.SolutionExplorer;
            EnvDTE.UIHierarchyItem selection = ((Array)uih.SelectedItems).GetValue(0) as EnvDTE.UIHierarchyItem;

            // Find the project object by name.
            // TODO: support nested projects, see http://www.wwwlicious.com/2011/03/29/envdte-getting-all-projects-html/
            EnvDTE.Project project = null;
            foreach (EnvDTE.Project p in app.Solution.Projects)
            {
                if (p == null) continue;
                if (p.Kind == EnvDTE80.ProjectKinds.vsProjectKindSolutionFolder) continue;
                if (p.Name != selection.Name) continue;
                project = p;
                break;
            }
            if (project == null)
            {
                ShowMessage(OLEMSGICON.OLEMSGICON_CRITICAL, "Error! Project not found: " + selection.Name);
                return;
            }
            var vsproject = project.Object as VSLangProj.VSProject;
            if (vsproject == null)
            {
                ShowMessage(OLEMSGICON.OLEMSGICON_CRITICAL, "Error! project is not C#/VB: " + selection.Name);
                return;
            }
            var steps = new List<string>();

            // Set copylocal value on references to neighboring items:
            // - set false on Rhino/Grasshopper DLLs
            // - set true on any others (DLLs and projects)
            // Non-grasshopper projects use the default value of True.
            // Unloaded projects return an empty path, and won't be touched.
            string parentdir = FullPath(Path.GetDirectoryName(project.FullName) + "\\..");
            foreach (VSLangProj.Reference r in vsproject.References)
            {
                bool nearby = FullPath(r.Path).StartsWith(parentdir, StringComparison.OrdinalIgnoreCase);
                if (nearby)
                {
                    bool rhino = r.Path.IndexOf("rhino", StringComparison.OrdinalIgnoreCase) >= 0;
                    bool grass = r.Path.IndexOf("grasshopper", StringComparison.OrdinalIgnoreCase) >= 0;
                    bool copylocal = (!rhino && !grass) || type == ProjectType.NONGHEXE;
                    try { r.CopyLocal = copylocal; } catch { }
                    steps.Add(string.Format("Set CopyLocal = {0} for reference to {1}", copylocal, r.Name));
                }
            }

            // Set post build events:
            // - rename DLL if it's a component
            // - always copy to ..\Output
            var pbsteps = new List<string>();
            if (type == ProjectType.GHCOMPONENT) 
            {
                pbsteps.Add(@"MOVE /Y ""$(TargetPath)"" ""$(TargetDir)$(ProjectName).gha""");
                steps.Add(string.Format("Set PostBuild to rename {0}.dll to {0}.gha", project.Name));
            }
            if (type == ProjectType.GHCOMPONENT || type == ProjectType.GHLIBRARY)
            {
                pbsteps.Add(@"XCOPY /Y /F ""$(TargetDir)*"" ""$(ProjectDir)..\Output\""");
                steps.Add("Set PostBuild to copy output to ..\\Output");
            }
            var postbuild = project.Properties.Item("PostBuildEvent");
            postbuild.Value = string.Join(" && ", pbsteps);
            
            // Find the installation directory for Rhino, preferring 64-bit.
            // See: http://developer.rhino3d.com/guides/cpp/finding-rhino-installation-folder/
            string path = null;
            const string RHINO = @"SOFTWARE\McNeel\Rhinoceros";
            RegistryKey local =
                RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64) ??
                RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
            RegistryKey key = local.OpenSubKey(RHINO);
            if (key != null)
            {
                string[] subs = key.GetSubKeyNames();
                if (subs.Length > 0)
                {
                    // Pick the last version directory, hoping to get the latest / newest install.
                    string version = subs[subs.Length - 1];
                    RegistryKey subkey = local.OpenSubKey(string.Format(@"{0}\{1}\Install", RHINO, version));
                    if (subkey != null)
                    {
                        object o = subkey.GetValue("Path");
                        path = o as string;
                    }
                }
            }
            // Set the project to launch Rhino when debugging.
            if (path != null)
            {
                // Update RHINO environment variable in registry and local process.
                try { Environment.SetEnvironmentVariable("RHINO", path, EnvironmentVariableTarget.User); } catch { }
                try { Environment.SetEnvironmentVariable("RHINO", path, EnvironmentVariableTarget.Process); } catch { }
                steps.Add(string.Format("Set Rhino path to {0}", path));

                // Set Rhino launch action for all configurations.
                bool rhino = type != ProjectType.NONGHEXE;
                foreach (EnvDTE.Configuration config in project.ConfigurationManager)
                {
                    if (rhino)
                    {
                        config.Properties.Item("StartAction").Value = VSLangProj.prjStartAction.prjStartActionProgram;
                        config.Properties.Item("StartProgram").Value = @"$(RHINO)";
                    }
                    else
                    {
                        config.Properties.Item("StartAction").Value = VSLangProj.prjStartAction.prjStartActionProject;
                        config.Properties.Item("StartProgram").Value = null;
                    }
                    config.Properties.Item("StartArguments").Value = null;
                    config.Properties.Item("StartWorkingDirectory").Value = null;
                }
                steps.Add(string.Format("Set debugging action to launch {0}", rhino ? "Rhino" : "Project"));
            }

            // Communicate success.
            string desc;
            switch (type)
            {
                case ProjectType.GHCOMPONENT: desc = "Grasshopper Component"; break;
                case ProjectType.GHLIBRARY: desc = "Grasshopper Shared Library"; break;
                default: desc = "Non-Grasshopper Executable"; break;
            }
            steps.Insert(0, string.Format("Updated project as {0}\r\n", desc));
            ShowMessage(OLEMSGICON.OLEMSGICON_INFO, string.Join("\r\n- ", steps));
        }

        private void ShowMessage(OLEMSGICON icon, string message)
        {
            VsShellUtilities.ShowMessageBox(
                this.ServiceProvider,
                message,
                "",
                icon,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
