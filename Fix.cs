using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Linq;

namespace FixProject
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Fix
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

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
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
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

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
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

            // Set launch action for all configurations.
            foreach (EnvDTE.Configuration config in project.ConfigurationManager)
            {
                string path = config.Properties.Item("OutputPath").Value.ToString();
                string projectdir = string.Concat(Enumerable.Repeat("..\\", path.Count(x => x == '\\')));
                config.Properties.Item("StartAction").Value = VSLangProj.prjStartAction.prjStartActionProgram;
                config.Properties.Item("StartProgram").Value = @"C:\WINDOWS\system32\cmd.exe";
                config.Properties.Item("StartArguments").Value = @"/c start_rhino.bat";
                config.Properties.Item("StartWorkingDirectory").Value = projectdir + @"..\Libraries\";
            }

            // Set copylocal false on all references.
            foreach (VSLangProj.Reference r in vsproject.References)
            {
                try { r.CopyLocal = false; } catch { }
            }

            // Set post build events (rename + copy).
            string rename = @"MOVE /Y ""$(TargetPath)"" ""$(TargetDir)$(ProjectName).gha""";
            string copy = @"XCOPY /Y /F ""$(TargetDir)$(ProjectName).*"" ""$(ProjectDir)..\Libraries\Output\""";
            var postbuild = project.Properties.Item("PostBuildEvent");
            postbuild.Value = rename + " && " + copy;

            // Communicate success.
            ShowMessage(OLEMSGICON.OLEMSGICON_INFO, "Updated project:\r\n" +
                "\r\n- Set all copylocal = false" +
                "\r\n- Set-post build rename to GHA" +
                "\r\n- Set-post build copy to Libraries\\Output" +
                "\r\n- Set start action to open Rhino");
        }
    }
}
