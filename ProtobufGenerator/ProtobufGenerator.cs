using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using Process = System.Diagnostics.Process;
using Task = System.Threading.Tasks.Task;

namespace ProtobufGenerator
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ProtobufGenerator
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("4d55cc65-dbd0-4959-8c85-ebc1c0df8606");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        private readonly string protocPath;

        /// <summary>
        /// Initializes a new instance of the <see cref="ProtobufGenerator"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ProtobufGenerator(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            menuItem.Supported = false;
            commandService.AddCommand(menuItem);

            this.protocPath = Path.GetDirectoryName(this.GetType().Assembly.Location) + "/protoc/protoc.exe";
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ProtobufGenerator Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
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
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Verify the current thread is the UI thread - the call to AddCommand in ProtobufGenerator's constructor requires
            // the UI thread.
            ThreadHelper.ThrowIfNotOnUIThread();

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new ProtobufGenerator(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            DTE dte = (DTE)await this.ServiceProvider.GetServiceAsync(typeof(DTE));
            foreach (SelectedItem item in dte.SelectedItems)
            {
                ProjectItem pi = item.ProjectItem;
                string type = null;
                switch (pi.ContainingProject.CodeModel.Language)
                {
                    case CodeModelLanguageConstants.vsCMLanguageCSharp:
                        type = "csharp_out";
                        break;
                    case CodeModelLanguageConstants.vsCMLanguageVC:
                        type = "cpp_out";
                        break;
                    default:
                        break;
                }
                if (type != null)
                {
                    string srcFile = pi.FileNames[0];
                    string tmpFile = Path.GetTempPath() + pi.GetHashCode() + ".proto";
                    File.Copy(srcFile, tmpFile, true);
                    this.ClearItems(pi);
                    string path = Path.GetTempPath() + pi.GetHashCode();
                    if (Directory.Exists(path))
                        Directory.Delete(path, true);
                    Directory.CreateDirectory(path);
                    ProcessStartInfo psi = new ProcessStartInfo(this.protocPath, $"\"{pi.Name}\" --{type}=\"{path}\"");
                    psi.UseShellExecute = false;
                    psi.CreateNoWindow = true;
                    psi.WorkingDirectory = Path.GetDirectoryName(srcFile);
                    using (Process p = Process.Start(psi))
                    {
                        if (p?.WaitForExit(10 * 1000) == true)
                        {
                            string[] dstFiles = Directory.GetFiles(path);
                            if (dstFiles.Length != 0)
                            {
                                string srcDir = Path.GetDirectoryName(srcFile);
                                string srcName = Path.GetFileNameWithoutExtension(srcFile);
                                int count = 0;
                                switch (type)
                                {
                                    case "csharp_out":
                                        count = 1;
                                        break;
                                    case "cpp_out":
                                        srcDir = Path.Combine(srcDir, "generate");
                                        srcName += ".pb";
                                        count = 2;
                                        break;
                                }
                                if (dstFiles.Length == count)
                                {
                                    if (Directory.Exists(srcDir) == false)
                                        Directory.CreateDirectory(srcDir);
                                    for (int i = 0; i < count; i++)
                                    {
                                        string dstExt = Path.GetExtension(dstFiles[i]);
                                        string dstFile = Path.Combine(srcDir, srcName + dstExt);
                                        File.Copy(dstFiles[i], dstFile, true);
                                        pi.ProjectItems.AddFromFile(dstFile);
                                    }
                                }
                            }
                        }
                    }
                    Directory.Delete(path, true);
                    File.Delete(tmpFile);
                }
            }
        }

        private void ClearItems(ProjectItem item)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            IList<ProjectItem> list = new List<ProjectItem>();
            foreach (ProjectItem pi in item.ProjectItems)
                list.Add(pi);
            foreach (ProjectItem pi in list)
                pi.Remove();
        }
    }
}
