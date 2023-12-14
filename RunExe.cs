using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

//https://learn.microsoft.com/en-us/visualstudio/extensibility/dynamically-adding-menu-items?view=vs-2022

namespace SSMSExec
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class RunExe
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("c7f81868-05c2-4a2d-877a-0f7decbf136b");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="RunExe"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private RunExe(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new OleMenuCommand((s, e) => package.JoinableTaskFactory.RunAsync(() => ExecuteAsync(s, e)), menuCommandID);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static RunExe Instance
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
            // Switch to the main thread - the call to AddCommand in RunExe's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new RunExe(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async Task ExecuteAsync(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            var generalOptions = (GeneralOptions)package.GetDialogPage(typeof(GeneralOptions));

            if (!generalOptions.ExeActive)
            {
                return;
            }

            EnvDTE.DTE dte = await this.ServiceProvider.GetServiceAsync(typeof(Microsoft.VisualStudio.Shell.Interop.SDTE)) as EnvDTE.DTE;

            //get currently active window
            Window activeWindow = ((DTE2)dte).ActiveWindow;

            if (activeWindow == null || activeWindow.Document == null)
            {
                MessageBox.Show("No active window. Please open a query window.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //get text from window
            TextDocument textDoc = (TextDocument)activeWindow.Document.Object("TextDocument");
            if (textDoc == null)
            {
                MessageBox.Show("No text document. Please open a query window.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            EditPoint editPoint = textDoc.StartPoint.CreateEditPoint();
            string sqlQuery = editPoint.GetText(textDoc.EndPoint);

            string arguments = $"{generalOptions.ExeParameter1} {generalOptions.ExeParameter2} {generalOptions.ExeParameter3} {generalOptions.ExeParameter4} {generalOptions.ExeParameter5} {generalOptions.ExeParameter6}";

            ProcessStartInfo startInfo = new ProcessStartInfo() {
                FileName = generalOptions.ExeLocation,
                Arguments = arguments,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };



            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo = startInfo;
            process.Start();

            await process.StandardInput.WriteLineAsync(sqlQuery);
            process.StandardInput.Close();

            try { 
                var updatedQuery = await process.StandardOutput.ReadToEndAsync();
                var processError = await process.StandardError.ReadToEndAsync();
                if (process.ExitCode != 0)
                {
                    MessageBox.Show("Error: Process exited unexpectedly", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                await process.WaitForExitAsync();

                if (!string.IsNullOrEmpty(processError))
                {
                    // Handle the error
                    MessageBox.Show("Error: " + processError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    editPoint.ReplaceText(textDoc.EndPoint, updatedQuery, (int)vsEPReplaceTextOptions.vsEPReplaceTextAutoformat);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            var menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                var generalOptions = (GeneralOptions)package.GetDialogPage(typeof(GeneralOptions));
                menuCommand.Visible = generalOptions.ExeActive;
                if (generalOptions.ExeActive)
                {
                    menuCommand.Text = generalOptions.ButtonText;
                }
            }
        }
    }
}
