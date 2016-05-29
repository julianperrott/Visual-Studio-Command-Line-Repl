namespace CommandLineRepl
{
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using System;
    using System.ComponentModel.Design;

    internal sealed class CommandLineRepl
    {
        private ConsoleRedirect redirect;

        public const int CommandId = 0x0100;

        public static readonly Guid CommandSet = new Guid("c22e4f5b-e1ba-4383-a8fa-96d99ea193fc");

        private readonly CommandLineReplPackage package;

        private readonly OutputPane outputPane;
        private readonly OutputPane lastOutputPane;
        private readonly EditorText editorText;

        private System.Windows.Forms.Form popup;

        private CommandLineRepl(CommandLineReplPackage package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            outputPane = new OutputPane(package, Microsoft.VisualStudio.VSConstants.OutputWindowPaneGuid.GeneralPane_guid);
            lastOutputPane = new OutputPane(package, Microsoft.VisualStudio.VSConstants.OutputWindowPaneGuid.DebugPane_guid);
            this.editorText = new EditorText(package, package);

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);

                redirect = new ConsoleRedirect("cmd", "/k echo type some commands e.g. (dir, cd, help, etc... :)", this.ServiceProvider);
                redirect.ReceivedOutput += new ReceivedOutputEventHandler(outputPane.OutputString);
                redirect.ReceivedOutput += new ReceivedOutputEventHandler(lastOutputPane.OutputString);
                redirect.Start();
            }
        }

        public static CommandLineRepl Instance
        {
            get;
            private set;
        }

        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        public static void Initialize(CommandLineReplPackage package)
        {
            Instance = new CommandLineRepl(package);
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            try
            {
                string command = editorText.GetText();
                if (!string.IsNullOrEmpty(command))
                {
                    lastOutputPane.Clear();
                    lastOutputPane.OutputString(command);
                    outputPane.OutputString(command);
                    redirect.SendInput(command + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                ShowMessageBox("Oh dear in " + this.GetType().FullName, ex.ToString());
            }
        }

        private void ShowMessageBox(string title, string message)
        {
            VsShellUtilities.ShowMessageBox(this.ServiceProvider, message, title, OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}