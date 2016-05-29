namespace CommandLineRepl
{
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using System;

    internal class OutputPane
    {
        private IServiceProvider ServiceProvider;
        private Guid guidPane;

        public OutputPane(IServiceProvider ServiceProvider, Guid guidPane)
        {
            this.ServiceProvider = ServiceProvider;
            this.guidPane = guidPane;
        }

        public void Clear()
        {
            GetOutputWindowPane(guidPane).Clear();
        }

        public void OutputString(string text)
        {
            IVsOutputWindowPane outputWindowPane = GetOutputWindowPane(guidPane);

            if (outputWindowPane != null)
            {
                try
                {
                    outputWindowPane.Activate();
                    outputWindowPane.OutputString(Environment.NewLine + text);
                }
                catch (Exception ex)
                {
                    ShowMessageBox("OutputString()", ex.Message);
                }
            }
        }

        private IVsOutputWindowPane GetOutputWindowPane(Guid guidPane)
        {
            try
            {
                const int VISIBLE = 1;
                const int DO_NOT_CLEAR_WITH_SOLUTION = 0;

                IVsOutputWindow outputWindow;
                IVsOutputWindowPane outputWindowPane = null;
                int hr;

                // Get the output window
                outputWindow = ServiceProvider.GetService(typeof(SVsOutputWindow)) as IVsOutputWindow;

                // The General pane is not created by default. We must force its creation
                if (guidPane == Microsoft.VisualStudio.VSConstants.OutputWindowPaneGuid.GeneralPane_guid)
                {
                    hr = outputWindow.CreatePane(guidPane, "General", VISIBLE, DO_NOT_CLEAR_WITH_SOLUTION);
                    Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(hr);
                }

                // Get the pane
                hr = outputWindow.GetPane(guidPane, out outputWindowPane);
                Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(hr);
                return outputWindowPane;
            }
            catch (Exception ex)
            {
                ShowMessageBox("GetOutputWindowPane()", ex.Message);
            }

            return null;
        }

        private void ShowMessageBox(string title, string message)
        {
            VsShellUtilities.ShowMessageBox(this.ServiceProvider, message, title, OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}