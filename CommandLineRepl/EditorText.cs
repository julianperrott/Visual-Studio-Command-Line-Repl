namespace CommandLineRepl
{
    using EnvDTE;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using System;

    public class EditorText
    {
        private readonly IServiceProvider ServiceProvider;
        private readonly IDteProvider dteProvider;

        public EditorText(IServiceProvider ServiceProvider, IDteProvider dteProvider)
        {
            this.ServiceProvider = ServiceProvider;
            this.dteProvider = dteProvider;
        }

        public string GetText()
        {
            TextSelection objTextSelection = GetTextSelection();

            if (objTextSelection == null)
            {
                return string.Empty;
            }

            if (!objTextSelection.IsEmpty)
            {
                return objTextSelection.Text;
            }

            return GetLineText(objTextSelection);
        }

        private TextSelection GetTextSelection()
        {
            try
            {
                EnvDTE.Document objDocument = this.dteProvider.Dte.ActiveDocument;
                if (objDocument == null)
                {
                    ShowMessageBox("GetTextSelection()", "ActiveDocument not found. Are you in a code editor window ?");
                    return null;
                }

                EnvDTE.TextDocument objTextDocument = (EnvDTE.TextDocument)objDocument.Object("TextDocument");
                EnvDTE.TextSelection objTextSelection = objTextDocument.Selection;
                return objTextSelection;
            }
            catch (Exception ex)
            {
                ShowMessageBox("GetTextSelection()", ex.Message);
            }
            return null;
        }

        private string GetLineText(TextSelection objTextSelection)
        {
            try
            {
                VirtualPoint objActive = objTextSelection.ActivePoint;
                objTextSelection.StartOfLine((EnvDTE.vsStartOfLineOptions)(0), true);
                string text = objTextSelection.Text;
                objTextSelection.EndOfLine(true);
                var result = text + objTextSelection.Text;
                objTextSelection.EndOfLine(false);
                return result;
            }
            catch (Exception ex)
            {
                ShowMessageBox("GetLineText()", ex.Message);
            }

            return string.Empty;
        }

        private void ShowMessageBox(string title, string message)
        {
            VsShellUtilities.ShowMessageBox(this.ServiceProvider, message, title, OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}