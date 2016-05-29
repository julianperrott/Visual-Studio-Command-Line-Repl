namespace CommandLineRepl
{
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using System;
    using System.Diagnostics.CodeAnalysis;
    using System.Runtime.InteropServices;

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(CommandLineReplPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class CommandLineReplPackage : Package, IDteProvider
    {
        public const string PackageGuidString = "a8aa2fc6-518c-48f0-9bcb-49026a8d8476";

        protected override void Initialize()
        {
            CommandLineRepl.Initialize(this);
            base.Initialize();
            this.InitializeDTE();
        }

        private DteInitializer dteInitializer;

        public EnvDTE80.DTE2 Dte
        {
            get; set;
        }

        private void InitializeDTE()
        {
            IVsShell shellService;

            this.Dte = this.GetService(typeof(Microsoft.VisualStudio.Shell.Interop.SDTE)) as EnvDTE80.DTE2;

            if (this.Dte == null) // The IDE is not yet fully initialized
            {
                shellService = this.GetService(typeof(SVsShell)) as IVsShell;
                this.dteInitializer = new DteInitializer(shellService, this.InitializeDTE);
            }
            else
            {
                this.dteInitializer = null;
            }
        }
    }
}