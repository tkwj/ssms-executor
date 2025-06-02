using Devvcat.SSMS.Options;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Devvcat.SSMS
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "3.0.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string, PackageAutoLoadFlags.BackgroundLoad)]
#if DEBUG
    [ProvideOptionPage(typeof(GeneralOptionsPage), "SSMS Executor", "General", 100, 101, true)]
#endif
    [Guid(PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class ExecutorPackage : AsyncPackage
    {
        private const string PackageGuidString = "a64d9865-b938-4543-bf8f-a553cc4f67f3";

        public ExecutorPackage()
        {
        }

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            base.Initialize();

            ExecutorCommand.Initialize(this);
        }

        protected override int QueryClose(out bool canClose)
        {
            SetSkipLoading();

            return base.QueryClose(out canClose);
        }

        void SetSkipLoading()
        {
            try
            {
                var registryKey = UserRegistryRoot.CreateSubKey(
                    string.Format("Packages\\{{{0}}}", PackageGuidString));

                registryKey.SetValue("SkipLoading", 1, RegistryValueKind.DWord);
                registryKey.Close();
            }
            catch
            { }
        }
    }
}
