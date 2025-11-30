using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.Win32;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace Devvcat.SSMS
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "3.1.0", IconResourceID = 400)] // Info on this package for Help/About
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

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            Initialize();

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
                    $"Packages\\{{{PackageGuidString}}}");

                if (registryKey != null)
                {
                    registryKey.SetValue("SkipLoading", 1, RegistryValueKind.DWord);
                    registryKey.Close();
                }
            }
            catch
            { }
        }
    }
}
