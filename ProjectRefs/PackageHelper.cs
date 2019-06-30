using System;

using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace ProjectRefs
{
    public class PackageHelper
    {
        private readonly AsyncPackage _package;

        private readonly Lazy<IVsOutputWindowPane> _outputWindowPane;

        public PackageHelper(AsyncPackage package)
        {
            Requires.NotNull(package, nameof(package));

            _package = package;

            _outputWindowPane = new Lazy<IVsOutputWindowPane>(() => 
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                var outputWindowPane = _package.GetOutputPane(VSConstants.OutputWindowPaneGuid.GeneralPane_guid, "ProjectRefs");
                ErrorHandler.ThrowOnFailure(outputWindowPane.Activate());
                return outputWindowPane;
            });
        }

        private IVsOutputWindowPane OutputWindowPane => _outputWindowPane.Value;

        public void LogInfo(string s)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            OutputWindowPane.OutputStringThreadSafe(s + Environment.NewLine);
        }

        public void LogDebug(string s)
        {
#if DEBUG
            ThreadHelper.ThrowIfNotOnUIThread();
            OutputWindowPane.OutputStringThreadSafe(s + Environment.NewLine);
#endif
        }
    }
}
