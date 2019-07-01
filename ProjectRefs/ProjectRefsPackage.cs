using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Events;
using Microsoft.VisualStudio.Shell.Interop;

using Task = System.Threading.Tasks.Task;

namespace ProjectRefs
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(ProjectRefsPackage.PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionOpening_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class ProjectRefsPackage : AsyncPackage
    {
        /// <summary>
        /// ProjectRefsPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "05956a94-db53-4a82-b0f2-6e85575382bf";

        private bool _solutionOpened = false;

        private PackageHelper _packageHelper;

        private SolutionHelper _solutionHelper;

        private ProjectHelper _projectHelper;

        private class VsRunningDocTableEvents : IVsRunningDocTableEvents, IVsRunningDocTableEvents4
        {
            private readonly ProjectRefsPackage _package;

            public VsRunningDocTableEvents(ProjectRefsPackage package)
            {
                _package = package;
            }

            public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
            {
                return VSConstants.S_OK;
            }

            public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
            {
                return VSConstants.S_OK;
            }

            public int OnAfterSave(uint docCookie)
            {
                return VSConstants.S_OK;
            }

            public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
            {
                return VSConstants.S_OK;
            }

            public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
            {
                return VSConstants.S_OK;
            }

            public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame)
            {
                return VSConstants.S_OK;
            }

            public int OnBeforeFirstDocumentLock(IVsHierarchy pHier, uint itemid, string pszMkDocument)
            {
                return VSConstants.S_OK;
            }

            public int OnAfterSaveAll()
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                _package._packageHelper.LogDebug("OnAfterSaveAll");
                _package.ProcessProjectRefsAsync();
                return VSConstants.S_OK;
            }

            public int OnAfterLastDocumentUnlock(IVsHierarchy pHier, uint itemid, string pszMkDocument, int fClosedWithoutSaving)
            {
                return VSConstants.S_OK;
            }
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            _packageHelper = new PackageHelper(this);

            var solution = await GetServiceAsync(typeof(SVsSolution)) as IVsSolution;
            Assumes.Present(solution);

            var runningDocumentTable = await GetServiceAsync(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;
            Assumes.Present(runningDocumentTable);

            runningDocumentTable.AdviseRunningDocTableEvents(new VsRunningDocTableEvents(this), out _);

            _solutionHelper = new SolutionHelper(solution, _packageHelper);
            _projectHelper = new ProjectHelper(_packageHelper);

            SolutionEvents.OnAfterOpenSolution += SolutionEvents_OnAfterOpenSolution;
            ////SolutionEvents.OnAfterOpenProject += SolutionEvents_OnAfterOpenProject;
            ////SolutionEvents.OnAfterLoadProject += SolutionEvents_OnAfterLoadProject;

            ErrorHandler.ThrowOnFailure(solution.GetProperty((int)__VSPROPID.VSPROPID_IsSolutionOpen, out object value));
            if (value is bool isOpened && isOpened)
            {
                SolutionEvents_OnAfterOpenSolution(null, null);
            }
        }

        ////private void SolutionEvents_OnAfterLoadProject(object sender, LoadProjectEventArgs e)
        ////{
        ////    ThreadHelper.ThrowIfNotOnUIThread();
        ////    _packageHelper.LogDebug("OnAfterLoadProject");
        ////    ProcessProjectRefs(e.RealHierarchy);
        ////}

        ////private void SolutionEvents_OnAfterOpenProject(object sender, OpenProjectEventArgs e)
        ////{
        ////    ThreadHelper.ThrowIfNotOnUIThread();
        ////    _packageHelper.LogDebug("OnAfterOpenProject");
        ////    ProcessProjectRefs(e.Hierarchy);
        ////}

        private void SolutionEvents_OnAfterOpenSolution(object sender, OpenSolutionEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _packageHelper.LogDebug("OnAfterOpenSolution");
            _solutionOpened = true;
            ProcessProjectRefs();
        }

        private async Task ProcessProjectRefsAsync(IVsHierarchy project = null)
        {
            await Task.Delay(TimeSpan.FromSeconds(1));
            await JoinableTaskFactory.SwitchToMainThreadAsync();
            ProcessProjectRefs(project);
        }

        private void ProcessProjectRefs(IVsHierarchy project = null)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (!_solutionOpened)
            {
                return;
            }

            if (project != null && _solutionHelper.IsRefProject(project))
            {
                return;
            }

            var projects = _solutionHelper.GetAllProjects().ToList();

            var addedPaths = new HashSet<string>();

            foreach (var item in projects.Where(item => !item.IsRef))
            {
                AddRefProjects(item.Path, item.Project, addedPaths);
            }

            foreach (var addedPath in addedPaths)
            {
                _packageHelper.LogDebug($"added path: {addedPath}");
            }

            foreach (var item in projects.Where(item => item.IsRef))
            {
                _packageHelper.LogDebug($"ref path: {item.Path}");
                if (!addedPaths.Contains(item.Path))
                {
                    _solutionHelper.RemoveProject(item.Path, item.Project);
                }
            }
        }

        private void AddRefProjects(string projectPath, IVsHierarchy project, ISet<string> addedPaths)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            foreach (var referencedProjectPath in _projectHelper.GetProjectReferencePaths(projectPath, project))
            {
                addedPaths.Add(referencedProjectPath);
                _packageHelper.LogDebug($"Try add: {referencedProjectPath}");
                var addedProject = _solutionHelper.AddRefProject(referencedProjectPath);
                if (addedProject != null)
                {
                    AddRefProjects(referencedProjectPath, addedProject, addedPaths);
                }
                else
                {
                    _packageHelper.LogInfo($"Got null: {referencedProjectPath}");
                }
            }
        }
    }
}
