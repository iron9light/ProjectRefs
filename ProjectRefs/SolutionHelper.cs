using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

using EnvDTE;

using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace ProjectRefs
{
    public class SolutionHelper
    {
        private readonly IVsSolution _solution;

        private readonly PackageHelper _packageHelper;

        private readonly Lazy<IVsHierarchy> _refsFolderHierarchy;

        public SolutionHelper(IVsSolution solution, PackageHelper packageHelper)
        {
            Requires.NotNull(solution, nameof(solution));
            Requires.NotNull(packageHelper, nameof(packageHelper));
            _solution = solution;
            _packageHelper = packageHelper;
            _refsFolderHierarchy = new Lazy<IVsHierarchy>(GetOrCreateRefsFolder);
        }

        public IVsHierarchy RefsFolderHierarchy => _refsFolderHierarchy.Value;

        public void RemoveProject(string projectPath, IVsHierarchy project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _packageHelper.LogInfo($"Removing reference project: {projectPath}");
            _solution.CloseSolutionElement((uint)__VSSLNCLOSEOPTIONS.SLNCLOSEOPT_DeleteProject, project, 0);
            _packageHelper.LogInfo($"Removed reference project: {projectPath}");
        }

        public IVsHierarchy AddRefProject(string projectPath)
        {
            Requires.NotNullOrEmpty(projectPath, nameof(projectPath));

            ThreadHelper.ThrowIfNotOnUIThread();

            ErrorHandler.ThrowOnFailure(_solution.GetSolutionInfo(out var solutionDir, out _, out _));

            var relativePath = GetRelativePath(projectPath, solutionDir);

            if (ErrorHandler.Succeeded(_solution.GetProjectOfUniqueName(relativePath, out var project)))
            {
                return project;
            }
            else
            {
                _packageHelper.LogInfo($"Adding reference project: {projectPath}");
                ErrorHandler.ThrowOnFailure(((IVsSolution6)_solution).AddExistingProject(projectPath, RefsFolderHierarchy, out project));
                _packageHelper.LogInfo($"Added reference project: {projectPath}");
                return project;
            }
        }

        public IEnumerable<(string Path, IVsHierarchy Project, bool IsRef)> GetAllProjects()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ErrorHandler.ThrowOnFailure(_solution.GetSolutionInfo(out var solutionDir, out _, out _));
            ErrorHandler.ThrowOnFailure(_solution.GetProjectEnum((uint)__VSENUMPROJFLAGS.EPF_ALLINSOLUTION, Guid.Empty, out var hierarchies));

            var hierarchy = new IVsHierarchy[1];

            while (ErrorHandler.Succeeded(hierarchies.Next(1, hierarchy, out var fetchedCount)) && fetchedCount == 1)
            {
                if (hierarchy[0] is IPersist persist)
                {
                    ErrorHandler.ThrowOnFailure(persist.GetClassID(out var classId));
                    if (classId == VSConstants.CLSID.MiscellaneousFilesProject_guid || classId == VSConstants.CLSID.SolutionFolderProject_guid)
                    {
                        continue;
                    }

                    if (ErrorHandler.Succeeded(_solution.GetUniqueNameOfProject(hierarchy[0], out var projectUniqueName)))
                    {
                        var projectPath = Path.Combine(solutionDir, projectUniqueName);
                        projectPath = Path.GetFullPath(projectPath);
                        yield return (projectPath, hierarchy[0], IsRefProject(hierarchy[0]));
                    }
                }
            }
        }

        private Project Hierarchy2Project(IVsHierarchy projectHierarchy)
        {
            Requires.NotNull(projectHierarchy, nameof(projectHierarchy));

            ThreadHelper.ThrowIfNotOnUIThread();
            ErrorHandler.ThrowOnFailure(projectHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_ExtObject, out var projectObject));
            return projectObject as Project;
        }

        private bool IsRefsFolder(Project project)
        {
            Requires.NotNull(project, nameof(project));

            ThreadHelper.ThrowIfNotOnUIThread();
            return project.Name == "refs" && project.ParentProjectItem == null;
        }

        ////public bool IsRefProject(IVsHierarchy projectHierarchy)
        ////{
        ////    ThreadHelper.ThrowIfNotOnUIThread();
        ////    var project = Hierarchy2Project(projectHierarchy);

        ////    var parent = project?.ParentProjectItem?.Collection?.Parent;
        ////    if (parent != null && parent is Project parentProject)
        ////    {
        ////        if (IsRefsFolder(parentProject))
        ////        {
        ////            return true;
        ////        }
        ////    }

        ////    return false;
        ////}

        public bool IsRefProject(IVsHierarchy projectHierarchy)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ErrorHandler.ThrowOnFailure(projectHierarchy.GetProperty((uint)VSConstants.VSITEMID.Root, (int)__VSHPROPID.VSHPROPID_ParentHierarchy, out var parentObject));
            var parent = parentObject as IVsHierarchy;
            Assumes.Present(parent);
            var parentProject = Hierarchy2Project(parent);
            return parentProject != null && IsRefsFolder(parentProject);
        }

        private IVsHierarchy GetOrCreateRefsFolder()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return GetRefsFolder() ?? CreateRefsFolder();
        }

        private IVsHierarchy GetRefsFolder()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ErrorHandler.ThrowOnFailure(_solution.GetProjectEnum((uint)(__VSENUMPROJFLAGS.EPF_ALLINSOLUTION | __VSENUMPROJFLAGS.EPF_MATCHTYPE), VSConstants.CLSID.SolutionFolderProject_guid, out var hierarchies));

            var hierarchy = new IVsHierarchy[1];

            while (ErrorHandler.Succeeded(hierarchies.Next(1, hierarchy, out var fetchedCount)) && fetchedCount == 1)
            {
                var project = Hierarchy2Project(hierarchy[0]);
                if (project != null && IsRefsFolder(project))
                {
                    return hierarchy[0];
                }
            }

            return null;
        }

        private IVsHierarchy CreateRefsFolder()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var iidProject = typeof(IVsHierarchy).GUID;
            ErrorHandler.ThrowOnFailure(_solution.CreateProject(VSConstants.CLSID.SolutionFolderProject_guid, null, null, "refs", 0, ref iidProject, out var ptr));
            if (ptr == IntPtr.Zero)
            {
                _packageHelper.LogInfo("Failed to create \"refs\" solution folder. The returned pointer is NULL");
                return null;
            }

            return (IVsHierarchy)Marshal.GetObjectForIUnknown(ptr);
        }

        private static string GetRelativePath(string path, string relativeTo)
        {
            path = Path.GetFullPath(path);
            relativeTo = Path.GetFullPath(relativeTo);
            Uri pathUri = new Uri(path);
            // Folders must end in a slash
            if (!relativeTo.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                relativeTo += Path.DirectorySeparatorChar;
            }
            Uri folderUri = new Uri(relativeTo);
            return Uri.UnescapeDataString(folderUri.MakeRelativeUri(pathUri).ToString().Replace('/', Path.DirectorySeparatorChar));
        }
    }
}
