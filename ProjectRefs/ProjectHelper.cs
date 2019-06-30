using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace ProjectRefs
{
    public class ProjectHelper
    {
        private readonly PackageHelper _packageHelper;

        public ProjectHelper(PackageHelper packageHelper)
        {
            _packageHelper = packageHelper;
        }

        public IEnumerable<string> GetProjectReferencePaths(string projectPath, IVsHierarchy project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            using (var fileStream = File.OpenRead(projectPath))
            {
                using (var xmlReader = XmlReader.Create(fileStream))
                {
                    if (xmlReader.ReadToDescendant("ProjectReference"))
                    {
                        do
                        {
                            if (xmlReader.MoveToAttribute("Include"))
                            {
                                var includeValue = xmlReader.ReadContentAsString();

                                if (!string.IsNullOrEmpty(includeValue))
                                {
                                    _packageHelper.LogDebug(Environment.NewLine + projectPath + "--->" + includeValue);

                                    var referencedProjectPath = ResolveMacrosInPath(project, includeValue);

                                    if (!Path.IsPathRooted(referencedProjectPath))
                                    {
                                        referencedProjectPath = Path.Combine(Path.GetDirectoryName(projectPath), referencedProjectPath);
                                    }

                                    // Save the canonical path of the project.
                                    yield return Path.GetFullPath(referencedProjectPath);
                                }
                            }
                        }
                        while (xmlReader.ReadToNextSibling("ProjectReference"));
                    }
                }
            }
        }

        private string ResolveMacrosInPath(IVsHierarchy projectHierarchy, string path)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if ((projectHierarchy == null) || string.IsNullOrEmpty(path))
            {
                return path;
            }

            int startIndex = 0;
            var i = path.IndexOf("$(", startIndex);

            if (i < 0)
            {
                // Path has no macros
                return path;
            }

            string resolvedPath = string.Empty;

            while (i >= 0)
            {
                var j = path.IndexOf(")", i);
                resolvedPath += path.Substring(startIndex, i - startIndex);

                var propertyName = path.Substring(i + 2, j - (i + 2));
                var propertyValue = GetProjectPropertyValue(projectHierarchy, propertyName);

                _packageHelper.LogDebug(propertyName + " = " + propertyValue);

                resolvedPath += propertyValue;

                startIndex = j + 1;
                i = path.IndexOf("$(", startIndex);
            }

            if (startIndex < path.Length)
            {
                resolvedPath += path.Substring(startIndex, path.Length - startIndex);
            }

            return resolvedPath;
        }

        /// <summary>
        /// Gets the value of a property defined in the passed-in project
        /// </summary>
        /// <param name="projectHierarchy">Project</param>
        /// <param name="propertyName">Name of the property</param>
        /// <returns>Value of the property. If there is an error, returns <code>string.Empty</code></returns>
        private string GetProjectPropertyValue(IVsHierarchy projectHierarchy, string propertyName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (projectHierarchy is IVsBuildPropertyStorage vsBuildPropertyStorage)
            {
                if (ErrorHandler.Succeeded(vsBuildPropertyStorage.GetPropertyValue(propertyName, null, (uint)_PersistStorageType.PST_PROJECT_FILE, out string propertyValue)))
                {
                    return propertyValue;
                }
            }

            _packageHelper.LogInfo("ERROR reading project property - " + propertyName ?? "<null>");
            return string.Empty;
        }
    }
}
