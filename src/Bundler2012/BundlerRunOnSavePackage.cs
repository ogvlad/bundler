using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Process = System.Diagnostics.Process;

namespace Bundler2012
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the informations needed to show the this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidBundlerRunOnSavePkgString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
    public sealed class BundlerRunOnSavePackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public BundlerRunOnSavePackage()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }

        private SolutionEventsListener _listener;
        private OutputWindowWriter _outputWindow;

        private DTE _dte;
        private DocumentEvents _documentEvents;
        private ProjectItemsEvents _projectItemEvents;

        //private static readonly string[] AllowedExtensions = new string[] { ".sass", ".less", ".css", ".coffee", ".js", ".bundle" };
        private static readonly string[] LessExtensions = new string[] { ".less" };
        private static readonly string[] CssExtensions = new string[] { ".css", "css.bundle" };
        private static readonly string[] JsExtensions = new string[] { ".js", ".js.bundle" };
        private static readonly string[] OtherExtensions = new string[] { ".sass", ".coffee" };

        private IDictionary<string, BundlerProcessInfo> bundlers = new Dictionary<string, BundlerProcessInfo>();

        /////////////////////////////////////////////////////////////////////////////
        // Overriden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initilaization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Trace.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            _outputWindow = new OutputWindowWriter(this, GuidList.guidBundlerRunOnSaveOutputWindowPane, "Bundler");

            _listener = new SolutionEventsListener();
            _listener.OnAfterOpenSolution += SolutionLoaded;
        }

        protected override void Dispose(bool disposing)
        {
            if (_listener != null) _listener.Dispose();
            _listener = null;

            base.Dispose(disposing);
        }
        #endregion


        public void SolutionLoaded()
        {
            _dte = (DTE)GetService(typeof(DTE));
            if (_dte == null)
            {
                Debug.WriteLine("Unable to get the EnvDTE.DTE service.");
                return;
            }

            var events = _dte.Events as Events2;
            if (events == null)
            {
                Debug.WriteLine("Unable to get the Events2.");
                return;
            }

            _documentEvents = events.get_DocumentEvents();
            _documentEvents.DocumentSaved += BundlerSaveOnLoadPackage_DocumentSaved;

            _projectItemEvents = events.ProjectItemsEvents;
            _projectItemEvents.ItemAdded += BundlerSaveOnLoadPackage_ItemAdded;
            _projectItemEvents.ItemRenamed += BundlerSaveOnLoadPackage_ItemRenamed;

            Debug.WriteLine("Solution loaded and listener document event save listener set up.");
        }

        public void BundlerSaveOnLoadPackage_DocumentSaved(Document document)
        {
            _outputWindow.WriteLine("Document saved. Running...");
            RunBundler(document.ProjectItem);
            _outputWindow.WriteLine("Finished");
        }

        public void BundlerSaveOnLoadPackage_ItemAdded(ProjectItem projectItem)
        {
            _outputWindow.WriteLine("Document added. Running...");
            RunBundler(projectItem);
            _outputWindow.WriteLine("Finished");
        }

        public void BundlerSaveOnLoadPackage_ItemRenamed(ProjectItem projectItem, string oldFileName)
        {
            _outputWindow.WriteLine("Document renamed. Running...");
            RunBundler(projectItem);
            _outputWindow.WriteLine("Finished");
        }

        private enum FileType
        {
            Unknown = 0,
            Other = 1,
            Less = 2,
            Css = 3,
            Js = 4
        }

        //private bool IsAllowedExtension(string filename)
        private bool CheckFilenameExtension(string filename, string[] extensions)
        {
            foreach (string extension in extensions)
            {
                if (filename.EndsWith(extension))
                {
                    return true;
                }
            }

            return false;
        }

        private FileType DetermineFileType(string filename)
        {
            if (CheckFilenameExtension(filename, LessExtensions))
            {
                return FileType.Less;
            }
            if (CheckFilenameExtension(filename, CssExtensions))
            {
                return FileType.Css;
            }
            if (CheckFilenameExtension(filename, JsExtensions))
            {
                return FileType.Js;
            }
            if (CheckFilenameExtension(filename, OtherExtensions))
            {
                return FileType.Other;
            }
            return FileType.Unknown;

            //var extensionIndex = filename.LastIndexOf('.');
            //if (extensionIndex < 0) return false;

            //var extension = filename.Substring(extensionIndex);
            //return AllowedExtensions.Contains(extension);
        }

        private void RunBundler(ProjectItem projectItem)
        {
            if (projectItem == null) return;

            try
            {
                if (projectItem.ContainingProject == null) return;

                // make sure this is a valid bundle file type
                //if (!IsAllowedExtension(projectItem.Name)) return;
                FileType fileType = DetermineFileType(projectItem.Name);
                if (fileType == FileType.Unknown)
                {
                    return;
                }

                // make sure the bundler exists
                var directory = new FileInfo(projectItem.ContainingProject.FileName).Directory;
                var bunderDirectory = directory.GetDirectories("bundler").FirstOrDefault();
                if (bunderDirectory == null) return;

                // make sure the files are in the bundler folder
                var fileNames = new List<string>();
                for (short i = 0; i < projectItem.FileCount; i += 1)
                    fileNames.Add(projectItem.FileNames[i]);

                if (fileNames.Any(m => m.StartsWith(bunderDirectory.FullName))) return;

                switch (fileType)
                {
                    case FileType.Less:
                        var bundleLessCommand = bunderDirectory.GetFiles("bundler-less.cmd").FirstOrDefault();
                        if (bundleLessCommand != null)
                        {
                            RunBundler(bundleLessCommand.FullName);
                            return;
                        }
                        break;

                    case FileType.Css:
                        var bundleCssCommand = bunderDirectory.GetFiles("bundler-css.cmd").FirstOrDefault();
                        if (bundleCssCommand != null)
                        {
                            RunBundler(bundleCssCommand.FullName);
                            return;
                        }
                        break;

                    case FileType.Js:
                        var bundleJsCommand = bunderDirectory.GetFiles("bundler-js.cmd").FirstOrDefault();
                        if (bundleJsCommand != null)
                        {
                            RunBundler(bundleJsCommand.FullName);
                            return;
                        }
                        break;
                }

                var bundleCommand = bunderDirectory.GetFiles("bundler.cmd").FirstOrDefault();
                if (bundleCommand == null) return;

                RunBundler(bundleCommand.FullName);
            }
            catch (Exception e)
            {
                // project item probably doesn't have a document
                Debug.WriteLine(e.Message);
            }
        }

        private void RunBundler(string bundleCommandFullName)
        {
            Debug.WriteLine("Running bundler");
            _outputWindow.WriteLine("Running bundler \"" + bundleCommandFullName + "\"");

            BundlerProcessInfo bundlerInfo = null;

            lock (bundlers)
            {
                if (!bundlers.ContainsKey(bundleCommandFullName))
                {
                    bundlers.Add(bundleCommandFullName, new BundlerProcessInfo());
                }

                bundlerInfo = bundlers[bundleCommandFullName];
            }

            lock (bundlerInfo)
            {
                if (bundlerInfo.Running)
                {
                    bundlerInfo.Queued = true;
                    return;
                }

                var process = bundlerInfo.Process = new Process();
                process.StartInfo = new ProcessStartInfo
                {
                    WindowStyle = ProcessWindowStyle.Hidden,
                    FileName = bundleCommandFullName,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                };

                process.Exited += (sender, args) =>
                {
                    lock (bundlerInfo)
                    {
                        bundlerInfo.Running = false;
                    }

                    if (bundlerInfo.Queued)
                    {
                        bundlerInfo.Queued = false;
                        RunBundler(bundleCommandFullName);
                    }

                    bundlerInfo.Process = null;
                };

                process.OutputDataReceived += (sender, args) => { if (args.Data != null) { _outputWindow.WriteLine("OUT: " + args.Data); } };
                process.ErrorDataReceived +=
                    (sender, args) =>
                    {
                        string errorString = args.Data;
                        if (errorString != null)
                        {
                            string taskItemFilePath;
                            uint taskItemLineNumber;
                            string taskItemText;

                            errorString = ProcessErrorString(errorString, out taskItemFilePath, out taskItemLineNumber, out taskItemText);
                            if (!string.IsNullOrEmpty(taskItemFilePath) && File.Exists(taskItemFilePath))
                            {
                                _outputWindow.AddTaskItem("ERR: " + errorString, taskItemText, taskItemFilePath, taskItemLineNumber);
                            }
                            else
                            {
                                _outputWindow.WriteLine("{0}", "ERR: " + errorString);
                            }
                        } 
                    };

                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
            }
        }

        private string ProcessErrorString(string errorString, out string taskItemFilePath, out uint taskItemLineNumber, out string taskItemText)
        {
            taskItemFilePath = null;
            taskItemLineNumber = 0;
            taskItemText = null;

            if (errorString == null)
            {
                return errorString;
            }

            errorString = new System.Text.RegularExpressions.Regex(System.Text.Encoding.ASCII.GetString(new byte[] { 27 }) + @"\[(\d)+m").Replace(errorString, string.Empty);

            try
            {

                const string errorStringPart0 = "ParseError: ";
                const string errorStringPart2 = " in ";
                const string errorStringPart4 = " on line ";
                const string errorStringPart6 = ", column ";
                int indexOfErrorStringPart1 = (errorString.StartsWith(errorStringPart0) ? errorStringPart0.Length : -1);
                int indexOfErrorStringPart2 = (indexOfErrorStringPart1 < 0 ? -1 : errorString.IndexOf(errorStringPart2, indexOfErrorStringPart1));
                int indexOfErrorStringPart3 = (indexOfErrorStringPart2 < 0 ? -1 : indexOfErrorStringPart2 + errorStringPart2.Length);
                int indexOfErrorStringPart4 = (indexOfErrorStringPart3 < 0 ? -1 : errorString.IndexOf(errorStringPart4, indexOfErrorStringPart3));
                int indexOfErrorStringPart5 = (indexOfErrorStringPart4 < 0 ? -1 : indexOfErrorStringPart4 + errorStringPart4.Length);
                int indexOfErrorStringPart6 = (indexOfErrorStringPart5 < 0 ? -1 : errorString.IndexOf(errorStringPart6, indexOfErrorStringPart5));

                if (!(indexOfErrorStringPart6 < 0))
                {
                    taskItemText = errorString.Substring(indexOfErrorStringPart1, indexOfErrorStringPart2 - indexOfErrorStringPart1);
                    taskItemFilePath = errorString.Substring(indexOfErrorStringPart3, indexOfErrorStringPart4 - indexOfErrorStringPart3);
                    taskItemLineNumber = uint.Parse(errorString.Substring(indexOfErrorStringPart5, indexOfErrorStringPart6 - indexOfErrorStringPart5));
                }
            }
            catch (Exception exception)
            {
                _outputWindow.WriteLine("{0}\n{1}", exception.ToString(), exception.StackTrace);
            }

            return errorString;
        }

        private class BundlerProcessInfo
        {
            public bool Running { get; set; }
            public bool Queued { get; set; }
            public Process Process { get; set; }
        }
    }
}
