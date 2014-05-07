using System;
using Microsoft.VisualStudio.Shell.Interop;

namespace Bundler2012
{
    internal class OutputWindowWriter
    {
        private IVsOutputWindowPane _outputWindowPane;

        
        public OutputWindowWriter(IServiceProvider serviceProvider, string outWindowGuid, string outWindowName)
        {
            var outputWindow = serviceProvider.GetService(typeof(SVsOutputWindow)) as IVsOutputWindow;
            if (outputWindow == null) throw new Exception("Unable to create an output pane.");

            var paneGuid = new Guid(outWindowGuid);
            outputWindow.GetPane(ref paneGuid, out _outputWindowPane);
            if (_outputWindowPane == null)
            {
                outputWindow.CreatePane(ref paneGuid, outWindowName, 1, 0);
                outputWindow.GetPane(ref paneGuid, out _outputWindowPane);
            }
        }

        public void Write(string format, params object[] parameters)
        {
            if (_outputWindowPane == null || format == null) return;

            _outputWindowPane.OutputString(String.Format(format, parameters));
        }

        public void WriteLine(string format, params object[] parameters)
        {
            Write(format + Environment.NewLine, parameters);
        }

        public void AddTaskItem(string taskItemOutputString, string taskItemText, string filePath, uint lineNumber)
        {
            if (_outputWindowPane == null) return;

            _outputWindowPane.OutputTaskItemString(taskItemOutputString + Environment.NewLine, VSTASKPRIORITY.TP_HIGH, VSTASKCATEGORY.CAT_CODESENSE, "SUBCATEGORY", (int)(Microsoft.VisualStudio.Shell.Interop._vstaskbitmap.BMP_COMPILE), filePath, (lineNumber - 1), taskItemText);
        }

        public void Clear()
        {
            _outputWindowPane.Clear();
        }
    }
}
