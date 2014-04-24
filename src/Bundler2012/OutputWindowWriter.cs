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

        public void Clear()
        {
            _outputWindowPane.Clear();
        }
    }
}
