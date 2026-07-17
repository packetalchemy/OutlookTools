using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using OutlookTools.Settings;

namespace OutlookTools
{
    /// <summary>
    /// OutlookTools v1.2.0 — Main add-in entry point.
    /// Without VSTO, this acts as a helper library.
    /// With VSTO, ThisAddIn is the Outlook entry point.
    /// </summary>
    public partial class ThisAddIn
    {
        private Application _application;

        public Application Application => _application;

        public void SetApplication(Application app)
        {
            _application = app;
            LogDebug("OutlookTools v1.2.0 started.");
        }

        public static void LogDebug(string message)
        {
            try
            {
                if (!SettingsManager.GetDebugLogEnabled()) return;
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookTools");
                Directory.CreateDirectory(dir);
                File.AppendAllText(Path.Combine(dir, "outlook-tools.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {message}{Environment.NewLine}");
            }
            catch { }
        }

        // VSTO handlers removed — standalone mode
    }

    /// <summary>
    /// Static accessor for the Application object.
    /// Used by all features to get the Outlook Application instance.
    /// </summary>
    public static class Globals
    {
        private static ThisAddIn _addIn;

        public static ThisAddIn ThisAddIn
        {
            get
            {
                if (_addIn == null)
                    _addIn = new ThisAddIn();
                return _addIn;
            }
        }
    }
}
