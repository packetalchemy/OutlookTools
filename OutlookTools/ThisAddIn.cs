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

        // Called by VSTO runtime when add-in loads
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _application = this.Application;
            LogDebug("OutlookTools v1.2.0 started.");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            LogDebug("OutlookTools v1.2.0 shut down.");
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

        #region VSTO generated
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
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
