using System;
using System.IO;

namespace OutlookTools.Settings
{
    /// <summary>
    /// OutlookTools — Settings Manager
    /// Stores settings in a simple JSON file in LocalAppData.
    /// No registry. No admin rights needed. Per-user only.
    /// 
    /// File: %LOCALAPPDATA%\OutlookTools\settings.json
    /// Format: Simple key-value JSON.
    /// </summary>
    public static class SettingsManager
    {
        private static readonly string SettingsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OutlookTools");

        private static readonly string SettingsFile = Path.Combine(SettingsDir, "settings.json");

        // Default values
        private static int _archiveAgeDays = 90;
        private static int _archiveHour = 6;
        private static int _followUpDays = 3;
        private static bool _autoArchiveEnabled = true;
        private static bool _autoReminderEnabled = true;
        private static bool _debugLogEnabled = false;
        private static bool _startupNotification = false;
        private static DateTime _lastDailyRun = DateTime.MinValue;
        private static bool _loaded = false;

        static SettingsManager()
        {
            Load();
        }

        public static int GetArchiveAgeDays() => _archiveAgeDays;
        public static int GetArchiveHour() => _archiveHour;
        public static int GetFollowUpDays() => _followUpDays;
        public static bool GetAutoArchiveEnabled() => _autoArchiveEnabled;
        public static bool GetAutoReminderEnabled() => _autoReminderEnabled;
        public static bool GetDebugLogEnabled() => _debugLogEnabled;
        public static bool GetStartupNotification() => _startupNotification;
        public static DateTime GetLastDailyRun() => _lastDailyRun;

        public static void SetArchiveAgeDays(int val) { _archiveAgeDays = val; Save(); }
        public static void SetArchiveHour(int val) { _archiveHour = val; Save(); }
        public static void SetFollowUpDays(int val) { _followUpDays = val; Save(); }
        public static void SetAutoArchiveEnabled(bool val) { _autoArchiveEnabled = val; Save(); }
        public static void SetAutoReminderEnabled(bool val) { _autoReminderEnabled = val; Save(); }
        public static void SetDebugLogEnabled(bool val) { _debugLogEnabled = val; Save(); }
        public static void SetStartupNotification(bool val) { _startupNotification = val; Save(); }
        public static void SetLastDailyRun(DateTime val) { _lastDailyRun = val; Save(); }

        private static void Load()
        {
            if (_loaded) return;
            try
            {
                if (!File.Exists(SettingsFile)) { _loaded = true; return; }

                string json = File.ReadAllText(SettingsFile);
                // Minimal JSON parser — no external dependencies
                _archiveAgeDays = GetInt(json, "archiveAgeDays", 90);
                _archiveHour = GetInt(json, "archiveHour", 6);
                _followUpDays = GetInt(json, "followUpDays", 3);
                _autoArchiveEnabled = GetBool(json, "autoArchiveEnabled", true);
                _autoReminderEnabled = GetBool(json, "autoReminderEnabled", true);
                _debugLogEnabled = GetBool(json, "debugLogEnabled", false);
                _startupNotification = GetBool(json, "startupNotification", false);
                _lastDailyRun = GetDateTime(json, "lastDailyRun", DateTime.MinValue);
                _loaded = true;
            }
            catch
            {
                // If settings file is corrupted, use defaults
                _loaded = true;
            }
        }

        private static void Save()
        {
            try
            {
                Directory.CreateDirectory(SettingsDir);
                string json = $@"{{
  ""archiveAgeDays"": {_archiveAgeDays},
  ""archiveHour"": {_archiveHour},
  ""followUpDays"": {_followUpDays},
  ""autoArchiveEnabled"": {(_autoArchiveEnabled ? "true" : "false")},
  ""autoReminderEnabled"": {(_autoReminderEnabled ? "true" : "false")},
  ""debugLogEnabled"": {(_debugLogEnabled ? "true" : "false")},
  ""startupNotification"": {(_startupNotification ? "true" : "false")},
  "lastDailyRun": "{_lastDailyRun:O}"
}}";
                File.WriteAllText(SettingsFile, json);
            }
            catch { /* settings save must not throw */ }
        }

        // ===== Minimal JSON helpers (no external dependency) =====

        private static int GetInt(string json, string key, int defaultVal)
        {
            try
            {
                int idx = json.IndexOf($"\"{key}\"");
                if (idx < 0) return defaultVal;
                int colon = json.IndexOf(':', idx);
                int end = json.IndexOfAny(new[] { ',', '}' }, colon);
                return int.Parse(json.Substring(colon + 1, end - colon - 1).Trim());
            }
            catch { return defaultVal; }
        }

        private static bool GetBool(string json, string key, bool defaultVal)
        {
            try
            {
                int idx = json.IndexOf($"\"{key}\"");
                if (idx < 0) return defaultVal;
                int colon = json.IndexOf(':', idx);
                int end = json.IndexOfAny(new[] { ',', '}' }, colon);
                string val = json.Substring(colon + 1, end - colon - 1).Trim();
                return val == "true";
            }
            catch { return defaultVal; }
        }

        private static DateTime GetDateTime(string json, string key, DateTime defaultVal)
        {
            try
            {
                int idx = json.IndexOf($"\"{key}\"");
                if (idx < 0) return defaultVal;
                int quote1 = json.IndexOf('"', idx + key.Length + 3);
                int quote2 = json.IndexOf('"', quote1 + 1);
                string val = json.Substring(quote1 + 1, quote2 - quote1 - 1);
                return DateTime.Parse(val);
            }
            catch { return defaultVal; }
        }
    }
}
