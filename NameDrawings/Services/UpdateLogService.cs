using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace EliteSheets.Services
{
    /// <summary>
    /// "What's new" popup, the same mechanism as the other RK Tools add-ins (Dali, ProSchedules, Revit.Nr):
    /// reads EliteSheets_CHANGELOG.json from the team Dropbox (…\Pluginad\UuendusteInfo) and, when it lists
    /// versions newer than the last one this Windows user has seen, shows them once.
    /// The last shown version is stored per user in %AppData%\EliteSheets\state.json.
    /// </summary>
    public static class UpdateLogService
    {
        private const string AppName = "EliteSheets";
        private const string StateFileName = "state.json";
        private const string ChangelogFileName = "EliteSheets_CHANGELOG.json";
        private const string EnvOverride = "ELITESHEETS_CHANGELOG_PATH";

        public class Changelog
        {
            [JsonProperty("versions")]
            public List<VersionEntry> Versions { get; set; }
        }

        public class VersionEntry
        {
            [JsonProperty("version")]
            public string Version { get; set; }

            [JsonProperty("released")]
            public string Released { get; set; }

            [JsonProperty("developer")]
            public string Developer { get; set; }

            /// <summary>A string or a list of strings in the JSON.</summary>
            [JsonProperty("notes")]
            public object NotesRaw { get; set; }

            [JsonIgnore]
            public List<string> Notes
            {
                get
                {
                    if (NotesRaw is string s) return new List<string> { s };
                    if (NotesRaw is Newtonsoft.Json.Linq.JArray arr) return arr.ToObject<List<string>>();
                    return new List<string>();
                }
            }
        }

        private class LocalState
        {
            public string LastShownVersion { get; set; } = "0.0.0";
        }

        /// <summary>Shows the unseen versions (newest first) on top of <paramref name="owner"/>, then marks them seen.</summary>
        public static void CheckAndShow(Window owner)
        {
            try
            {
                if (!TryResolveChangelogPath(out var path)) return;

                var changelog = JsonConvert.DeserializeObject<Changelog>(File.ReadAllText(path));
                if (changelog?.Versions == null || changelog.Versions.Count == 0) return;

                var state = LoadState();
                var lastShown = ParseVersion(state.LastShownVersion);

                var unseen = changelog.Versions
                    .Where(v => ParseVersion(v.Version) > lastShown)
                    .OrderByDescending(v => ParseVersion(v.Version))
                    .ToList();
                if (unseen.Count == 0) return;

                // Modeless: a modal loop here can suspend Revit's dispatcher during startup external events.
                var window = new ChangelogWindow(owner, unseen);
                window.Show();

                state.LastShownVersion = unseen[0].Version;
                SaveState(state);
            }
            catch (Exception ex)
            {
                Logger.Log("Update log check failed.", ex);
            }
        }

        private static bool TryResolveChangelogPath(out string path)
        {
            path = Environment.GetEnvironmentVariable(EnvOverride);
            if (!string.IsNullOrEmpty(path) && File.Exists(path)) return true;

            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            foreach (var root in new[] { "EULE Dropbox", "Dropbox" })
            {
                path = Path.Combine(userProfile, root, "0_EULE  Team folder (kogu kollektiiv)", "02_EULE REVIT TEMPLATE",
                                    "099-scriptid", "Pluginad", "UuendusteInfo", ChangelogFileName);
                if (File.Exists(path)) return true;
            }

            path = null;
            return false;
        }

        private static string StatePath =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), AppName, StateFileName);

        private static LocalState LoadState()
        {
            try
            {
                if (File.Exists(StatePath))
                    return JsonConvert.DeserializeObject<LocalState>(File.ReadAllText(StatePath)) ?? new LocalState();
            }
            catch (Exception ex)
            {
                Logger.Log("Failed to read update log state.", ex);
            }
            return new LocalState();
        }

        private static void SaveState(LocalState state)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(StatePath));
                File.WriteAllText(StatePath, JsonConvert.SerializeObject(state, Formatting.Indented));
            }
            catch (Exception ex)
            {
                Logger.Log("Failed to save update log state.", ex);
            }
        }

        private static Version ParseVersion(string v)
        {
            v = (v ?? string.Empty).Replace("v", "").Trim();
            return Version.TryParse(v, out var result) ? result : new Version(0, 0, 0);
        }
    }
}
