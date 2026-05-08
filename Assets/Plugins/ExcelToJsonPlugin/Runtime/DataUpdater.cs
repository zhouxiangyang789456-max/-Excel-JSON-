using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using UnityEngine;
using UnityEngine.Networking;

namespace ExcelToJsonPlugin.Runtime
{
    /// <summary>
    /// Runtime data hot-update manager.
    /// Checks a CDN for newer versions of data tables and downloads JSON updates.
    ///
    /// Usage:
    ///   var updater = new DataUpdater("https://cdn.example.com/game-data/");
    ///   yield return updater.CheckAndUpdateAll(OnTableUpdated);
    /// </summary>
    public class DataUpdater
    {
        private readonly string cdnBaseUrl;
        private readonly string localCachePath;

        /// <param name="cdnBaseUrl">CDN base URL, e.g. "https://cdn.example.com/game-data/"</param>
        /// <param name="localCachePath">Local cache path, defaults to Application.persistentDataPath + "/Data/"</param>
        public DataUpdater(string cdnBaseUrl, string localCachePath = null)
        {
            this.cdnBaseUrl = cdnBaseUrl.TrimEnd('/') + "/";
            this.localCachePath = localCachePath
                ?? Path.Combine(Application.persistentDataPath, "Data");
            Directory.CreateDirectory(this.localCachePath);
        }

        /// <summary>
        /// Fetch the version manifest from CDN.
        /// Returns: { "Weapon": "a1b2c3", "Skill": "d4e5f6" }
        /// </summary>
        public IEnumerator FetchManifest(Action<Dictionary<string, string>> onComplete,
            Action<string> onError = null)
        {
            var url = cdnBaseUrl + "version.json";
            using (var req = UnityWebRequest.Get(url))
            {
                yield return req.SendWebRequest();

                if (req.result != UnityWebRequest.Result.Success)
                {
                    onError?.Invoke($"Failed to fetch manifest: {req.error}");
                    yield break;
                }

                try
                {
                    var manifest = ParseManifest(req.downloadHandler.text);
                    onComplete?.Invoke(manifest);
                }
                catch (Exception ex)
                {
                    onError?.Invoke($"Failed to parse manifest: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Download updated JSON for a table.
        /// </summary>
        public IEnumerator DownloadTableJson(string tableName, string remoteHash,
            Action<string> onComplete, Action<string> onError = null)
        {
            var url = $"{cdnBaseUrl}{tableName}.{remoteHash}.json";
            using (var req = UnityWebRequest.Get(url))
            {
                yield return req.SendWebRequest();

                if (req.result != UnityWebRequest.Result.Success)
                {
                    onError?.Invoke($"Failed to download {tableName}: {req.error}");
                    yield break;
                }

                var json = req.downloadHandler.text;
                var path = Path.Combine(localCachePath, $"{tableName}.json");
                File.WriteAllText(path, json, Encoding.UTF8);

                // Save hash for future comparison
                PlayerPrefs.SetString($"data_hash_{tableName}", remoteHash);

                onComplete?.Invoke(json);
            }
        }

        /// <summary>
        /// Check for updates and download all changed tables.
        /// </summary>
        public IEnumerator CheckAndUpdateAll(
            Action<string, int, int> onProgress = null,
            Action<string, string> onTableUpdated = null,
            Action<string> onError = null)
        {
            Dictionary<string, string> manifest = null;
            string error = null;

            yield return FetchManifest(
                m => manifest = m,
                e => error = e);

            if (error != null || manifest == null)
            {
                onError?.Invoke(error ?? "Manifest is null");
                yield break;
            }

            var updates = new List<KeyValuePair<string, string>>();
            foreach (var kv in manifest)
            {
                var localHash = PlayerPrefs.GetString($"data_hash_{kv.Key}", "");
                if (localHash != kv.Value)
                    updates.Add(kv);
            }

            if (updates.Count == 0)
            {
                Debug.Log("[DataUpdater] All tables up to date.");
                yield break;
            }

            Debug.Log($"[DataUpdater] Found {updates.Count} tables to update.");

            for (int i = 0; i < updates.Count; i++)
            {
                var update = updates[i];
                onProgress?.Invoke(update.Key, i + 1, updates.Count);

                bool done = false;
                string json = null;
                string err = null;

                yield return DownloadTableJson(update.Key, update.Value,
                    j => { json = j; done = true; },
                    e => { err = e; done = true; });

                while (!done) yield return null;

                if (json != null)
                {
                    onTableUpdated?.Invoke(update.Key, json);
                    Debug.Log($"[DataUpdater] Updated: {update.Key} ({update.Value.Substring(0, 8)}...)");
                }
                else
                {
                    Debug.LogWarning($"[DataUpdater] Failed to update {update.Key}: {err}");
                }
            }
        }

        /// <summary>
        /// Load locally cached JSON for a table.
        /// </summary>
        public string LoadCachedJson(string tableName)
        {
            var path = Path.Combine(localCachePath, $"{tableName}.json");
            return File.Exists(path) ? File.ReadAllText(path) : null;
        }

        /// <summary>
        /// Clear all locally cached data.
        /// </summary>
        public void ClearCache()
        {
            if (Directory.Exists(localCachePath))
            {
                foreach (var f in Directory.GetFiles(localCachePath, "*.json"))
                    File.Delete(f);
            }
        }

        /// <summary>
        /// Compute MD5 hash of a string (for version comparison).
        /// </summary>
        public static string ComputeHash(string content)
        {
            using (var md5 = MD5.Create())
            {
                var bytes = md5.ComputeHash(Encoding.UTF8.GetBytes(content));
                return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant();
            }
        }

        private Dictionary<string, string> ParseManifest(string json)
        {
            var result = new Dictionary<string, string>();
            json = json.Trim();
            if (!json.StartsWith("{") || !json.EndsWith("}"))
                return result;

            json = json.Substring(1, json.Length - 2);
            var entries = json.Split(',');
            foreach (var entry in entries)
            {
                var parts = entry.Split(':');
                if (parts.Length >= 2)
                {
                    var key = parts[0].Trim().Trim('"');
                    var val = parts[1].Trim().Trim('"');
                    if (!string.IsNullOrEmpty(key))
                        result[key] = val;
                }
            }
            return result;
        }
    }
}
