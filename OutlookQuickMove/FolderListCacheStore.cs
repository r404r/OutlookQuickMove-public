using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OutlookQuickMove
{
    /// <summary>
    /// Disk-backed cache for folder candidates. The payload contains only strings, never COM
    /// objects, so it is safe to reuse after Outlook restarts. It is keyed by the same store-filter
    /// signature used by the in-memory cache.
    /// </summary>
    internal static class FolderListCacheStore
    {
        private const string CacheFileName = "folder-list-cache.txt";
        private const string Version = "1";
        private static readonly object Gate = new object();

        public static List<FolderCandidate> Load(string signature, TimeSpan maxAge, out long cacheAgeMs)
        {
            var folders = new List<FolderCandidate>();
            cacheAgeMs = -1;
            if (string.IsNullOrEmpty(signature))
            {
                return folders;
            }

            try
            {
                lock (Gate)
                {
                    var path = GetCachePath();
                    if (!File.Exists(path))
                    {
                        return folders;
                    }

                    var lines = File.ReadAllLines(path, Encoding.UTF8);
                    if (lines.Length < 3)
                    {
                        return folders;
                    }

                    if (!string.Equals(lines[0], Version, StringComparison.Ordinal))
                    {
                        return folders;
                    }

                    if (!string.Equals(Decode(lines[1]), signature, StringComparison.Ordinal))
                    {
                        return folders;
                    }

                    long ticks;
                    if (!long.TryParse(lines[2], out ticks))
                    {
                        QuickMoveLog.Write("ignored folder list cache because its timestamp is invalid.");
                        return folders;
                    }

                    var savedAtUtc = new DateTime(ticks, DateTimeKind.Utc);
                    var age = DateTime.UtcNow - savedAtUtc;
                    if (age < TimeSpan.Zero)
                    {
                        QuickMoveLog.Write("ignored folder list cache because its timestamp is in the future.");
                        return folders;
                    }

                    cacheAgeMs = (long)age.TotalMilliseconds;
                    if (age > maxAge)
                    {
                        QuickMoveLog.Write("ignored folder list cache because it expired: ageMs="
                            + cacheAgeMs + ", maxAgeMs=" + (long)maxAge.TotalMilliseconds + ".");
                        return folders;
                    }

                    for (int i = 3; i < lines.Length; i++)
                    {
                        var folder = ParseFolder(lines[i]);
                        if (folder != null)
                        {
                            folders.Add(folder);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to load folder list cache.", ex);
                return new List<FolderCandidate>();
            }

            if (folders.Count > 0)
            {
                QuickMoveLog.Write("loaded " + folders.Count + " folder candidate(s) from disk cache: ageMs="
                    + cacheAgeMs + ".");
            }

            return folders;
        }

        public static void Save(string signature, List<FolderCandidate> folders)
        {
            if (string.IsNullOrEmpty(signature) || folders == null || folders.Count == 0)
            {
                return;
            }

            try
            {
                var path = GetCachePath();
                var builder = new StringBuilder();
                builder.AppendLine(Version);
                builder.AppendLine(Encode(signature));
                builder.AppendLine(DateTime.UtcNow.Ticks.ToString());
                foreach (var folder in folders)
                {
                    builder.Append(Encode(folder.DisplayPath)).Append('\t')
                        .Append(Encode(folder.EntryId)).Append('\t')
                        .Append(Encode(folder.StoreId)).Append('\n');
                }

                lock (Gate)
                {
                    WriteAllTextAtomically(path, builder.ToString());
                }

                QuickMoveLog.Write("saved " + folders.Count + " folder candidate(s) to disk cache.");
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to save folder list cache.", ex);
            }
        }

        public static void Clear()
        {
            try
            {
                lock (Gate)
                {
                    var path = GetCachePath();
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                        QuickMoveLog.Write("cleared folder list cache.");
                    }
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to clear folder list cache.", ex);
            }
        }

        private static FolderCandidate ParseFolder(string line)
        {
            try
            {
                var fields = line.Split('\t');
                if (fields.Length < 3)
                {
                    return null;
                }

                var displayPath = Decode(fields[0]);
                var entryId = Decode(fields[1]);
                var storeId = Decode(fields[2]);
                if (string.IsNullOrEmpty(displayPath) || string.IsNullOrEmpty(entryId) || string.IsNullOrEmpty(storeId))
                {
                    return null;
                }

                return new FolderCandidate(displayPath, entryId, storeId);
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("Failed to parse folder list cache record.", ex);
                return null;
            }
        }

        private static string GetCachePath()
        {
            return QuickMovePaths.DataFile(CacheFileName);
        }

        private static void WriteAllTextAtomically(string path, string content)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            var tempPath = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(tempPath, content, Encoding.UTF8);
                if (File.Exists(path))
                {
                    File.Replace(tempPath, path, null);
                }
                else
                {
                    File.Move(tempPath, path);
                }
            }
            finally
            {
                TryDelete(tempPath);
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        private static string Encode(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static string Decode(string value)
        {
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value ?? string.Empty));
            }
            catch (FormatException)
            {
                return string.Empty;
            }
        }
    }
}
