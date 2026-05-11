using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ReportLineOAForDebtAndBranch
{
    public record HistoryEntry(
        string Id,
        string Sql,
        DateTime ExecutedAt,
        bool IsPinned
    );

    public static class HistoryStore
    {
        private static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ReportLineOAForDebtAndBranch", "history.json");

        private const int MaxUnpinned = 200;
        private static readonly JsonSerializerOptions Opts = new() { WriteIndented = true };

        public static List<HistoryEntry> Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new();
                return JsonSerializer.Deserialize<List<HistoryEntry>>(
                    File.ReadAllText(FilePath), Opts) ?? new();
            }
            catch { return new(); }
        }

        public static void Save(List<HistoryEntry> entries)
        {
            // Pinned always kept; unpinned capped at MaxUnpinned
            var pinned   = entries.FindAll(e => e.IsPinned);
            var unpinned = entries.FindAll(e => !e.IsPinned);
            if (unpinned.Count > MaxUnpinned)
                unpinned = unpinned.GetRange(0, MaxUnpinned);

            var all = new List<HistoryEntry>(pinned);
            all.AddRange(unpinned);

            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(all, Opts));
        }

        public static HistoryEntry AddEntry(string sql)
        {
            var list    = Load();
            string norm = sql.Trim();

            // Deduplicate: if same SQL exists in recent 10 unpinned, just bump its timestamp
            int recent = Math.Min(10, list.Count);
            for (int i = 0; i < recent; i++)
            {
                if (!list[i].IsPinned && list[i].Sql.Trim() == norm)
                {
                    var updated = list[i] with { ExecutedAt = DateTime.Now };
                    list.RemoveAt(i);
                    list.Insert(0, updated);
                    Save(list);
                    return updated;
                }
            }

            var entry = new HistoryEntry(Guid.NewGuid().ToString("N"), sql, DateTime.Now, false);
            list.Insert(0, entry);
            Save(list);
            return entry;
        }

        public static void SetPinned(string id, bool pinned)
        {
            var list = Load();
            int idx  = list.FindIndex(e => e.Id == id);
            if (idx < 0) return;
            list[idx] = list[idx] with { IsPinned = pinned };
            Save(list);
        }

        public static void Delete(string id)
        {
            var list = Load();
            list.RemoveAll(e => e.Id == id);
            Save(list);
        }

        public static void ClearUnpinned()
        {
            var list = Load();
            list.RemoveAll(e => !e.IsPinned);
            Save(list);
        }

        public static string RelativeTime(DateTime dt)
        {
            var diff = DateTime.Now - dt;
            if (diff.TotalSeconds < 60)  return "Just now";
            if (diff.TotalMinutes < 60)  return $"{(int)diff.TotalMinutes}m ago";
            if (diff.TotalHours   < 24)  return $"{(int)diff.TotalHours}h ago";
            if (diff.TotalDays    < 7)   return $"{(int)diff.TotalDays}d ago";
            return dt.ToString("dd/MM/yy");
        }
    }
}
