using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ReportLineOAForDebtAndBranch
{
    public record TabSession(string Name, string QueryText);

    public static class SessionStore
    {
        private static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ReportLineOAForDebtAndBranch", "session.json");

        private static readonly JsonSerializerOptions Opts = new() { WriteIndented = true };

        public static List<TabSession> Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new();
                return JsonSerializer.Deserialize<List<TabSession>>(
                    File.ReadAllText(FilePath), Opts) ?? new();
            }
            catch { return new(); }
        }

        public static void Save(List<TabSession> sessions)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(sessions, Opts));
        }
    }
}
