using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ReportLineOAForDebtAndBranch
{
    public record SavedConnection(string Name, string ConnectionString);

    public static class ConnectionStore
    {
        private static readonly string FilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ReportLineOAForDebtAndBranch",
            "connections.json");

        private static readonly JsonSerializerOptions JsonOpts = new() { WriteIndented = true };

        public static List<SavedConnection> Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new();
                string json = File.ReadAllText(FilePath);
                return JsonSerializer.Deserialize<List<SavedConnection>>(json, JsonOpts) ?? new();
            }
            catch
            {
                return new();
            }
        }

        public static void Save(List<SavedConnection> connections)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(connections, JsonOpts));
        }

        public static void Upsert(string name, string connectionString)
        {
            var list = Load();
            int idx = list.FindIndex(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            var entry = new SavedConnection(name, connectionString);
            if (idx >= 0) list[idx] = entry;
            else list.Add(entry);
            Save(list);
        }

        public static void Delete(string name)
        {
            var list = Load();
            list.RemoveAll(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            Save(list);
        }
    }
}
