using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using CablePlan.Models;

namespace CablePlan.Services;

public static class CableStorage
{
    private static readonly JsonSerializerOptions Opts = new()
    {
        WriteIndented = true
    };

    public static void Save(string jsonPath, PlanData data)
    {
        var json = JsonSerializer.Serialize(data, Opts);
        File.WriteAllText(jsonPath, json, Encoding.UTF8);
    }

    public static PlanData LoadOrNew(string jsonPath)
    {
        if (!File.Exists(jsonPath)) return new PlanData();
        var json = File.ReadAllText(jsonPath, Encoding.UTF8);
        return JsonSerializer.Deserialize<PlanData>(json, Opts) ?? new PlanData();
    }

    public static string ComputeFileHash(string filePath)
    {
        using var sha = SHA256.Create();
        using var fs = File.OpenRead(filePath);
        var hash = sha.ComputeHash(fs);
        return Convert.ToHexString(hash);
    }
}
