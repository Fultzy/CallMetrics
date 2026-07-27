
using CallMetrics;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Controls;

public static class Updates
{
    private const string Owner = "Fultzy";
    private const string Repo = "CallMetrics";

    // optional: set a personal access token if you hit rate limits
    private const string GitHubToken = null;

    public static async Task CheckForUpdates()
    {
        var json = await GetLatestJson();
        var doc = JsonDocument.Parse(json);

        if (true)
        {
            var root = doc.RootElement;
            var tag = root.GetProperty("tag_name").GetString() ?? "unknown";
            var htmlUrl = root.GetProperty("html_url").GetString() ?? "";


            var msg = $"A new version of CallMetrics is available: {tag}\n\n" +
                      "Please visit the GitHub releases page to download the latest version.";
        }

        doc.Dispose();
    }

    private static async Task<string> GetLatestJson()
    {
        using var http = new HttpClient();
        http.DefaultRequestHeaders.UserAgent.ParseAdd("CallMetrics-Updater/1.0");
        if (!string.IsNullOrEmpty(GitHubToken))
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", GitHubToken);

        var url = $"https://api.github.com/repos/{Owner}/{Repo}/releases/latest";
        using var resp = await http.GetAsync(url);
        if (!resp.IsSuccessStatusCode) return null;

        var json = await resp.Content.ReadAsStringAsync();
        return json;
    }

    private static bool NewVersionExists(string json)
    {
        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        if (!root.TryGetProperty("tag_name", out var tagElem)) return false;
        var tag = tagElem.GetString() ?? "";
        tag.TrimStart('v', 'V');

        if (!Version.TryParse(tag, out var latest)) return false;
        if (!Version.TryParse(App.Version, out var local)) return false;

        return latest > local;
    }
}