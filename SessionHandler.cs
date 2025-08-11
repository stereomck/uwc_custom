// TARGET:dummy.exe
// START_IN:
using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Components;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;

public class Default : ScriptBase
{
    void Execute()
    {
        Wait(2);

        try
        {
            Log("=== Test 1: Search for Login button ===");
        }
        catch (Exception ex)
        {
            Log($"Error in Execute: {ex.Message}");
            Log($"Stack trace: {ex.StackTrace}");
        }
    }
    
    private void LogResults(string searchTerm, List<OCRMatch> results)
    {
        Log($"Search for '{searchTerm}': Found {results.Count} matches");
        for (int i = 0; i < results.Count; i++)
        {
            var match = results[i];
            Log($"  Match {i}: '{match.Text}' at ({match.CenterX}, {match.CenterY}) confidence: {match.Confidence}");
        }
    }
}

public class OCRMatch
{
    public string Text { get; set; }
    public int Left { get; set; }
    public int Top { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public int CenterX { get; set; }
    public int CenterY { get; set; }
    public double Confidence { get; set; }
    public string Type { get; set; }
    public int WordCount { get; set; }

    public OCRMatch()
    {
        Text = "";
        Type = "";
    }
}

public class OCRWorkflow
{
    private readonly string OCR_SCRIPT_PATH = @"C:\Users\mkent-admin\Documents\GitHub\uwc_custom\OCRFunctions.ps1";
    private int _screenshotCounter = 0;

    /// <summary>
    /// Activate window by partial title match
    /// </summary>
    private bool ActivateWindowByTitle(string partialTitle, int waitTimeMs = 2000)
    {
        try
        {
            string command = $@"
                . '{OCR_SCRIPT_PATH}'
                
                # Find windows by title
                $windows = Find-WindowByTitle -Title '{partialTitle}'
                
                if ($windows.Count -eq 0) {{
                    throw 'No windows found with title containing: {partialTitle}'
                }}
                
                # Activate the first found window
                $targetWindow = $windows[0]
                [Win32]::SetForegroundWindow($targetWindow)
                
                # Wait and verify activation
                Start-Sleep -Milliseconds 500
                
                Write-Output 'SUCCESS'";
            
            string result = RunPoSHCommand(command);
            
            // Additional wait for window to become active
            Thread.Sleep(waitTimeMs);
            
            return result.Trim().Contains("SUCCESS");
        }
        catch (Exception ex)
        {
            throw new Exception($"Window activation failed: {ex.Message}");
        }
    }
        /// <summary>
    /// Click at specific coordinates
    /// </summary>
    private bool ClickAtCoordinates(int x, int y)
    {
        try
        {
            string command = $@"
                . '{OCR_SCRIPT_PATH}'
                Click-Coordinates -X {x} -Y {y}
                Write-Output 'CLICK_SUCCESS'";
            
            string result = RunPoSHCommand(command);
            return result.Trim().Contains("CLICK_SUCCESS");
        }
        catch (Exception ex)
        {
            throw new Exception($"Click operation failed: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Execute PowerShell command with enhanced error handling
    /// </summary>
    private string RunPoSHCommand(string command)
    {
        try
        {
            if (!File.Exists(OCR_SCRIPT_PATH))
            {
                throw new FileNotFoundException($"OCR script not found at: {OCR_SCRIPT_PATH}");
            }

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-ExecutionPolicy Bypass -NoProfile -Command \"{command.Replace("\"", "`\"")}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process powershell = new Process { StartInfo = startInfo })
            {
                powershell.Start();
                
                string output = powershell.StandardOutput.ReadToEnd();
                string errors = powershell.StandardError.ReadToEnd();
                
                powershell.WaitForExit();
                
                if (powershell.ExitCode != 0)
                {
                    throw new Exception($"PowerShell command failed (Exit code: {powershell.ExitCode}): {errors}");
                }
                
                return output;
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to execute PowerShell command: {ex.Message}");
        }
    }
}

// Simple JSON parser for OCR results
public static class SimpleJsonParser
{
    public static List<OCRMatch> ParseOCRResults(string json)
    {
        var results = new List<OCRMatch>();
        
        if (string.IsNullOrWhiteSpace(json) || json.Trim() == "null" || json.Trim() == "[]")
        {
            return results;
        }

        try
        {
            json = json.Trim();
            
            if (json.StartsWith("[") && json.EndsWith("]"))
            {
                var objects = ExtractJsonObjects(json);
                foreach (var obj in objects)
                {
                    var match = ParseSingleOCRMatch(obj);
                    if (match != null)
                    {
                        results.Add(match);
                    }
                }
            }
            else if (json.StartsWith("{") && json.EndsWith("}"))
            {
                var match = ParseSingleOCRMatch(json);
                if (match != null)
                {
                    results.Add(match);
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to parse JSON: {ex.Message}. JSON: {json}");
        }

        return results;
    }

    private static List<string> ExtractJsonObjects(string jsonArray)
    {
        var objects = new List<string>();
        string content = jsonArray.Substring(1, jsonArray.Length - 2).Trim();
        
        if (string.IsNullOrEmpty(content))
        {
            return objects;
        }

        int braceCount = 0;
        int startIndex = 0;
        bool inString = false;
        bool escapeNext = false;

        for (int i = 0; i < content.Length; i++)
        {
            char c = content[i];

            if (escapeNext)
            {
                escapeNext = false;
                continue;
            }

            if (c == '\\')
            {
                escapeNext = true;
                continue;
            }

            if (c == '"' && !escapeNext)
            {
                inString = !inString;
                continue;
            }

            if (!inString)
            {
                if (c == '{')
                {
                    braceCount++;
                }
                else if (c == '}')
                {
                    braceCount--;
                    
                    if (braceCount == 0)
                    {
                        string obj = content.Substring(startIndex, i - startIndex + 1).Trim();
                        objects.Add(obj);
                        
                        while (i + 1 < content.Length && (content[i + 1] == ',' || char.IsWhiteSpace(content[i + 1])))
                        {
                            i++;
                        }
                        startIndex = i + 1;
                    }
                }
            }
        }

        return objects;
    }

    private static OCRMatch ParseSingleOCRMatch(string jsonObject)
    {
        try
        {
            var match = new OCRMatch();

            match.Text = ExtractStringValue(jsonObject, "Text") ?? "";
            match.Left = ExtractIntValue(jsonObject, "Left");
            match.Top = ExtractIntValue(jsonObject, "Top");
            match.Width = ExtractIntValue(jsonObject, "Width");
            match.Height = ExtractIntValue(jsonObject, "Height");
            match.CenterX = ExtractIntValue(jsonObject, "CenterX");
            match.CenterY = ExtractIntValue(jsonObject, "CenterY");
            match.Confidence = ExtractDoubleValue(jsonObject, "Confidence");
            match.Type = ExtractStringValue(jsonObject, "Type") ?? "";
            match.WordCount = ExtractIntValue(jsonObject, "WordCount");

            return match;
        }
        catch (Exception ex)
        {
            throw new Exception($"Failed to parse OCR match object: {ex.Message}");
        }
    }

    private static string ExtractStringValue(string json, string propertyName)
    {
        string pattern = $"\"({propertyName})\"\\s*:\\s*\"([^\"\\\\]*(\\\\.[^\"\\\\]*)*)\"";
        var match = Regex.Match(json, pattern, RegexOptions.IgnoreCase);
        
        if (match.Success)
        {
            string value = match.Groups[2].Value;
            value = value.Replace("\\\"", "\"")
                        .Replace("\\\\", "\\")
                        .Replace("\\n", "\n")
                        .Replace("\\r", "\r")
                        .Replace("\\t", "\t");
            return value;
        }
        
        return null;
    }

    private static int ExtractIntValue(string json, string propertyName)
    {
        string pattern = $"\"({propertyName})\"\\s*:\\s*(-?\\d+)";
        var match = Regex.Match(json, pattern, RegexOptions.IgnoreCase);
        
        if (match.Success && int.TryParse(match.Groups[2].Value, out int value))
        {
            return value;
        }
        
        return 0;
    }

    private static double ExtractDoubleValue(string json, string propertyName)
    {
        string pattern = $"\"({propertyName})\"\\s*:\\s*(-?\\d+(?:\\.\\d+)?)";
        var match = Regex.Match(json, pattern, RegexOptions.IgnoreCase);
        
        if (match.Success && double.TryParse(match.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
        {
            return value;
        }
        
        return 0.0;
    }
}
