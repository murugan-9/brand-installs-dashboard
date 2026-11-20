# dump-ess.ps1
# Dumps EVERY .xlsx in the ESS folder to JSON (ESS_Installs.json + SNs_Config_SVL.json) into an
# isolated staging folder. Always re-dumps (fresh) so no stale data can be picked up.
# Does not touch the legacy xlsx-dumps folders and never touches git.
param(
    [string]$EssDir = 'C:\Users\MuruganVenugopal\.bob\ESS',
    [string]$Stage  = 'C:\Users\MuruganVenugopal\.bob\playground\ess-staging'
)
$ErrorActionPreference = "Stop"
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

public static class XlsxDumperEss
{
    public static int[] DumpToJson(string xlsxPath, string outDir, string essFileName, string svlFileName)
    {
        Directory.CreateDirectory(outDir);
        using (var zip = ZipFile.OpenRead(xlsxPath))
        {
            var sst        = ReadSharedStrings(zip);
            var sheetNames = ReadSheetNames(zip);

            int essIdx = -1, svlIdx = -1;
            for (int i = 0; i < sheetNames.Count; i++)
            {
                if (sheetNames[i].IndexOf("ESS Installs", StringComparison.OrdinalIgnoreCase) >= 0) essIdx = i + 1;
                if (sheetNames[i].IndexOf("SNs Config",   StringComparison.OrdinalIgnoreCase) >= 0) svlIdx = i + 1;
            }
            if (essIdx < 0) essIdx = 2;
            if (svlIdx < 0) svlIdx = 1;

            int essRows = DumpSheet(zip, essIdx, sst, Path.Combine(outDir, essFileName));
            int svlRows = DumpSheet(zip, svlIdx, sst, Path.Combine(outDir, svlFileName));
            return new int[] { essRows, svlRows };
        }
    }

    static List<string> ReadSharedStrings(ZipArchive zip)
    {
        var result = new List<string>(131072);
        var entry  = zip.GetEntry("xl/sharedStrings.xml");
        if (entry == null) return result;

        using (var stream = entry.Open())
        using (var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore }))
        {
            bool inSi = false, inT = false;
            var sb = new StringBuilder(256);
            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                        if      (reader.LocalName == "si") { inSi = true; sb.Clear(); }
                        else if (reader.LocalName == "t" && inSi) { inT = true; }
                        break;
                    case XmlNodeType.Text:
                    case XmlNodeType.SignificantWhitespace:
                        if (inT) sb.Append(reader.Value);
                        break;
                    case XmlNodeType.EndElement:
                        if      (reader.LocalName == "t")  { inT = false; }
                        else if (reader.LocalName == "si") { result.Add(sb.ToString()); inSi = false; sb.Clear(); }
                        break;
                }
            }
        }
        return result;
    }

    static List<string> ReadSheetNames(ZipArchive zip)
    {
        var names = new List<string>();
        var entry = zip.GetEntry("xl/workbook.xml");
        if (entry == null) return names;
        using (var stream = entry.Open())
        using (var reader = XmlReader.Create(stream))
        {
            while (reader.Read())
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheet")
                    names.Add(reader.GetAttribute("name") ?? "");
        }
        return names;
    }

    static int DumpSheet(ZipArchive zip, int sheetIdx, List<string> sst, string outPath)
    {
        var entryName = "xl/worksheets/sheet" + sheetIdx + ".xml";
        var entry = zip.GetEntry(entryName);
        if (entry == null) throw new Exception("Sheet not found: " + entryName);

        using (var inStream  = entry.Open())
        using (var outStream = new FileStream(outPath, FileMode.Create, FileAccess.Write, FileShare.None, 131072))
        using (var writer    = new StreamWriter(outStream, new UTF8Encoding(false), 131072))
        {
            writer.Write("{\"headers\":");
            var xmlSettings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore, IgnoreWhitespace = true };
            using (var reader = XmlReader.Create(inStream, xmlSettings))
            {
                string[] headers = null;
                bool headersWritten = false, firstDataRow = true;
                int rowCount = 0, currentRowNum = 0;
                bool inRow = false, inCell = false, inV = false, inT = false;
                string cellRef = "", cellType = "";
                var cellVal = new StringBuilder(128);
                var rowData = new Dictionary<int, string>(256);

                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:
                            switch (reader.LocalName)
                            {
                                case "row":
                                    var ra = reader.GetAttribute("r");
                                    currentRowNum = ra != null ? int.Parse(ra) : currentRowNum + 1;
                                    rowData.Clear(); inRow = true;
                                    break;
                                case "c":
                                    cellRef  = reader.GetAttribute("r") ?? "";
                                    cellType = reader.GetAttribute("t") ?? "";
                                    cellVal.Clear();
                                    inCell = true; inV = false; inT = false;
                                    if (reader.IsEmptyElement) inCell = false;
                                    break;
                                case "v": if (inCell) { inV = true; cellVal.Clear(); } break;
                                case "t": if (inCell)   inT = true; break;
                            }
                            break;
                        case XmlNodeType.Text:
                        case XmlNodeType.SignificantWhitespace:
                            if (inV || inT) cellVal.Append(reader.Value);
                            break;
                        case XmlNodeType.EndElement:
                            switch (reader.LocalName)
                            {
                                case "v": inV = false; break;
                                case "t": inT = false; break;
                                case "c":
                                    if (inCell)
                                    {
                                        string resolved = cellVal.ToString();
                                        if (cellType == "s")
                                        {
                                            int si = 0;
                                            resolved = (int.TryParse(resolved, out si) && si < sst.Count) ? sst[si] : "";
                                        }
                                        else if (cellType == "b")
                                            resolved = resolved == "1" ? "TRUE" : "FALSE";
                                        rowData[ColIndex(cellRef)] = resolved;
                                        inCell = false;
                                    }
                                    break;
                                case "row":
                                    if (!inRow) break;
                                    inRow = false;
                                    if (currentRowNum == 1)
                                    {
                                        int maxCol = 0;
                                        foreach (var k in rowData.Keys) if (k > maxCol) maxCol = k;
                                        headers = new string[maxCol + 1];
                                        for (int i = 0; i <= maxCol; i++)
                                            headers[i] = rowData.ContainsKey(i) ? rowData[i] : "";
                                        writer.Write("[");
                                        for (int i = 0; i < headers.Length; i++)
                                        { if (i > 0) writer.Write(","); writer.Write("\""); writer.Write(JEsc(headers[i])); writer.Write("\""); }
                                        writer.Write("],\"rows\":[");
                                        headersWritten = true;
                                    }
                                    else if (headers != null)
                                    {
                                        if (!firstDataRow) writer.Write(",");
                                        firstDataRow = false;
                                        writer.Write("[");
                                        for (int i = 0; i < headers.Length; i++)
                                        {
                                            if (i > 0) writer.Write(",");
                                            string v;
                                            if (rowData.TryGetValue(i, out v))
                                            { writer.Write("\""); writer.Write(JEsc(v)); writer.Write("\""); }
                                            else writer.Write("null");
                                        }
                                        writer.Write("]");
                                        rowCount++;
                                    }
                                    break;
                            }
                            break;
                    }
                }
                if (!headersWritten) writer.Write("[],\"rows\":[");
                writer.Write("]}");
                return rowCount;
            }
        }
    }

    static int ColIndex(string cellRef)
    {
        int idx = 0;
        foreach (char c in cellRef) { if (c < 'A' || c > 'Z') break; idx = idx * 26 + (c - 'A' + 1); }
        return idx - 1;
    }

    static string JEsc(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        var sb = new StringBuilder(s.Length + 4);
        foreach (char c in s)
        {
            if      (c == '"')  sb.Append("\\\"");
            else if (c == '\\') sb.Append("\\\\");
            else if (c == '\n' || c == '\r' || c == '\t') sb.Append(" ");
            else if (c < 0x20)  sb.AppendFormat("\\u{0:x4}", (int)c);
            else sb.Append(c);
        }
        return sb.ToString();
    }
}
'@ -ReferencedAssemblies "System.IO.Compression","System.IO.Compression.FileSystem","System.Xml"

Write-Host "C# class compiled OK" -ForegroundColor Green
if (-not (Test-Path $EssDir)) { throw "ESS folder not found: $EssDir" }
if (Test-Path $Stage) { Remove-Item $Stage -Recurse -Force }
New-Item -ItemType Directory -Path $Stage -Force | Out-Null

$files = Get-ChildItem $EssDir -Filter *.xlsx -File | Where-Object { $_.Name -notlike '~$*' } | Sort-Object Name
Write-Host ("Found {0} xlsx file(s) in {1}" -f $files.Count, $EssDir) -ForegroundColor Cyan
$summary = @()
foreach ($f in $files) {
    $name   = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    $outDir = Join-Path $Stage $name
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $c  = [XlsxDumperEss]::DumpToJson($f.FullName, $outDir, "ESS_Installs.json", "SNs_Config_SVL.json")
    $sw.Stop()
    Write-Host ("  {0,-24} ESS={1,6}  SVL={2,7}  {3,5:N1}s" -f $name, $c[0], $c[1], $sw.Elapsed.TotalSeconds) -ForegroundColor Green
    $summary += [ordered]@{ program=$name; file=$f.Name; sizeBytes=$f.Length; modified=$f.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'); essRows=$c[0]; svlRows=$c[1] }
}
[System.IO.File]::WriteAllText((Join-Path $Stage '_summary.json'), ($summary | ConvertTo-Json -Depth 4), (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Staging complete: $Stage" -ForegroundColor Cyan