# dump-3nov.ps1
# Dumps the 4 "3rd Nov" xlsx files to JSON for the dashboard refresh.
# Reuses the compiled C# streamer from dump-6oct-fast.ps1 ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Â same logic, new file list.

$ErrorActionPreference = "Stop"

$playground = "C:\Users\MuruganVenugopal\.bob\playground"
$dumpBase   = "C:\Users\MuruganVenugopal\.bob\playground\.bob\tmp\xlsx-dumps\Brands_Installs Data\.bob\tmp\xlsx-dumps"

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

public static class XlsxDumper2
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
                if (!headersWritten) writer.Write("[]");
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

$files = @(
    @{ XlsxName="P11 (2025) - 3rd Nov.xlsx";          DumpLabel="P11 (2025) - 3rd Nov" },
    @{ XlsxName="FlashSystem - 3rd Nov.xlsx";         DumpLabel="FLASHSYSTEM - 3rd Nov" },
    @{ XlsxName="z Mid Range (2026) - 3rd Nov.xlsx";   DumpLabel="z Mid Range (2026) - 3rd Nov" },
    @{ XlsxName="P11 Balcones (2026) - 3rd Nov.xlsx";  DumpLabel="P11 Balcones (2026) - 3rd Nov" }
)

foreach ($f in $files) {
    $xlsxPath = Join-Path $playground $f.XlsxName
    Write-Host ""
    Write-Host "=== $($f.DumpLabel) ===" -ForegroundColor Cyan

    if (-not (Test-Path $xlsxPath)) { Write-Host "  ERROR: Not found: $xlsxPath" -ForegroundColor Red; continue }

    $md5    = [System.Security.Cryptography.MD5]::Create()
    $stream = [System.IO.File]::OpenRead($xlsxPath)
    $hash   = ($md5.ComputeHash($stream) | ForEach-Object { $_.ToString("x2") }) -join ""
    $stream.Close(); $md5.Dispose()
    $outDir = Join-Path $dumpBase "$($f.DumpLabel)-$($hash.Substring(0,16))"

    if ((Test-Path "$outDir\ESS_Installs.json") -and (Test-Path "$outDir\SNs_Config_SVL.json") -and (Test-Path "$outDir\manifest.json")) {
        Write-Host "  Already complete, skipping." -ForegroundColor DarkGray; continue
    }

    Write-Host "  Source : $(Split-Path $xlsxPath -Leaf)" -ForegroundColor White
    Write-Host "  OutDir : $(Split-Path $outDir -Leaf)" -ForegroundColor White

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $counts = [XlsxDumper2]::DumpToJson($xlsxPath, $outDir, "ESS_Installs.json", "SNs_Config_SVL.json")
    $sw.Stop()

    $essRows = $counts[0]; $svlRows = $counts[1]
    Write-Host "  ESS Installs   : $essRows rows" -ForegroundColor Green
    Write-Host "  SNs Config SVL : $svlRows rows" -ForegroundColor Green
    Write-Host "  Time           : $($sw.Elapsed.TotalSeconds.ToString('F1'))s" -ForegroundColor White

    $manifestObj = [ordered]@{
        sheets = @(
            [ordered]@{ name="SNs Config &  SVL"; file="SNs_Config_SVL.json"; rowCount=$svlRows },
            [ordered]@{ name="ESS Installs";       file="ESS_Installs.json";  rowCount=$essRows }
        )
    }
    [System.IO.File]::WriteAllText(
        "$outDir\manifest.json",
        ($manifestObj | ConvertTo-Json -Depth 4),
        (New-Object System.Text.UTF8Encoding($false))
    )
    Write-Host "  DONE" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "All dumps complete." -ForegroundColor Green
Write-Host ""
Write-Host "Dump folders:" -ForegroundColor White
Get-ChildItem $dumpBase -Directory | Where-Object { $_.Name -match "3rd Nov" } | ForEach-Object {
    $fs = (Get-ChildItem $_.FullName | Select-Object -ExpandProperty Name) -join ", "
    Write-Host "  $($_.Name)  [$fs]" -ForegroundColor Green
}
