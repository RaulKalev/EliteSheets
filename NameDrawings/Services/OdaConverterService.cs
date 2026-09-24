using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace EliteSheets.Services
{
    /// <summary>
    /// Wraps the free ODA File Converter command line to turn DXF files into DWG.
    /// </summary>
    public class OdaConverterService
    {
        private const string ExeName = "ODAFileConverter.exe";
        private const string OutputVersion = "ACAD2018";
        private static readonly TimeSpan Timeout = TimeSpan.FromMinutes(10);

        /// <summary>
        /// Looks for ODAFileConverter.exe under Program Files\ODA\ODAFileConverter*\.
        /// Returns the newest install, or null when ODA is not installed.
        /// </summary>
        public static string FindConverter()
        {
            var roots = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
            };

            var candidates = new List<string>();
            foreach (var root in roots.Where(r => !string.IsNullOrEmpty(r)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var odaRoot = Path.Combine(root, "ODA");
                if (!Directory.Exists(odaRoot)) continue;

                try
                {
                    foreach (var dir in Directory.EnumerateDirectories(odaRoot, "ODAFileConverter*"))
                    {
                        var exe = Path.Combine(dir, ExeName);
                        if (File.Exists(exe)) candidates.Add(exe);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"Failed to scan {odaRoot} for ODA File Converter.", ex);
                }
            }

            return candidates
                .OrderByDescending(p => File.GetLastWriteTimeUtc(p))
                .FirstOrDefault();
        }

        /// <summary>
        /// Converts every *.dxf in <paramref name="inputFolder"/> to DWG (AutoCAD 2018) in <paramref name="outputFolder"/>.
        /// Returns false with an error message if the converter failed or timed out.
        /// </summary>
        public bool ConvertFolder(string converterPath, string inputFolder, string outputFolder, out string errorMessage)
        {
            errorMessage = null;
            try
            {
                Directory.CreateDirectory(outputFolder);

                // Args: <in folder> <out folder> <version> <type> <recurse> <audit> [filter]
                var psi = new ProcessStartInfo
                {
                    FileName = converterPath,
                    Arguments = $"\"{inputFolder.TrimEnd('\\')}\" \"{outputFolder.TrimEnd('\\')}\" {OutputVersion} DWG 0 1 \"*.DXF\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var proc = Process.Start(psi))
                {
                    if (proc == null)
                    {
                        errorMessage = "ODA File Converter could not be started.";
                        return false;
                    }

                    if (!proc.WaitForExit((int)Timeout.TotalMilliseconds))
                    {
                        try { proc.Kill(); } catch { }
                        errorMessage = "ODA File Converter timed out.";
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                Logger.Log("ODA File Converter failed.", ex);
                return false;
            }
        }
    }
}
