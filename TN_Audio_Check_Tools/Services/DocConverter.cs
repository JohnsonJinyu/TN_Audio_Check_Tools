using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace TN_Audio_Check_Tools.Services
{
    /// <summary>
    /// 提供将旧版 .doc 转换为 .docx 的工具方法。
    /// 优先尝试使用本机已安装的 Microsoft Word COM（需 STA），回退尝试使用 LibreOffice 的 soffice 命令行。
    /// 若转换成功，返回生成的 .docx 路径；失败返回 null。
    /// </summary>
    public static class DocConverter
    {
        public static string? ConvertDocToDocx(string docPath, int timeoutSeconds = 30)
        {
            return ConvertDocToDocxWithDetails(docPath, timeoutSeconds).Path;
        }

        /// <summary>
        /// 转换并返回详细信息：生成路径、使用的方法（Interop/LibreOffice/None）以及错误信息（如果有）。
        /// </summary>
        public static (string? Path, string Method, string? Error) ConvertDocToDocxWithDetails(string docPath, int timeoutSeconds = 30)
        {
            if (string.IsNullOrWhiteSpace(docPath)) return (null, "None", "Invalid path");
            var ext = Path.GetExtension(docPath);
            if (ext != null && ext.Equals(".docx", StringComparison.OrdinalIgnoreCase)) return (docPath, "AlreadyDocx", null);
            if (!ext.Equals(".doc", StringComparison.OrdinalIgnoreCase)) return (null, "None", "Not a .doc file");

            var dest = Path.Combine(Path.GetDirectoryName(docPath) ?? Path.GetTempPath(), Path.GetFileNameWithoutExtension(docPath) + ".docx");

            // 1) Interop (STA)
            var mre = new ManualResetEventSlim(false);
            string? interopResult = null;
            string? interopError = null;

            var t = new Thread(() =>
            {
                try
                {
                    Type? wordType = Type.GetTypeFromProgID("Word.Application");
                    if (wordType == null)
                    {
                        interopResult = null;
                        interopError = "Word COM not available";
                        return;
                    }

                    object? wordApp = null;
                    object? doc = null;
                    try
                    {
                        wordApp = Activator.CreateInstance(wordType);
                        if (wordApp == null) { interopError = "Failed to create Word instance"; return; }
                        dynamic w = wordApp;
                        try { w.Visible = false; } catch { }

                        // Ensure absolute path for SaveAs2
                        var destAbs = Path.GetFullPath(dest);

                        // CRITICAL: Open document with ReadOnly=false to allow SaveAs to work
                        doc = w.Documents.Open(docPath, false, false, false);

                        // Delete target file if it exists to avoid conflicts
                        try { if (File.Exists(destAbs)) File.Delete(destAbs); } catch { }

                        // SaveAs2: Save document in .docx format (FileFormat 12 = wdFormatXMLDocument)
                        w.ActiveDocument.SaveAs2(destAbs, 12, Type.Missing, Type.Missing, true, Type.Missing, false);

                        // Close the document
                        try { w.ActiveDocument.Close(false); } catch { }

                        // Quit Word application - this is critical!
                        try { w.Quit(false); } catch { }

                        // Release COM objects immediately
                        try { if (doc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(doc); } catch { }
                        doc = null;
                        try { if (wordApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); } catch { }
                        wordApp = null;

                        // Wait for Word process to completely exit and file to be flushed
                        Thread.Sleep(1500);

                        // Verify file was actually created by attempting to read file info
                        if (File.Exists(destAbs))
                        {
                            try
                            {
                                var fi = new FileInfo(destAbs);
                                // Try to get file length to ensure it's really accessible
                                var size = fi.Length;
                                if (size > 0)
                                {
                                    interopResult = destAbs;
                                }
                                else
                                {
                                    interopError = $"SaveAs2 created empty file: {destAbs}";
                                }
                            }
                            catch (Exception fileEx)
                            {
                                interopError = $"File exists but cannot be accessed: {fileEx.Message}";
                            }
                        }
                        else
                        {
                            // Fallback: search for .docx with same base name in target directory
                            try
                            {
                                var dir = Path.GetDirectoryName(destAbs) ?? Path.GetTempPath();
                                var baseName = Path.GetFileNameWithoutExtension(destAbs) ?? string.Empty;
                                if (Directory.Exists(dir))
                                {
                                    var matches = Directory.GetFiles(dir, "*.docx");
                                    foreach (var m in matches)
                                    {
                                        if (string.Equals(Path.GetFileNameWithoutExtension(m), baseName, StringComparison.OrdinalIgnoreCase))
                                        {
                                            try
                                            {
                                                var fi = new FileInfo(m);
                                                if (fi.Length > 0)
                                                {
                                                    interopResult = m;
                                                    break;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                                if (string.IsNullOrEmpty(interopResult))
                                {
                                    interopError = $"Expected file not found: {destAbs}";
                                }
                            }
                            catch (Exception searchEx)
                            {
                                interopError = $"File search failed: {searchEx.Message}";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        interopError = ex.Message;
                    }
                    finally
                    {
                        // Cleanup - ensure COM objects are released
                        try { if (doc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(doc); } catch { }
                        try { if (wordApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); } catch { }
                    }
                }
                catch (Exception ex)
                {
                    interopError = ex.Message;
                }
                finally
                {
                    mre.Set();
                }
            });

            t.SetApartmentState(ApartmentState.STA);
            t.IsBackground = true;
            t.Start();
            if (!mre.Wait(TimeSpan.FromSeconds(timeoutSeconds)))
            {
                try { t.Abort(); } catch { }
            }

            if (!string.IsNullOrEmpty(interopResult) && File.Exists(interopResult)) return (interopResult, "Interop", null);

            // 2) LibreOffice
                try
                {
                    var outDir = Path.GetDirectoryName(docPath) ?? Path.GetTempPath();
                    var psi = new ProcessStartInfo
                    {
                        FileName = "soffice",
                        Arguments = $"--headless --convert-to docx --outdir \"{outDir}\" \"{docPath}\"",
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };

                    using var p = Process.Start(psi);
                    string stdOut = string.Empty;
                    string stdErr = string.Empty;
                    if (p != null)
                    {
                        // read outputs (avoid deadlock by reading after ensuring start)
                        stdOut = p.StandardOutput.ReadToEnd();
                        stdErr = p.StandardError.ReadToEnd();
                        if (!p.WaitForExit(timeoutSeconds * 1000))
                        {
                            try { p.Kill(); } catch { }
                        }
                    }

                    // soffice normally writes converted file to outDir with same base name
                    if (File.Exists(dest)) return (dest, "LibreOffice", stdOut + (string.IsNullOrEmpty(stdErr) ? "" : "\nERR:\n" + stdErr));

                    // fallback: search for any .docx with same base name (case-insensitive)
                    try
                    {
                        var baseName = Path.GetFileNameWithoutExtension(docPath) ?? string.Empty;
                        var candidates = Directory.GetFiles(outDir, "*.docx");
                        foreach (var c in candidates)
                        {
                            if (string.Equals(Path.GetFileNameWithoutExtension(c), baseName, StringComparison.OrdinalIgnoreCase))
                            {
                                return (c, "LibreOffice", stdOut + (string.IsNullOrEmpty(stdErr) ? "" : "\nERR:\n" + stdErr));
                            }
                        }
                    }
                    catch { }

                    var msg = "soffice did not produce output" + (string.IsNullOrEmpty(stdErr) ? "" : ": " + stdErr);
                    return (null, "LibreOffice", msg);
                }
            catch (Exception ex)
            {
                var err = interopError ?? ex.Message;
                return (null, "None", err);
            }
        }
    }
}
