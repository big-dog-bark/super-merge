using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace SuperMerge
{
    public static class MergeEngine
    {
        public static void MergeAndExport(string templateName, string outputDir, bool exportAsPdf = false)
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document mainDoc = app.ActiveDocument;
            Word.MailMerge merge = mainDoc.MailMerge;

            if (merge.MainDocumentType != Word.WdMailMergeMainDocType.wdFormLetters)
                throw new InvalidOperationException("当前文档不是邮件合并主文档。");

            int recordCount = merge.DataSource.RecordCount;

            for (int i = 1; i <= recordCount; i++)
            {
                // 👇 正确的方式：只合并当前一条
                merge.DataSource.ActiveRecord = (Word.WdMailMergeActiveRecord)i;
                merge.DataSource.FirstRecord = i;
                merge.DataSource.LastRecord = i;//Make it work properly by lzy
                string outputName = GenerateOutputName(templateName, merge);
                if (string.IsNullOrWhiteSpace(outputName))
                    outputName = $"Record_{i}";

                // 获取已有文档引用
                var beforeDocs = app.Documents
                    .Cast<Word.Document>()
                    .Select(d => d.Name) // 注意不要用 FullName，未保存文档可能为空
                    .ToHashSet();

                merge.Destination = Word.WdMailMergeDestination.wdSendToNewDocument;
                merge.Execute(false); // 👈 合并当前记录

                // 找出新文档
                Word.Document newDoc = app.Documents
                    .Cast<Word.Document>()
                    .FirstOrDefault(d => !beforeDocs.Contains(d.Name));

                if (newDoc == null)
                    throw new Exception("未能捕获生成的新文档。");

                string outputPath = Path.Combine(outputDir, outputName + (exportAsPdf ? ".pdf" : ".docx"));

                try
                {
                    if (exportAsPdf)
                    {
                        newDoc.ExportAsFixedFormat(outputPath, Word.WdExportFormat.wdExportFormatPDF);
                    }
                    else
                    {
                        newDoc.SaveAs2(outputPath);
                    }

                    System.Threading.Thread.Sleep(150);
                }
                finally
                {
                    try { newDoc.Close(false); } catch { }
                    Marshal.ReleaseComObject(newDoc);
                }
            }

            mainDoc.Activate(); // 激活回来
        }

        private static string GenerateOutputName(string template, Word.MailMerge merge)
        {
            string result = template;
            var matches = Regex.Matches(template, @"<([^<>]+)>");

            var fieldNames = matches.Cast<Match>()
                                    .Select(m => m.Groups[1].Value.Trim())
                                    .Distinct();

            foreach (var field in fieldNames)
            {
                try
                {
                    string value = merge.DataSource.DataFields[field]?.Value ?? "未知";
                    result = result.Replace($"<{field}>", SanitizeFileName(value));
                }
                catch
                {
                    result = result.Replace($"<{field}>", "缺失字段");
                }
            }

            return result.Trim();
        }

        private static string SanitizeFileName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name.Trim();
        }
    }
}
