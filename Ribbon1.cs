using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SuperMerge;

namespace SuperMerge
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnExportDOCX_Click(object sender, RibbonControlEventArgs e)
        {
            bool isPDF = false; // 设置为导出DOCX
            AskAndExport(isPDF);
        }

        private static bool AskAndExport(bool isPDF)
        {
            // 1. 弹inputbox输入文件名模板
            string templateName = Interaction.InputBox(
                "请输入导出文件名模板，支持用<字段>替换",
                "输入文件名模板",
                "<姓名>_合同");
            if (string.IsNullOrWhiteSpace(templateName))
            {
                MessageBox.Show("未输入文件名模板，取消导出");
                return false;
            }

            // 2. 弹文件夹选择对话框
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.Description = "请选择导出文件夹";
                if (fbd.ShowDialog() != DialogResult.OK)
                {
                    MessageBox.Show("未选择导出目录，取消导出");
                    return false;
                }

                string outputDir = fbd.SelectedPath;

                try
                {
                    MergeEngine.MergeAndExport(templateName, outputDir, isPDF);
                    MessageBox.Show("导出成功！");
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("导出失败：" + ex.Message);
                }
            }

            return true;
        }

        private void btnExportPDF_Click(object sender, RibbonControlEventArgs e)
        {
            bool isPDF = true; // 设置为导出PDF
            AskAndExport(isPDF);
        }
    }
}
