using System.IO;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System;

namespace PaperFormatDetection.Tools
{
    /**
     * 文件转换的工具类
     * 所有方法都是静态方法，可以直接使用"类名.方法名"方式调用
     */
    public class Converter
    {
        
        /* 将doc格式文件转为docx格式 */
        public static string convertDocToDocx(string docPath) {
            Console.WriteLine("0 => 正在转换文档格式");
            FileInfo file = new FileInfo(docPath);
            string docxPath = file.FullName;
            if (file.Extension.ToLower() == ".doc")
            {
                Application application = null;
                Document document = null;
                try
                {
                    application = new Application();
                    application.Visible = false;
                    application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                    object fileformat = WdSaveFormat.wdFormatXMLDocument;


                    string filename = file.FullName;
                    docxPath = file.FullName.ToLower().Replace(".doc", ".docx");
                    document = application.Documents.Open(filename);
                    document.SaveAs2(docxPath, WdSaveFormat.wdFormatXMLDocument,
                             CompatibilityMode: WdCompatibilityMode.wdWord2010);
                }
                finally {
                    document.Close();
                    document = null;
                    application.Quit();
                    application = null;
                }
                
            }
            Console.WriteLine("0 => 文档格式转换完成");
            return docxPath;
        }
        
        /* 将docx文件另存为新的docx */
        public static string saveAsDocx(string docxPath, string suffix) {
            FileInfo file = new FileInfo(docxPath);
            string newDocxPath = file.FullName;
            if (file.Extension.ToLower() == ".docx")
            {
                Application application = null;
                Document document = null;
                try
                {
                    application = new Application();
                    application.Visible = false;
                    application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                    object fileformat = WdSaveFormat.wdFormatXMLDocument;


                    string filename = file.FullName;
                    docxPath = file.FullName.ToLower().Replace(".doc", suffix + ".docx");
                    document = application.Documents.Open(filename);
                    document.SaveAs2(docxPath, WdSaveFormat.wdFormatXMLDocument,
                             CompatibilityMode: WdCompatibilityMode.wdWord2010);
                }
                finally
                {
                    document.Close();
                    document = null;
                    application.Quit();
                    application = null;
                }
            }
            return newDocxPath;
        }

        /* 
         * 使用Wordconv.exe将doc格式转为docx格式，暂不可用
         * 若使用，请将本机的Wordconv.exe绝对路径赋值给converter.StartInfo.FileName 
         */
        public static void convert(string fileIn, string fileOut) {
            Process converter = new Process();
            converter.StartInfo.FileName = "";//Wordconv绝对路径
            converter.StartInfo.Arguments = string.Format("-oice -nme \"{0}\" \"{1}\"", fileIn, fileOut);
            converter.StartInfo.CreateNoWindow = true;
            converter.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            converter.StartInfo.UseShellExecute = false;
            converter.StartInfo.RedirectStandardError = true;
            converter.StartInfo.RedirectStandardOutput = true;
            converter.Start();
            converter.WaitForExit();
            int exitCode = converter.ExitCode;
        }
    }
}
