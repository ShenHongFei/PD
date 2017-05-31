using System;
using System.IO;
using PaperFormatDetection.Tools;
using System.Text.RegularExpressions;




namespace PaperFormatDetection.Frame
{
    /**
     * 程序调用的入口
     * Main()方法负责处理程序调用的参数 
     */
    public class Program
    {
        public static void Main(string[] args)
        {
            DateTime start = DateTime.Now;
            if (args.Length == 5)
            {
                string templatePath = args[0];//模板文件路径
                string paperPath = args[1];//待测文件路径
                string codeStr = args[2]; //检测代码参数
                string masterStr = args[3];//硕士类型参数 
                string usePageStr = args[4];//是否使用页码定位

                bool codeDetect = false;//默认是不检测代码
                int masterType = 1; //默认是专业硕士
                bool usePageLocator = false; //默认是不使用页码定位
                if (codeStr.Equals("true"))
                {
                    codeDetect = true;
                }
                if (masterStr.Equals("0"))
                {
                    masterType = 0;//学术硕士
                }
                if (usePageStr.Equals("true"))
                {
                    usePageLocator = true; //使用页码定位
                }
                if (File.Exists(templatePath) && File.Exists(paperPath))
                {
                    FileInfo paper = new FileInfo(paperPath);
                    if (paper.Length <= 0)
                    {
                        Console.WriteLine("文档不能为空！");
                    }
                    else
                    {
                        if (paper.Extension.Equals(".doc"))
                        {
                            paperPath = Converter.convertDocToDocx(paperPath);
                            PaperDetection pd = new PaperDetection(codeDetect, masterType, usePageLocator);
                            pd.detect(templatePath, paperPath);
                        }
                        else if (paper.Extension.Equals(".docx"))
                        {
                            PaperDetection pd = new PaperDetection(codeDetect, masterType, usePageLocator);
                            pd.detect(templatePath, paperPath);
                        }
                        else
                        {
                            Console.WriteLine("文档格式不对，请传入doc或者docx格式的文件！");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("路径输入错误或文件不存在！");
                }
            }
            else
            {
                Console.WriteLine("命令行参数不完整！");
                printHint();
                string templatePath = @"H:\本科\PaperFormatDetection\PaperFormatDetection\bin\Debug\temp.docx";
                string paperPath = @"H:\本科\PaperFormatDetection\PaperFormatDetection\bin\Debug\singletest\我的论文.docx";
                creatcopy(paperPath);
                paperPath = creatcopy(paperPath);
                string codeStr = "false"; //检测代码参数
                string masterStr = "1";//硕士类型参数
                string usePageStr = "false";//是否使用页码定位
                bool codeDetect = false;//默认是不检测代码
                int masterType = 1; //默认是工程硕士
                bool usePageLocator = false; //默认是不使用页码定位
                if (codeStr.Equals("true"))
                {
                    codeDetect = true;
                }
                if (masterStr.Equals("0"))
                {
                    masterType = 0;//学术硕士
                }
                if (usePageStr.Equals("true"))
                {
                    usePageLocator = true; //使用页码定位
                }
                if (File.Exists(templatePath) && File.Exists(paperPath))
                {
                    FileInfo paper = new FileInfo(paperPath);
                    if (paper.Length <= 0)
                    {
                        Console.WriteLine("文档不能为空！");
                    }
                    else
                    {
                        if (paper.Extension.Equals(".doc"))
                        {
                            paperPath = Converter.convertDocToDocx(paperPath);
                            PaperDetection pd = new PaperDetection(codeDetect, masterType, usePageLocator);
                            pd.detect(templatePath, paperPath);
                        }
                        else if (paper.Extension.Equals(".docx"))
                        {
                            PaperDetection pd = new PaperDetection(codeDetect, masterType, usePageLocator);
                            pd.detect(templatePath, paperPath);
                        }
                        else
                        {
                            Console.WriteLine("文档格式不对，请传入doc或者docx格式的文件！");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("路径输入错误或文件不存在！");
                }
            }
            DateTime end = DateTime.Now;
            Console.WriteLine(" <= 检测用时： " + DateDiff(start, end) + " =>");
            Console.ReadLine();
        }
        //生成副本
        public static string creatcopy(string paperPath)
        {
            FileInfo file = new FileInfo(paperPath);
            if (file.Exists)
            {
                //为副本文件加"副本"2字
                try
                {
                    Match match = Regex.Match(paperPath, @"(\w+)\.docx");
                    paperPath = paperPath.Substring(0, match.Index) + match.Groups[1] + "副本.docx";
                    // true is overwrite
                    file.CopyTo(paperPath, true);
                }
                catch
                {
                    Console.WriteLine("文件格式错误！");
                }
            }
            
            return paperPath;
        }

        

        /* 时间差计算 */
        private static string DateDiff(DateTime start, DateTime end)
        {
            string dateDiff = null;
            TimeSpan ts = start.Subtract(end).Duration();
            dateDiff = ts.Hours.ToString() + "小时" + ts.Minutes.ToString() + "分钟" + ts.Seconds.ToString() + "秒" + ts.Milliseconds.ToString() + "毫秒";
            return dateDiff;
        }

        /* 打印调用参数解释 */
        private static void printHint()
        {
            Console.WriteLine("程序调用需要参数列表如下：");
            Console.WriteLine("    参数1：模板文件路径");
            Console.WriteLine("    参数2：待测文件路径");
            Console.WriteLine("    参数3：检测代码参数 true-检测 false-不检测  默认不检测(false)");
            Console.WriteLine("    参数4：硕士类型参数 1-工程硕士 0-学术硕士  默认工程硕士(1)");
            Console.WriteLine("    参数5：是否使用页码定位 true-使用 false-不使用  默认不使用(false)");
            Console.WriteLine("======================");
        }
    }
}
