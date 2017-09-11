using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using PaperFormatDetection.Format;
using PaperFormatDetection.Tools;

namespace PaperFormatDetection.Frame
{
    public class StyleExtractor : ModuleProcessor
    {
        private PageLocator locator;
        private int masterType;

        /* 构造函数 */
        public StyleExtractor(string templatePath, string paperPath, List<Module> modList, int masterType, bool usePageLocator) : base(templatePath, paperPath, modList)
        {

            locator = new PageLocator(paperPath, usePageLocator);
            this.masterType = masterType;
        }

        /* 执行函数，继承自ModuleProcessor */
        public override void excute()
        {
            extractStyle();
        }

        private void extractStyle() {
            getTemplateStyle(templatePath, modList);
            getPaperStyle(templatePath, paperPath, modList);
        }

        /* 模板格式提取 */
        private void getTemplateStyle(string templatePath, List<Module> modList)
        {
            
            string templateName = Path.GetFileNameWithoutExtension(templatePath);// 没有扩展名的文件名
            string templateFolder = "Templates\\" + templateName;
            try
            {
                if (!Directory.Exists(templateFolder))
                {
                    Console.WriteLine("论文模板格式提取开始！");
                    Directory.CreateDirectory(templateFolder);
                    WordprocessingDocument wd = WordprocessingDocument.Open(templatePath, true);
                    getModulesStyle(wd, templateFolder, modList);
                    wd.Close();
                }
                else {
                    Console.WriteLine("论文模板已存在，不再提取格式！");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("论文模板文件夹创建失败: ", e.ToString());
            }
        }

        /* 待测文件格式提取 */
        private void getPaperStyle(string templatePath, string paperPath, List<Module> modList)
        {
            Console.WriteLine("待测论文格式提取开始！");
            string paperName = Path.GetFileNameWithoutExtension(paperPath);// 没有扩展名的文件名
            string paperFolder = "Papers\\" + paperName;
            try
            {
                if (!Directory.Exists(paperFolder))
                {
                    Directory.CreateDirectory(paperFolder);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("待测论文文件夹创建失败: ", e.ToString());
            }
            WordprocessingDocument wd = WordprocessingDocument.Open(paperPath, true);
            getModulesStyle(wd, paperFolder, modList);
            wd.Close();
        }

        /* 使用各模块的检测方法提取格式 */
        private void getModulesStyle(WordprocessingDocument wd, string folderPath, List<Module> modList)
        {
            //遍历所有子节点 
            foreach (Module mod in modList)
            {
                if (mod.Detect == true)
                {
                    string className = mod.ClassName;
                    string path = "PaperFormatDetection.Format." + className;
                    //反射创建类的实例，返回为 object 类型
                    Console.WriteLine("  ==> [" + mod.CnName + "]正在提取格式");

                    Type type = Type.GetType(path);
                    object[] parameters = new object[3];
                    parameters[0] = this.modList;
                    parameters[1] = this.locator;
                    parameters[2] = this.masterType;
                    object obj = Activator.CreateInstance(type, parameters);
                    ModuleFormat imod = obj as ModuleFormat;
                    imod.getStyle(wd, folderPath);
                }
            }
        }
    }
}
