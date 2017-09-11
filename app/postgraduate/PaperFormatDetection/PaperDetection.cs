using System.Xml;
using System.Collections.Generic;
using System;

namespace PaperFormatDetection.Frame
{
    /**
     * 论文检测类，整合论文检测的流程 
     */
    class PaperDetection
    {
        private List<Module> modList = new List<Module>();
        private int masterType;
        private bool usePageLocator;

        public PaperDetection(bool codeDetect, int masterType, bool usePageLocator)
        {
            this.init(codeDetect);
            this.masterType = masterType;
            this.usePageLocator = usePageLocator;
        }

        /* 初始化加载函数 */
        private void init(bool codeDetect)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load("modules.xml"); //加载xml文件
            //获取paper节点的所有子节点  
            XmlNodeList moduleList = xmlDoc.SelectSingleNode("paper").ChildNodes;

            //遍历所有子节点 
            foreach (XmlNode xmod in moduleList)
            {
                XmlElement mod = (XmlElement)xmod; //转换类型
                if (mod.GetAttribute("detect").Equals("true")) {
                    string className = mod.SelectSingleNode("className").InnerText;
                    string cnName = mod.SelectSingleNode("cnName").InnerText;
                    string rule = mod.SelectSingleNode("rule").InnerText;
                    bool detect = mod.GetAttribute("detect").Equals("true");
                    bool xqueryUse = mod.GetAttribute("xqueryUse").Equals("true");                
                    Module module = new Module(className, cnName, rule, detect, xqueryUse);
                    modList.Add(module);                        
                }
            }
            xmlDoc.Save("modules.xml");//保存
        }

        /* 论文检测，处理整个流程 */
        public void detect(string templatePath, string paperPath)
        {
            extractStyle(templatePath, paperPath);
            generateRules(templatePath, paperPath);
            compare(paperPath);
            generateReport(paperPath);
        }

        /* 论文格式提取 */
        public void extractStyle(string templatePath, string paperPath)
        {
            StyleExtractor se = new StyleExtractor(templatePath, paperPath, this.modList, this.masterType, this.usePageLocator);
            Console.WriteLine("1 => 论文格式提取");
            se.excute();
        }

        /* 生成各模块的对比规则文件 */
        private void generateRules(string templatePath, string paperPath)
        {
            Rules rl = new Rules(templatePath, paperPath, this.modList);
            Console.WriteLine("2 => 生成对比规则文件");
            rl.excute();
        }

        /* 格式文件对比 */
        private void compare(string paperPath) {
            RuleComparator rc = new RuleComparator(paperPath, this.modList);
            Console.WriteLine("3 => 对比格式文件");
            rc.excute();
        }

        /* 错误报告生成 */
        private void generateReport(string paperPath) {
            ReportGenerator rg = new ReportGenerator(paperPath, this.modList);
            Console.WriteLine("4 => 生成错误报告");
            rg.excute();
        }

        /* 打印modList,用于测试 */
        public void printModList() {
            Console.WriteLine("Module List = [ ");
            foreach (Module mod in modList){
                Console.WriteLine(mod.ToString());
            }
            Console.WriteLine(" ]");
        }
    }
}
