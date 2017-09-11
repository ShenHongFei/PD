using System;
using Saxon.Api;
using System.Xml;
using System.IO;
using System.Collections.Generic;
namespace PaperFormatDetection.Frame
{
    /**
     * 报告生成器类，继承自ModuleProcessor
     */
    public class ReportGenerator :ModuleProcessor
    {
        static int count = 1;
        int errModuelCount = 0;

        public ReportGenerator(string paperPath, List<Module> modList) : base(paperPath, modList)
        {

        }

        public override void excute()
        {
            creatReport();
        }

        /* 生成报告 */
        private void creatReport()
        {
            count = 1;
            try{
                string paperName = Path.GetFileNameWithoutExtension(paperPath);//待测论文文件夹名称
    
                string reportPath = "Papers\\" + paperName + "\\report.txt";
                FileStream fs = new FileStream(reportPath, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                //报告前面的文本
                sw.WriteLine("       红蚂蚁实验室提供技术支持");
                //sw.WriteLine("当前版本无法检测公式，代码，参考文献的格式");
                sw.WriteLine("");
                

                //遍历所有子节点  
                foreach (Module mod in modList)
                {
                    string className = mod.ClassName;//模块名称
                    //这个是partName，就是每个部分名称
                    //（比如：-----------------正文-----------------）
                    writePartNameInReport(paperName, className, sw);
                    if (mod.XqueryUse == true)
                    {
                        writeModuleReportXQuery(paperName, className,sw);
                    }
                    //这个是把spErroInfo这个标签下的都拿出来
                    writeSpErroInfoInReport(paperName, className, sw);
                }
                //如果每个部分spErroInfo下面都没有东西
                if (errModuelCount == 0)
                {
                    sw.WriteLine("恭喜您，在本系统检测范围内，你的论文无错误！");
                }
    
                //清空缓冲区
                sw.Flush();
                //关闭流
                sw.Close();
                fs.Close();
            } catch(IOException e){
                Console.WriteLine(e.ToString());
            }
        }

        private void writeModuleReportXQuery(string paperName, string className, StreamWriter sw) {
            Processor pr = new Processor();
            XQueryCompiler xqc = pr.NewXQueryCompiler();

            string errPath = "Papers\\" + paperName + "\\" + className + "_err.xml";
            XmlDocument errdoc = new XmlDocument();
            errdoc.Load(errPath);
            XmlElement root = errdoc.DocumentElement;
            XmlNodeList ruleNodes = root.GetElementsByTagName("err-rule");
            foreach (XmlNode xrule in ruleNodes)
            {
                string rule = xrule.InnerText;
                //Console.WriteLine(rule);

                XQueryExecutable xqe = xqc.Compile(rule);
                XQueryEvaluator xqev = xqe.Load();
                XdmValue result = xqev.Evaluate(); //对比结果


                foreach (XdmItem rs in result)
                {
                    sw.WriteLine(count + ". " + rs.ToString().Trim());
                    count++;
                }
            }
        }

        private void writeModuleReport(string paperName, string className, StreamWriter sw)
        {

            string configPath = "Config\\" + className + ".config";
            XmlDocument confdoc = new XmlDocument();
            confdoc.Load(configPath);
            XmlNode xconf = confdoc.SelectSingleNode("conf");

            string errPath = "Papers\\" + paperName + "\\" + className + "_err.xml";
            XmlDocument errdoc = new XmlDocument();
            errdoc.Load(errPath);
            XmlElement root = errdoc.DocumentElement;
            XmlNodeList ruleNodes = root.GetElementsByTagName("err-rule");

            foreach (XmlNode xrule in ruleNodes)
            {
                string rule = xrule.InnerText;
                string errmsg = count + ". " + getErrorMsg(xconf, rule);
                sw.WriteLine(errmsg.Trim());
                count++;
            }
        }

        private string getErrorMsg(XmlNode xconf, string errstr)
        {
            string errmsg = "";
            string[] errarr = errstr.Split('/');
            if (errarr.Length == 3){
                XmlNode xmod = xconf.SelectSingleNode("module");
                XmlElement mod = (XmlElement)xmod;
                if (mod.GetAttribute("tag").Equals(errarr[0])) {
                    XmlNodeList seclist = xmod.ChildNodes;
                    foreach (XmlNode xsec in seclist) {
                        XmlElement sec = (XmlElement)xsec;
                        if (sec.GetAttribute("tag").Equals(errarr[1])) {
                            XmlNodeList itemlist = xsec.ChildNodes;
                            foreach (XmlNode xitem in itemlist) {
                                XmlElement item = (XmlElement)xitem;
                                if (item.GetAttribute("tag").Equals(errarr[2])) {
                                    errmsg = item.FirstChild.InnerText;
                                }
                            }
                        }
                    }
                }
            }
            else {
                XmlNode xmod = xconf.SelectSingleNode("module");
                XmlElement mod = (XmlElement)xmod;
                if (mod.GetAttribute("tag").Equals(errarr[0]))
                {
                    XmlNodeList seclist = xmod.ChildNodes;
                    
                    foreach (XmlNode xsec in seclist)
                    {
                        XmlElement sec = (XmlElement)xsec;
                        if (sec.GetAttribute("tag").Equals(errarr[1]))
                        {
                            XmlNodeList partlist = xsec.ChildNodes;
                            foreach (XmlNode xpart in partlist) {
                                XmlElement part = (XmlElement)xpart;
                                if (part.GetAttribute("tag").Equals(errarr[2])) {
                                    XmlNodeList itemlist = xpart.ChildNodes;

                                    foreach (XmlNode xitem in itemlist)
                                    {
                                        XmlElement item = (XmlElement)xitem;
                                        if (item.GetAttribute("tag").Equals(errarr[3]))
                                        {
                                            errmsg = item.FirstChild.InnerText;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return errmsg;
        }

        private void writeSpErroInfoInReport(string paperName, string className, StreamWriter sw)
        {
            String spInfoPoision = className + "/spErroInfo";
            string errPath = "Papers\\" + paperName + "\\" + className + ".xml";
            XmlDocument errdoc = new XmlDocument();
            errdoc.Load(errPath);
            XmlNode root = errdoc.SelectSingleNode(spInfoPoision);
            if(root != null)
            {
                foreach (XmlNode xrule in root.ChildNodes)
                {
                    string info = count+". "+xrule.InnerText;
                    sw.WriteLine(info.Trim());
                    count++;
                }
            }
            
        }

        private void writePartNameInReport(string paperName, string className, StreamWriter sw)
        {
            String spInfoPoision = className + "/partName";
            String flagPath= className + "/spErroInfo";
            string errPath = "Papers\\" + paperName + "\\" + className + ".xml";
            XmlDocument errdoc = new XmlDocument();
            errdoc.Load(errPath);
            XmlNode root = errdoc.SelectSingleNode(spInfoPoision);
            XmlNode spRoot = errdoc.SelectSingleNode(flagPath);
            if (root != null)
            {
                if(spRoot != null)
                {
                    if(spRoot.ChildNodes.Count > 0)
                    {
                        errModuelCount++;
                        foreach (XmlNode xrule in root.ChildNodes)
                        {
                            string info = xrule.InnerText;
                            sw.WriteLine(info.Trim());
                        }
                    }
                    
                }
                
            }

        }
    }
}
