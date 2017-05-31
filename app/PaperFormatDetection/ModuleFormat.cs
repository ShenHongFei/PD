using DocumentFormat.OpenXml.Packaging;
using PaperFormatDetection.Frame;
using PaperFormatDetection.Tools;
using System;
using System.Collections.Generic;

namespace PaperFormatDetection.Format
{
    public abstract class ModuleFormat
    {
        protected List<Module> modList; //检测模块列表
        protected PageLocator locator; //页定位器
        protected int masterType; //硕士类型
        protected int pageNum; //页码

        /* 构造函数 */
        public ModuleFormat(List<Module> modList, PageLocator locator, int masterType) {
            this.modList = modList;
            this.locator = locator;
            this.masterType = masterType;
            pageNum = 1;
        }
        
        /* 检测函数 */
        public abstract void getStyle(WordprocessingDocument doc, String fileName);

        /* 获取所在页码 */
        protected int getPageNum(string text)
        {
            return this.locator.findPageNum(text);
        }

        /* 获取所在页码 */
        protected int getPageNum(int pageNum, string text)
        {
            return this.locator.findPageNumStartWith(pageNum, text);
        }

        /* 包装页码信息 */
        protected string addPageInfo(int pageNum)
        {
            if (pageNum == -1)
            {
                return "";
            }
            return "【第" + pageNum + "页】";
        }
    }
}
