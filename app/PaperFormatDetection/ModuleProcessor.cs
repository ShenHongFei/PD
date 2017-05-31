using System.Collections.Generic;

namespace PaperFormatDetection.Frame
{
    /**
     * 模块处理流程的基类
     */
    public abstract class ModuleProcessor
    {
        protected string templatePath;
        protected string paperPath;
        protected List<Module> modList;

        /* 构造函数 */
        public ModuleProcessor(string paperPath, List<Module> modList) {
            this.templatePath = null;
            this.paperPath = paperPath;
            this.modList = modList;
        }

        /* 构造函数 */
        public ModuleProcessor(string templatePath, string paperPath, List<Module> modList) {
            this.templatePath = templatePath;
            this.paperPath = paperPath;
            this.modList = modList;
        }

        /* 执行函数，抽象方法 */
        public abstract void excute();
    }
}
