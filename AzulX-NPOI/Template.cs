using System;
using System.Collections.Generic;

namespace AzulX.NPOI
{
    public class Template
    {
        public string Sheet;
        public int HeaderStartAt;
        public int ContentStartAt;
        public List<string> Headers;
        public List<string> Contents;

        /// <summary>
        /// 展示当前模板信息
        /// </summary>
        public void Show()
        {
            Console.WriteLine("Page:"+Sheet);
            string header = null;
            for (int i = 0; i < Headers.Count; i += 1)
            {
                header += Headers[i] + "、";
            }
            Console.WriteLine($"头部开始在 { HeaderStartAt } 信息为: {header}");
            string content = null;
            for (int i = 0; i < Contents.Count; i += 1)
            {
                content += Contents[i] + "、";
            }
            Console.WriteLine($"内容开始在 { ContentStartAt } 信息为: {content}");
        }
    }
}
