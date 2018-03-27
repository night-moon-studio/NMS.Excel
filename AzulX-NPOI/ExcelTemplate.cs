using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace AzulX.NPOI
{
    public class ExcelTemplate
    {
        private Regex _endReg;
        private Regex _nameReg;
        private Regex _sheetReg;
        private Regex _spiltReg;
        private string _split;
        private Regex _headerReg;
        private Regex _startAtReg;
        private Regex _contentReg;

        public HashSet<Template> Templates;
        public Dictionary<string, Template> TemplatesMapping;
        public ExcelTemplate()
        {
            Templates = new HashSet<Template>();
            TemplatesMapping = new Dictionary<string, Template>();
            _endReg = new Regex("@End");
            _nameReg = new Regex("(@Name[:|：](?<result>.*))", RegexOptions.Compiled);
            _sheetReg = new Regex("(@Sheet[:|：](?<result>.*))", RegexOptions.Compiled);
            _spiltReg = new Regex("(@Split[:|：](?<result>.*))", RegexOptions.Compiled);
            _headerReg = new Regex("(@Header[:|：](?<result>.*))", RegexOptions.Compiled);
            _startAtReg = new Regex("(@StartAt[:|：](?<result>\\d*?)\\D)", RegexOptions.Compiled);
            _contentReg = new Regex("(@Content[:|：](?<result>.*))", RegexOptions.Compiled);
        }

        /// <summary>
        /// 根据Key筛选模板
        /// </summary>
        /// <param name="key">对应配置文件中的@Name</param>
        /// <returns></returns>
        public Template this[string key]
        {
            get
            {
                if (TemplatesMapping.ContainsKey(key))
                {
                    return TemplatesMapping[key];
                }
                return null;
            }
            set
            {
                TemplatesMapping[key] = value;
            }
        }

        /// <summary>
        /// 从文件中读取配置模块
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns></returns>
        public ExcelTemplate TemplateFromFile(string path)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException("文件不存在，请检查文件在磁盘上的位置！");
            }
            string content = File.ReadAllText(path);
            return Template(content);
        }

        /// <summary>
        /// 根据内容解析模板
        /// </summary>
        /// <param name="content">字符串内容</param>
        /// <returns></returns>
        public ExcelTemplate Template(string content)
        {
            string[] contents = content.Split("\r\n");
            Template template = new Template();
            string Name = null;
            for (int i = 0; i < contents.Length; i += 1)
            {
                //匹配@Name
                Match match = _nameReg.Match(contents[i]);
                if (match.Success)
                {
                    Name = match.Groups["result"].Value.Trim();
                    continue;
                }
                //匹配@End
                match = _endReg.Match(contents[i]);
                if (match.Success)
                {
                    Templates.Add(template);
                    if (Name != null)
                    {
                        TemplatesMapping[Name] = template;
                    }
                    template = new Template();
                    Name = null;
                    continue;
                }

                //匹配@Spilt
                match = _spiltReg.Match(contents[i]);
                if (match.Success)
                {
                    _split = match.Groups["result"].Value;
                    continue;
                }

                //匹配@Sheet
                match = _sheetReg.Match(contents[i]);
                if (match.Success)
                {
                    template.Sheet = match.Groups["result"].Value.Trim();
                    continue;
                }


                //匹配@Header
                match = _headerReg.Match(contents[i]);
                if (match.Success)
                {
                    string temp = match.Groups["result"].Value.Trim();

                    //匹配@StartAt
                    match = _startAtReg.Match(temp);
                    if (match.Success)
                    {
                        string index = match.Groups["result"].Value.Trim();
                        template.HeaderStartAt = int.Parse(index) - 1;
                        temp = _startAtReg.Replace(temp, "");
                    }
                    //获取头信息
                    template.Headers = new List<string>(temp.Trim().Split(_split));
                    continue;
                }

                //匹配@Content
                match = _contentReg.Match(contents[i]);
                if (match.Success)
                {
                    string temp = match.Groups["result"].Value.Trim();

                    //匹配@StartAt
                    match = _startAtReg.Match(temp);
                    if (match.Success)
                    {
                        string index = match.Groups["result"].Value.Trim();
                        template.ContentStartAt = int.Parse(index) - 1;
                        temp = _startAtReg.Replace(temp, "");
                    }
                    //获取头信息
                    template.Contents = new List<string>(temp.Trim().Split(_split));
                    continue;
                }
            }
            return this;
        }

        /// <summary>
        /// 展示模板信息
        /// </summary>
        public void Show()
        {
            Console.WriteLine("分隔符为：" + _split);
            if (Templates != null)
            {
                foreach (var item in Templates)
                {
                    item.Show();
                }
            }
        }
    }
}
