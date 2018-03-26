using AzulX.NPOI;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace System
{
    public class ExcelFile : IDisposable
    {
        private string _path;
        private ExcelVersion _version;
        internal IWorkbook _workbook;
        private ISheet _sheet;
        private IRow _row;
        private ICell _col;
        private Stream stream;

        private HashSet<string> _sheetNames;
        private HashSet<int> _sheetIndexs;

        public int CurrentCol;
        public int CurrentRow;

        public ExcelFile(string path, ExcelVersion version = ExcelVersion.V2007)
        {
            _path = path;
            _version = version;
            _sheetNames = new HashSet<string>();
            _sheetIndexs = new HashSet<int>();
            FileInit();
        }
        private void FileInit()
        {
            if (File.Exists(_path))
            {
                StreamReader reader = new StreamReader(_path, Encoding.UTF8);
                stream = reader.BaseStream;
                if (_version == ExcelVersion.V2007)
                {
                    _workbook = new XSSFWorkbook(stream);
                }
                else
                {
                    _workbook = new HSSFWorkbook(stream);
                }
            }
            else
            {
                FileStream writer = new FileStream(_path, FileMode.Create, FileAccess.Write);
                stream = writer;
                if (_version == ExcelVersion.V2007)
                {
                    _workbook = new XSSFWorkbook();
                }
                else
                {
                    _workbook = new HSSFWorkbook();
                }
            }
            RefreshSheets();
        }

        /// <summary>
        /// 刷新sheet的名与索引缓存
        /// </summary>
        private void RefreshSheets()
        {
            _sheetNames.Clear();
            _sheetIndexs.Clear();
            int sheets = _workbook.NumberOfSheets;
            for (int i = 0; i < sheets; i += 1)
            {
                _sheetNames.Add(_workbook.GetSheetName(i));
                _sheetIndexs.Add(i);
            }
        }
        /// <summary>
        /// 根据名字选择一个工作表，没有会自动创建
        /// </summary>
        /// <param name="sheetName">sheet的名字</param>
        /// <returns></returns>
        public ExcelFile Select(string sheetName)
        {
            if (!_sheetNames.Contains(sheetName))
            {
                _sheet = _workbook.CreateSheet(sheetName);
            }
            _sheet = _workbook.GetSheet(sheetName);
            RefreshSheets();
            MoveToRow(0);
            MoveToCol(0);
            return this;
        }
        /// <summary>
        /// 根据索引选择一个工作表，没有会自动创建
        /// </summary>
        /// <param name="index">sheet的索引</param>
        /// <returns></returns>
        public ExcelFile Select(int index)
        {
            index -= 1;
            if (index<0)
            {
                index = 0;
            }
            while (!_sheetIndexs.Contains(index))
            {
                _sheet = _workbook.CreateSheet();
                RefreshSheets();
            }
            _sheet = _workbook.GetSheetAt(index);
            MoveToRow(0);
            MoveToCol(0);
            return this;
        }


        #region 行列位置操作
        /// <summary>
        /// 移动到指定列
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        public ExcelFile MoveToCol(int column)
        {
            CurrentCol = column;
            _col = _row.GetCell(CurrentCol);
            if (_col == null)
            {
                _col = _row.CreateCell(CurrentCol);
            }
            return this;
        }
        /// <summary>
        /// 移动到指定行
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public ExcelFile MoveToRow(int row)
        {
            CurrentRow = row;
            _row = _sheet.GetRow(CurrentRow);
            if (_row == null)
            {
                _row = _sheet.CreateRow(CurrentRow);
            }
            return this;
        }
        #endregion

        #region 设置单元格

        /// <summary>
        /// 在当前列位置进行填充
        /// </summary>
        /// <param name="value">需要填充的值</param>
        /// <param name="style">单元格样式</param>
        /// <returns></returns>
        public ExcelFile Cell(string value, ICellStyle style = null)
        {
            _col.SetCellValue(value);
            if (style != null)
            {
                _col.CellStyle = style;
            }
            return this;
        }
        /// <summary>
        /// 在当前列位置进行填充
        /// </summary>
        /// <param name="value">需要填充的值</param>
        /// <param name="style">单元格样式</param>
        /// <returns></returns>
        public ExcelFile Cell(bool value, ICellStyle style = null)
        {
            _col.SetCellValue(value);
            if (style != null)
            {
                _col.CellStyle = style;
            }
            return this;
        }
        /// <summary>
        /// 在当前列位置进行填充
        /// </summary>
        /// <param name="value">需要填充的值</param>
        /// <param name="style">单元格样式</param>
        /// <returns></returns>
        public ExcelFile Cell(double value, ICellStyle style = null)
        {
            _col.SetCellValue(value);
            if (style != null)
            {
                _col.CellStyle = style;
            }
            return this;
        }
        /// <summary>
        /// 在当前列位置进行填充
        /// </summary>
        /// <param name="value">需要填充的值</param>
        /// <param name="style">单元格样式</param>
        /// <returns></returns>
        public ExcelFile Cell(IRichTextString value, ICellStyle style = null)
        {
            _col.SetCellValue(value);
            if (style != null)
            {
                _col.CellStyle = style;
            }
            return this;
        }
        /// <summary>
        /// 在当前位置，向下一列进行填充
        /// </summary>
        /// <param name="value">需要填充的值</param>
        /// <param name="style">单元格样式</param>
        /// <returns></returns>
        public ExcelFile NextCell(string value, ICellStyle style = null)
        {
            CurrentCol += 1;
            MoveToCol(CurrentCol);
            _col.SetCellValue(value);
            if (style != null)
            {
                _col.CellStyle = style;
            }
            return this;
        }
        /// <summary>
        /// 在当前位置，向下一列进行填充
        /// </summary>
        /// <param name="value">需要填充的值</param>
        /// <param name="style">单元格样式</param>
        /// <returns></returns>
        public ExcelFile NextCell(double value, ICellStyle style = null)
        {
            CurrentCol += 1;
            MoveToCol(CurrentCol);
            _col.SetCellValue(value);
            if (style != null)
            {
                _col.CellStyle = style;
            }
            return this;
        }
        /// <summary>
        /// 在当前位置，向下一列进行填充
        /// </summary>
        /// <param name="value">需要填充的值</param>
        /// <param name="style">单元格样式</param>
        /// <returns></returns>
        public ExcelFile NextCell(IRichTextString value, ICellStyle style = null)
        {
            CurrentCol += 1;
            MoveToCol(CurrentCol);
            _col.SetCellValue(value);
            if (style != null)
            {
                _col.CellStyle = style;
            }
            return this;
        }
        /// <summary>
        /// 在当前位置，向下一列进行填充
        /// </summary>
        /// <param name="value">需要填充的值</param>
        /// <param name="style">单元格样式</param>
        /// <returns></returns>
        public ExcelFile NextCell(bool value, ICellStyle style = null)
        {
            CurrentCol += 1;
            MoveToCol(CurrentCol);
            _col.SetCellValue(value);
            if (style != null)
            {
                _col.CellStyle = style;
            }
            return this;
        }
        /// <summary>
        /// 指定单元格进行填充
        /// </summary>
        /// <param name="index">列索引</param>
        /// <param name="value">需要填充的值</param>
        /// <param name="style">单元格样式</param>
        /// <returns></returns>
        public ExcelFile SpecialCell(int index, bool value, ICellStyle style = null)
        {
            MoveToCol(index);
            _col.SetCellValue(value);
            if (style != null)
            {
                _col.CellStyle = style;
            }
            return this;
        }
        /// <summary>
        /// 指定单元格进行填充
        /// </summary>
        /// <param name="index">列索引</param>
        /// <param name="value">需要填充的值</param>
        /// <param name="style">单元格样式</param>
        /// <returns></returns>
        public ExcelFile SpecialCell(int index, IRichTextString value, ICellStyle style = null)
        {
            MoveToCol(index);
            _col.SetCellValue(value);
            if (style != null)
            {
                _col.CellStyle = style;
            }
            return this;
        }
        /// <summary>
        /// 指定单元格进行填充
        /// </summary>
        /// <param name="index">列索引</param>
        /// <param name="value">需要填充的值</param>
        /// <param name="style">单元格样式</param>
        /// <returns></returns>
        public ExcelFile SpecialCell(int index, double value, ICellStyle style = null)
        {
            MoveToCol(index);
            _col.SetCellValue(value);
            if (style != null)
            {
                _col.CellStyle = style;
            }
            return this;
        }
        /// <summary>
        /// 指定单元格进行填充
        /// </summary>
        /// <param name="index">列索引</param>
        /// <param name="value">需要填充的值</param>
        /// <param name="style">单元格样式</param>
        /// <returns></returns>
        public ExcelFile SpecialCell(int index, string value, ICellStyle style = null)
        {
            MoveToCol(index);
            _col.SetCellValue(value);
            if (style != null)
            {
                _col.CellStyle = style;
            }
            return this;
        }
        #endregion

        #region 设置行
        /// <summary>
        /// 移动到下一行
        /// </summary>
        /// <param name="isResetCol">是否将列重置为首个单元格</param>
        /// <returns></returns>
        public ExcelFile NextRow(bool isResetCol = true)
        {
            MoveToRow(CurrentRow + 1);
            if (isResetCol)
            {
                MoveToCol(0);
            }
            return this;
        }
        /// <summary>
        /// 移动到上一行
        /// </summary>
        /// <param name="isResetCol">是否将列重置为首个单元格</param>
        /// <returns></returns>
        public ExcelFile PrewRow(bool isResetCol = true)
        {
            MoveToRow(CurrentRow - 1);
            if (isResetCol)
            {
                MoveToCol(0);
            }
            return this;
        }
        #endregion

        #region Template , JObject映射
        private Template _template;
        private ExcelTemplate _excelTemplate;

        public ExcelFile LoadTemplate(string path)
        {
            _excelTemplate = new ExcelTemplate();
            _excelTemplate.TemplateFromFile(path);
            return this;
        }
        /// <summary>
        /// 使用模板
        /// </summary>
        /// <param name="templateKey">模板文件中的Key</param>
        /// <returns></returns>
        public ExcelFile UseTemplate(string templateKey)
        {
            _template = _excelTemplate[templateKey];
            if (_template == null)
            {
                throw new NullReferenceException("未找到该模板模块，请核对key值~");
            }
            if (_template.Sheet != null)
            {
                Select(_template.Sheet);
            }
            return this;
        }
        /// <summary>
        /// 绘制模板头部
        /// </summary>
        /// <param name="styles">单元格样式</param>
        /// <returns>link</returns>
        public ExcelFile FillHeader(params ICellStyle[] styles)
        {
            int count = styles.Length;
            ICellStyle style = null;
            if (_template.Headers != null)
            {
                if (_template.HeaderStartAt != 0)
                {
                    MoveToCol(_template.HeaderStartAt);
                }
                for (int i = 0; i < _template.Headers.Count; i += 1)
                {
                    if (i < count)
                    {
                        style = styles[i];
                    }
                    NextCell(_template.Headers[i], style);
                }
                NextRow();
            }
            return this;
        }
        /// <summary>
        /// 绘制内容
        /// </summary>
        /// <typeparam name="T">实例类型</typeparam>
        /// <param name="collection">实例集合</param>
        /// <param name="styles">单元格样式</param>
        /// <returns></returns>
        public ExcelFile FillCollection<T>(IEnumerable<T> collection, params ICellStyle[] styles)
        {
            foreach (var item in collection)
            {
                Fill(item);
            }
            return this;
        }
        /// <summary>
        /// 绘制单条内容
        /// </summary>
        /// <typeparam name="T">实例类型</typeparam>
        /// <param name="t">实例</param>
        /// <param name="styles">单元格样式</param>
        /// <returns></returns>
        public ExcelFile Fill<T>(T t, params ICellStyle[] styles)
        {
            if (_template.ContentStartAt != 0)
            {
                MoveToCol(_template.ContentStartAt);
            }
            int count = styles.Length;
            ICellStyle style = null;
            JObject tempObject = JObject.FromObject(t);
            if (_template.Contents != null)
            {
                for (int i = 0; i < _template.Contents.Count; i += 1)
                {
                    string key = _template.Contents[i];
                    if (tempObject.ContainsKey(key))
                    {
                        if (i < count)
                        {
                            style = styles[i];
                        }
                        NextCell(tempObject[key].ToString(), style);
                    }
                }
            }
            NextRow();
            return this;
        }
        #endregion

        #region 末尾操作
        /// <summary>
        /// 保存文件
        /// </summary>
        public void Save()
        {
            stream.Dispose();
            stream = new FileStream(_path, FileMode.Create, FileAccess.Write);
            _workbook?.Write(stream);
        }
        /// <summary>
        /// 销毁
        /// </summary>
        public void Dispose()
        {
            stream?.Dispose();
            _sheet = null;
            _row = null;
            if (_workbook != null)
            {
                _workbook.Close();
                _workbook = null;
            }
        }
        #endregion
    }

    public enum ExcelVersion
    {
        V2003,
        V2007
    }
}
