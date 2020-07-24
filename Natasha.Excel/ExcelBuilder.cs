using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Natasha.Excel
{
    public class ExcelBuilder : IDisposable
    {
        internal IWorkbook _workbook;
        private ISheet _sheet;
        private Stream _stream;
        private readonly string _path;
        private readonly ExcelVersion _version;
        private readonly HashSet<string> _sheetNames;
        private readonly HashSet<int> _sheetIndexs;

        public int CurrentCol;
        public int CurrentRow;

        public ExcelBuilder(string path, ExcelVersion version = ExcelVersion.V2007)
        {
            _path = path;
            _version = version;
            _sheetNames = new HashSet<string>();
            _sheetIndexs = new HashSet<int>();
            if (File.Exists(_path))
            {
                StreamReader reader = new StreamReader(_path, Encoding.UTF8);
                _stream = reader.BaseStream;
                if (_version == ExcelVersion.V2007)
                {
                    _workbook = new XSSFWorkbook(_stream);
                }
                else
                {
                    _workbook = new HSSFWorkbook(_stream);
                }
            }
            else
            {
                FileStream writer = new FileStream(_path, FileMode.Create, FileAccess.Write);
                _stream = writer;
                if (_version == ExcelVersion.V2007)
                {
                    _workbook = new XSSFWorkbook();
                }
                else
                {
                    _workbook = new HSSFWorkbook();
                }
            }

            if (_workbook.NumberOfSheets != 0)
            {
                _sheet = _workbook.GetSheetAt(0);
            }
           
        }



        public int Count
        {
            get { return _sheet.LastRowNum; }
        }


        #region Sheet页操作
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
        /// 检测Sheet页是否存在
        /// </summary>
        /// <param name="index">Sheet页索引</param>
        /// <returns></returns>
        public bool HasSheet(int index)
        {
            return _sheetIndexs.Contains(index);
        }
        /// <summary>
        /// 检测Sheet页是否存在
        /// </summary>
        /// <param name="name">Sheet页名字</param>
        /// <returns></returns>
        public bool HashSheet(string name)
        {
            return _sheetNames.Contains(name);
        }
        #endregion

        #region 行列位置操作
        public ISheet this[int sheetIndex]
        {
            get
            {
                sheetIndex -= 1;

                if (sheetIndex < 0) { sheetIndex = 0; }

                while (!_sheetIndexs.Contains(sheetIndex))
                {
                    _sheet = _workbook.CreateSheet();
                    RefreshSheets();
                }
                return _sheet = _workbook.GetSheetAt(sheetIndex);
            }
        }

        public ISheet this[string sheetIndex]
        {
            get
            {
                if (!_sheetNames.Contains(sheetIndex))
                {
                    _sheet = _workbook.CreateSheet(sheetIndex);
                }
                _sheet = _workbook.GetSheet(sheetIndex);
                RefreshSheets();
                return _sheet;
            }
        }
        #endregion


        /// <summary>
        /// 保存文件
        /// </summary>
        public void Save()
        {
            _stream.Dispose();
            _stream = new FileStream(_path, FileMode.Create, FileAccess.Write);
            _workbook?.Write(_stream);
        }
        /// <summary>
        /// 销毁
        /// </summary>
        public void Dispose()
        {
            _stream?.Dispose();
            _sheet = null;
            if (_workbook != null)
            {
                _workbook.Close();
                _workbook = null;
            }
        }
    }

    public enum ExcelVersion
    {
        V2003,
        V2007
    }
}
