using NPOI.SS.UserModel;
using System;

namespace Natasha.Excel
{
    public class ColOperator
    {
        private readonly IRow _row;
        private ICell _cell;
        private int _column;
        public ColOperator(IRow row) { 
            _row = row;
            _cell = _row.GetCell(_column);
            if (_cell == null)
            {
                _cell = _row.CreateCell(_column);
            }
        }

        public ColOperator this[int colIndex]
        {
            get
            {
                _column = colIndex;
                _cell = _row.GetCell(_column);
                if (_cell == null)
                {
                    _cell = _row.CreateCell(_column);
                }
                return this;
            }
        }

        public ColOperator Next
        {
            get
            {
                _column += 1;
                return this[_column];
            }
        }

        public ColOperator Pre
        {
            get
            {
                _column -= 1;
                return this[_column];
            }
        }

        #region 获取/设置数据
        public void SetValue(bool value)
        {
            _cell.SetCellValue(value);
        }
        public void SetValue(long value)
        {
            _cell.SetCellValue(value);
        }
        public void SetValue(int value)
        {
            _cell.SetCellValue(value);
        }
        public void SetValue(short value)
        {
            _cell.SetCellValue(value);
        }
        public void SetValue(byte value)
        {
            _cell.SetCellValue(value);
        }
        public void SetValue(string value)
        {
            _cell.SetCellValue(value);
        }
        public void SetValue(double value)
        {
            _cell.SetCellValue(value);
        }
        public void SetValue(DateTime value)
        {
            _cell.SetCellValue(value);
        }
        public string StringValue
        {
            get
            {
                return _cell.StringCellValue;
            }
            set
            {
                _cell.SetCellValue(value);
            }

        }
        public bool BoolValue
        {
            get
            {
                return _cell.BooleanCellValue;
            }
            set
            {
                
                _cell.SetCellValue(value);
            }
        }
        public DateTime DateValue
        {
            get
            {
                return _cell.DateCellValue;
            }
            set
            {
                _cell.SetCellValue(value);
            }
        }
        public double NumValue
        {
            get
            {
                return _cell.NumericCellValue;
            }
            set
            {
                _cell.SetCellValue(value);
            }
        }
        public IRichTextString RichValue
        {
            get
            {
                return _cell.RichStringCellValue;
            }
            set
            {
                _cell.SetCellValue(value);
            }
        }
        #endregion
    }
}
