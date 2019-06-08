using NPOI.SS.UserModel;
using System;

namespace AzulX.NPOI
{
    public class ColOperator
    {
        private readonly IRow _row;
        private ICell _cell;
        private int _column;
        public ColOperator(IRow row) => _row = row;

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

        public ColOperator NextCol
        {
            get
            {
                _column += 1;
                return this[_column];
            }
        }

        public ColOperator PreCol
        {
            get
            {
                _column -= 1;
                return this[_column];
            }
        }

        #region 获取/设置数据

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
