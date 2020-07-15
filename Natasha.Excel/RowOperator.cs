using NPOI.SS.UserModel;

namespace Natasha.Excel
{
    public class RowOperator
    {
        private readonly ISheet _sheet;
        private IRow _row;
        private int _rowIndex;
        public RowOperator(ISheet sheet) {

            _sheet = sheet;
            _row = _sheet.GetRow(_rowIndex);
            if (_row == null)
            {
                _row = _sheet.CreateRow(_rowIndex);
            }
        } 

        public short Count { 

            get 
            { 
                return _row.LastCellNum; 
            } 

        }

        public RowOperator this[int rowIndex]
        {
            get
            {
                _rowIndex = rowIndex;
                _row = _sheet.GetRow(_rowIndex);
                if (_row == null)
                {
                    _row = _sheet.CreateRow(_rowIndex);
                }
                return this;
            }
        }

        public ColOperator Columns
        {
            get { return new ColOperator(_row); }
        }

        public RowOperator Next
        {
            get
            {
                _rowIndex += 1;
                return this[_rowIndex];
            }
        }

        public RowOperator Pre
        {
            get
            {
                _rowIndex -= 1;
                return this[_rowIndex];
            }
        }
    }
}
