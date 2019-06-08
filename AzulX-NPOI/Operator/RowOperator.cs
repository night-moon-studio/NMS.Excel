using NPOI.SS.UserModel;

namespace AzulX.NPOI
{
    public class RowOperator
    {
        private readonly ISheet _sheet;
        private IRow _row;
        private int _rowIndex;
        public RowOperator(ISheet sheet) => _sheet = sheet;

        public ColOperator this[int rowIndex]
        {
            get
            {
                _rowIndex = rowIndex;
                _row = _sheet.GetRow(_rowIndex);
                if (_row == null)
                {
                    _row = _sheet.CreateRow(_rowIndex);
                }
                return new ColOperator(_row);
            }
        }

        public ColOperator NextRow
        {
            get
            {
                _rowIndex += 1;
                return this[_rowIndex];
            }
        }

        public ColOperator PreRow
        {
            get
            {
                _rowIndex -= 1;
                return this[_rowIndex];
            }
        }
    }
}
