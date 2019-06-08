using NPOI.SS.UserModel;
using System;

namespace AzulX.NPOI
{
    public class ExcelStyle
    {
        private IWorkbook _workbook;

        public ExcelStyle(IWorkbook workbook)
        {
            _workbook = workbook;
        }
        public static ExcelStyle Create(ExcelOperator file)
        {
            return new ExcelStyle(file._workbook);
        }

        public ICellStyle Header()
        {
            ICellStyle style = _workbook.CreateCellStyle();
            IFont font = _workbook.CreateFont();
            font.IsBold = true;
            style.SetFont(font);
            return style;
        }
    }
}
