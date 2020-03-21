using System;

using Excel = Microsoft.Office.Interop.Excel;

namespace Vng.Common
{
    class LibToExcel
    {
        public void CellMerge(string title, string cell1, string cell2, double tWidth,
                            bool wrpText, double tFont, char tHor, char tVer,
                            int tOrient, Excel.Worksheet xlSh)
        {
            Excel.Range xlSheetRange;               //Выделеная область

            // диапазон
            xlSheetRange = xlSh.get_Range(cell1, cell2);
            // Объединяем ячейки
            xlSheetRange.Merge(Type.Missing);
            xlSheetRange.Value2 = title;
            xlSheetRange.Orientation = tOrient;

            if (wrpText)
            { xlSheetRange.WrapText = wrpText; }

            if (tWidth > 0)
            { xlSheetRange.ColumnWidth = tWidth; }

            if (tFont > 0)
            { xlSheetRange.Font.Size = tFont; }

            switch (tHor)
            {
                case 'C':                       //Задаем выравнивание по центру
                    xlSheetRange.HorizontalAlignment = Excel.Constants.xlCenter;
                    break;
                case 'L':
                    xlSheetRange.HorizontalAlignment = Excel.Constants.xlLeft;
                    break;
                case 'R':
                    xlSheetRange.HorizontalAlignment = Excel.Constants.xlRight;
                    break;
                default:
                    break;
            }

            switch (tVer)
            {
                case 'C':
                    //Задаем выравнивание по центру
                    xlSheetRange.VerticalAlignment = Excel.Constants.xlCenter;
                    break;
                case 'T':
                    xlSheetRange.VerticalAlignment = Excel.Constants.xlTop;
                    break;
                case 'B':
                    xlSheetRange.VerticalAlignment = Excel.Constants.xlBottom;
                    break;
                default:
                    break;
            }
        }

        public void ColumnFormat(int column, int topRow, int bottomRow, bool wrpText,
                             double tFont, char tHor, string tFormat, Excel.Worksheet xlSh)
        {
            Excel.Range c1 = (Excel.Range)xlSh.Cells[topRow, column];              //"B10"
            Excel.Range c2 = (Excel.Range)xlSh.Cells[bottomRow, column];
            Excel.Range range = xlSh.get_Range(c1, c2);

            if (wrpText)
            { range.WrapText = wrpText; }

            switch (tHor)
            {
                case 'C':                       //Задаем выравнивание по центру
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    break;
                case 'L':
                    range.HorizontalAlignment = Excel.Constants.xlLeft;
                    break;
                case 'R':
                    range.HorizontalAlignment = Excel.Constants.xlRight;
                    break;
                default:
                    break;
            }

            if (tFont > 0)
            { range.Font.Size = tFont; }

            //range.Font.Name = "Arial";
            //range.NumberFormat = "@";

            range.NumberFormat = tFormat;
        }
    }
}
