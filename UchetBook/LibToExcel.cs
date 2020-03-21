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
    }
}
