using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace npoiDemo
{
    public static class ExportExcelExtensions
    {
        public static ICell CreateCellWithCellStyle(this IRow row, int columnIndex, ICellStyle cellStyle, IComment cellComment = null)
        {
            var cell = row.CreateCell(columnIndex);
            cell.CellStyle = cellStyle;

            if (cellComment != null)
            {
                cell.CellComment = cellComment;
            }

            return cell;
        }

        

        public static IRow CreateRowWithHeightInPoints(this ISheet sheet, int rowIndex, short heightInPoints)
        {
            var row = sheet.CreateRow(rowIndex);
            row.HeightInPoints = heightInPoints;
            return row;
        }

    }
}
