using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace npoiDemo
{
    class ExportExcelDemo
    {
        private readonly IWorkbook _workbook;

        public ExportExcelDemo()
        {
            _workbook = new XSSFWorkbook() ;
        }


        public void ExportExcel()
        {
            const string newFile = @"项目动态收益管理.xlsx";

            using (var fs = new FileStream(newFile, FileMode.Create, FileAccess.Write))
            {
                
                ExportExcelProjectBasicInfo();

                _workbook.Write(fs);

            }
        }

        public void ExportExcelProjectBasicInfo()
        {

            var sheet = _workbook.CreateSheet("项目基本信息");
            
            var rowIndex = 0;
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowIndex, rowIndex, 0, 10));
            var row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("基本信息");
            
            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("项目编码");
            row.CreateCell(1).SetCellValue("022DC.001ygha01");
            row.CreateCell(2).SetCellValue("项目案名");
            row.CreateCell(3).SetCellValue("阳光海岸");
            row.CreateCell(4).SetCellValue("项目推广名");
            row.CreateCell(5).SetCellValue("雍海苑");
   

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("项目名称");
            row.CreateCell(1).SetCellValue("阳光海岸一期");
            row.CreateCell(2).SetCellValue("项目地块名称");
            row.CreateCell(3).SetCellValue("一期");
            row.CreateCell(4).SetCellValue("法人公司名称");
            row.CreateCell(5).SetCellValue("力高（天津）地产有限公司");
   
            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("产品系列");
            row.CreateCell(1).SetCellValue("雍系");
            row.CreateCell(2).SetCellValue("生产状态");
            row.CreateCell(3).SetCellValue("结案");
            row.CreateCell(4).SetCellValue("权益比例");
            row.CreateCell(5).SetCellValue(1);
       

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("土地获取日期");
            row.CreateCell(1).SetCellValue(DateTime.Parse("2009-04-29"));
            row.CreateCell(2).SetCellValue("土地开始使用日期");
            row.CreateCell(3).SetCellValue(DateTime.Parse("2009-09-16"));
            row.CreateCell(4).SetCellValue("土地截止使用日期");
            row.CreateCell(5).SetCellValue(DateTime.Parse("2059-09-15"));
          

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("用地性质");
            row.CreateCell(1).SetCellValue("其他");
            row.CreateCell(2).SetCellValue("土地获取方式");
            row.CreateCell(3).SetCellValue("其他");
            row.CreateCell(4).SetCellValue("项目地址");
            row.CreateCell(5).SetCellValue("天津滨海新区中新生态城海滨路");
         

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("所属区域");
            row.CreateCell(1).SetCellValue("省会重点城市");
            row.CreateCell(2).SetCellValue("所属城市");
            row.CreateCell(3).SetCellValue("");
            row.CreateCell(4).SetCellValue("所属城市行政区域");
            row.CreateCell(5).SetCellValue("");
         

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("项目负责人");
            row.CreateCell(1).SetCellValue("6938");
            row.CreateCell(2).SetCellValue("项目创建人");
            row.CreateCell(3).SetCellValue("李丽");
            row.CreateCell(4).SetCellValue("项目创建日期");
            row.CreateCell(5).SetCellValue(DateTime.Parse("2019-03-27"));
    

            rowIndex++;
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex + 1, 0, 0));
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex + 1, 1, 10));
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("项目描述");
            row.CreateCell(1).SetCellValue("一期及二期已经竣备，三期、四期南、四期北在施");
            
            
            for (var i = 0; i < 6; i++)
            {
                sheet.AutoSizeColumn(i);   
            }

            var fontHeader = _workbook.CreateFont();
            fontHeader.FontHeightInPoints = (short) 14;
            var cellStyleHeader = _workbook.CreateCellStyle();
            cellStyleHeader.SetFont(fontHeader);
            cellStyleHeader.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            cellStyleHeader.FillPattern = FillPattern.SolidForeground;
            sheet.GetRow(0).Cells[0].CellStyle = cellStyleHeader;
            sheet.GetRow(0).HeightInPoints = (short) 20;

            var cellStyleBody = _workbook.CreateCellStyle();
            cellStyleBody.BorderBottom = BorderStyle.Thin;

            for (var i = 1; i < sheet.LastRowNum; i++)
            {
                var tempRow = sheet.GetRow(i);
                for (var j = 0; j < 6; j++)
                {
                    if (j % 2 != 0)
                    {
                        tempRow.Cells[j].CellStyle = cellStyleBody;
                    }
                     
                }
                
            }

            var cellStylePercent = _workbook.CreateCellStyle();
            cellStylePercent.CloneStyleFrom(cellStyleBody);
            cellStylePercent.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");
            sheet.GetRow(3).Cells[5].CellStyle = cellStylePercent;

            var cellStyleDate = _workbook.CreateCellStyle();
            cellStyleDate.CloneStyleFrom(cellStyleBody);
            var dataFormatDate = _workbook.CreateDataFormat();
            cellStyleDate.DataFormat = dataFormatDate.GetFormat("yyyy-MM-dd");
            sheet.GetRow(4).Cells[1].CellStyle = cellStyleDate;
            sheet.GetRow(4).Cells[3].CellStyle = cellStyleDate;
            sheet.GetRow(4).Cells[5].CellStyle = cellStyleDate;
            sheet.GetRow(7).Cells[5].CellStyle = cellStyleDate;

            var cellStyleMergeLabel = _workbook.CreateCellStyle();
            cellStyleMergeLabel.CloneStyleFrom(cellStyleBody);
            cellStyleMergeLabel.Alignment = HorizontalAlignment.Left;
            cellStyleMergeLabel.VerticalAlignment = VerticalAlignment.Center;
            sheet.GetRow(8).Cells[0].CellStyle = cellStyleMergeLabel;

            var cellStyleMergeContent = _workbook.CreateCellStyle();
            cellStyleMergeContent.CloneStyleFrom(cellStyleBody);
            cellStyleMergeContent.Alignment = HorizontalAlignment.Left;
            cellStyleMergeContent.VerticalAlignment = VerticalAlignment.Bottom;
            cellStyleMergeContent.BorderBottom = BorderStyle.Thin;
            sheet.GetRow(8).Cells[1].CellStyle = cellStyleMergeContent;

            return ;

        }
    }
}
