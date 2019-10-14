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

        public const int MergeMaxColumnNum = 10;

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

                ExportExcelProjectInfo();

                ExportExcelPlanInfo();

                _workbook.Write(fs);

            }
        }

        public ICellStyle CellStyleHeader
        {
            get
            {
                var fontHeader = _workbook.CreateFont();
                fontHeader.FontHeightInPoints = (short)ExportExcelFontHeight.Header;
                var cellStyleHeader = _workbook.CreateCellStyle();
                cellStyleHeader.SetFont(fontHeader);
                cellStyleHeader.FillForegroundColor = HSSFColor.Grey40Percent.Index;
                cellStyleHeader.FillPattern = FillPattern.SolidForeground;
                return cellStyleHeader;
            }
        }

        public ICellStyle CellStyleTableHeader
        {
            get
            {
                var fontHeader = _workbook.CreateFont();
                fontHeader.FontHeightInPoints = (short)ExportExcelFontHeight.TableHeader;
                var cellStyleHeader = _workbook.CreateCellStyle();
                cellStyleHeader.SetFont(fontHeader);
                cellStyleHeader.FillForegroundColor = HSSFColor.Grey25Percent.Index;
                cellStyleHeader.FillPattern = FillPattern.SolidForeground;
                return cellStyleHeader;
            }
        }


        public ICellStyle CellStyleBody
        {
            get
            {
                var cellStyleBody = _workbook.CreateCellStyle();
                cellStyleBody.BorderBottom = BorderStyle.Thin;
                return cellStyleBody;
            }
        }

        public ICellStyle CellStyleDate
        {
            get
            {
                var cellStyleDate = _workbook.CreateCellStyle();
                cellStyleDate.CloneStyleFrom(CellStyleBody);
                var dataFormatDate = _workbook.CreateDataFormat();
                cellStyleDate.DataFormat = dataFormatDate.GetFormat("yyyy-MM-dd");
                return cellStyleDate;

            }
        }

        public ICellStyle CellStylePercent
        {
            get
            {
                var cellStylePercent = _workbook.CreateCellStyle();
                cellStylePercent.CloneStyleFrom(CellStyleBody);
                cellStylePercent.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00%");
                return cellStylePercent;
            }
        }

        public ICellStyle CellStyleArea
        {
            get
            {
                var cellStyle = _workbook.CreateCellStyle();
                cellStyle.CloneStyleFrom(CellStyleBody);
                cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0.00");
                return cellStyle;
            }
        }

        public ICellStyle CellStyleMergeLabel
        {
            get
            {
                var cellStyleMergeLabel = _workbook.CreateCellStyle();
                cellStyleMergeLabel.CloneStyleFrom(CellStyleBody);
                cellStyleMergeLabel.Alignment = HorizontalAlignment.Left;
                cellStyleMergeLabel.VerticalAlignment = VerticalAlignment.Center;
                return cellStyleMergeLabel;
            }
        }

        public ICellStyle CellStyleMergeContent
        {
            get
            {
                var cellStyleMergeContent = _workbook.CreateCellStyle();
                cellStyleMergeContent.CloneStyleFrom(CellStyleBody);
                cellStyleMergeContent.Alignment = HorizontalAlignment.Left;
                cellStyleMergeContent.VerticalAlignment = VerticalAlignment.Bottom;
                cellStyleMergeContent.BorderBottom = BorderStyle.Thin;
                return cellStyleMergeContent;
            }
        }

        /// <summary>
        /// 项目基本信息
        /// </summary>
        public void ExportExcelProjectInfo()
        {
            var sheet = _workbook.CreateSheet("项目基本信息");

            ExportExcelProjectBasicInfo();
            ExportExcelProjectProductInfo();
            ExportExcelProjectSectionInfo();

            for (var i = 0; i < 6; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            return;

        }

        /// <summary>
        /// </summary>
        public void ExportExcelPlanInfo()
        {
            var sheet = _workbook.CreateSheet("项目规划指标");

            ExportExcelPlanBasicInfo();
            ExportExcelPlanProjectInfo();
            ExportExcelPlanSectionInfo();

            for (var i = 0; i < 6; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            return;
        }


        public void ExportExcelProjectBasicInfo()
        {

            var sheet = _workbook.GetSheet("项目基本信息");

            var rowIndex = 0;
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowIndex, rowIndex, 0, MergeMaxColumnNum));
            var row = sheet.CreateRowWithHeightInPoints(rowIndex, (short)ExportExcelEnumRowPointHeight.Header);
            row.CreateCellWithCellStyle(0, CellStyleHeader).SetCellValue("基本信息");
         
            
            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("项目编码");
            row.CreateCellWithCellStyle(1, CellStyleBody).SetCellValue("022DC.001ygha01");
            row.CreateCell(2).SetCellValue("项目案名");
            row.CreateCellWithCellStyle(3, CellStyleBody).SetCellValue("阳光海岸");
            row.CreateCell(4).SetCellValue("项目推广名");
            row.CreateCellWithCellStyle(5, CellStyleBody).SetCellValue("雍海苑");
   

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("项目名称");
            row.CreateCellWithCellStyle(1, CellStyleBody).SetCellValue("阳光海岸一期");
            row.CreateCell(2).SetCellValue("项目地块名称");
            row.CreateCellWithCellStyle(3, CellStyleBody).SetCellValue("一期");
            row.CreateCell(4).SetCellValue("法人公司名称");
            row.CreateCellWithCellStyle(5, CellStyleBody).SetCellValue("力高（天津）地产有限公司");
   
            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("产品系列");
            row.CreateCellWithCellStyle(1, CellStyleBody).SetCellValue("雍系");
            row.CreateCell(2).SetCellValue("生产状态");
            row.CreateCellWithCellStyle(3, CellStyleBody).SetCellValue("结案");
            row.CreateCell(4).SetCellValue("权益比例");
            row.CreateCellWithCellStyle(5, CellStylePercent).SetCellValue(1);
            
            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("土地获取日期");
            row.CreateCellWithCellStyle(1, CellStyleDate).SetCellValue(DateTime.Parse("2009-04-29"));
            row.CreateCell(2).SetCellValue("土地开始使用日期");
            row.CreateCellWithCellStyle(3, CellStyleDate).SetCellValue(DateTime.Parse("2009-09-16"));
            row.CreateCell(4).SetCellValue("土地截止使用日期");
            row.CreateCellWithCellStyle(5, CellStyleDate).SetCellValue(DateTime.Parse("2059-09-15"));
            
            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("用地性质");
            row.CreateCellWithCellStyle(1, CellStyleBody).SetCellValue("其他");
            row.CreateCell(2).SetCellValue("土地获取方式");
            row.CreateCellWithCellStyle(3,CellStyleBody).SetCellValue("其他");
            row.CreateCell(4).SetCellValue("项目地址");
            row.CreateCellWithCellStyle(5,CellStyleBody).SetCellValue("天津滨海新区中新生态城海滨路");
            
            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("所属区域");
            row.CreateCellWithCellStyle(1, CellStyleBody).SetCellValue("省会重点城市");
            row.CreateCell(2).SetCellValue("所属城市");
            row.CreateCellWithCellStyle(3,CellStyleBody).SetCellValue("");
            row.CreateCell(4).SetCellValue("所属城市行政区域");
            row.CreateCellWithCellStyle(5,CellStyleBody).SetCellValue("");
            
            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("项目负责人");
            row.CreateCellWithCellStyle(1,CellStyleBody).SetCellValue("6938");
            row.CreateCell(2).SetCellValue("项目创建人");
            row.CreateCellWithCellStyle(3,CellStyleBody).SetCellValue("李丽");
            row.CreateCell(4).SetCellValue("项目创建日期");
            row.CreateCellWithCellStyle(5,CellStyleDate).SetCellValue(DateTime.Parse("2019-03-27"));
            
            rowIndex++;
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex + 1, 0, 0));
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex + 1, 1, MergeMaxColumnNum));
            row = sheet.CreateRow(rowIndex);
            row.CreateCellWithCellStyle(0, CellStyleMergeLabel).SetCellValue("项目描述");
            row.CreateCellWithCellStyle(1, CellStyleMergeContent).SetCellValue("一期及二期已经竣备，三期、四期南、四期北在施");
            

        }

        public void ExportExcelProjectProductInfo()
        {
            var sheet = _workbook.GetSheet("项目基本信息");
            var rowIndex = sheet.LastRowNum;

            rowIndex++;
            rowIndex++;
            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, MergeMaxColumnNum));
            var row = sheet.CreateRowWithHeightInPoints(rowIndex, (short)ExportExcelEnumRowPointHeight.Header);
            row.CreateCellWithCellStyle(0, CellStyleHeader).SetCellValue("业态信息");
            
            rowIndex++;
            row = sheet.CreateRowWithHeightInPoints(rowIndex, (short)ExportExcelEnumRowPointHeight.TableHeader);
            row.CreateCellWithCellStyle(0, CellStyleTableHeader).SetCellValue("产品大类");
            row.CreateCellWithCellStyle(1, CellStyleTableHeader).SetCellValue("产品名称");
            row.CreateCellWithCellStyle(2, CellStyleTableHeader).SetCellValue("产品类型");
            row.CreateCellWithCellStyle(3, CellStyleTableHeader).SetCellValue("产品系列");

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("产品大类") ;
            row.CreateCell(1).SetCellValue("车位");
            row.CreateCell(2).SetCellValue("");
            row.CreateCell(3).SetCellValue("雍系");

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("配套");
            row.CreateCell(1).SetCellValue("配建");
            row.CreateCell(2).SetCellValue("");
            row.CreateCell(3).SetCellValue("雍系");


            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("商业");
            row.CreateCell(1).SetCellValue("商业");
            row.CreateCell(2).SetCellValue("");
            row.CreateCell(3).SetCellValue("雍系");


            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("住宅");
            row.CreateCell(1).SetCellValue("住宅");
            row.CreateCell(2).SetCellValue("");
            row.CreateCell(3).SetCellValue("雍系");

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("住宅");
            row.CreateCell(1).SetCellValue("住宅");
            row.CreateCell(2).SetCellValue("");
            row.CreateCell(3).SetCellValue("雍系");

            return;

        }

        public void ExportExcelProjectSectionInfo()
        {
            var sheet = _workbook.GetSheet("项目基本信息");
           
            var rowIndex = sheet.LastRowNum;
            rowIndex ++;

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, MergeMaxColumnNum));

            var row = sheet.CreateRowWithHeightInPoints(rowIndex,(short)ExportExcelEnumRowPointHeight.Header);
            row.CreateCellWithCellStyle(0, CellStyleHeader).SetCellValue("标段信息");

            rowIndex++;
            row = sheet.CreateRowWithHeightInPoints(rowIndex, (short)ExportExcelEnumRowPointHeight.TableHeader);
            
            row.CreateCellWithCellStyle(0, CellStyleTableHeader).SetCellValue("标段名称");
            row.CreateCellWithCellStyle(1,CellStyleTableHeader).SetCellValue("对应楼栋");
            row.CreateCellWithCellStyle(2, CellStyleTableHeader).SetCellValue("产品名称");

            for (int i = 0; i < 49; i++)
            {
                rowIndex++;
                row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("一标段");
                row.CreateCell(1).SetCellValue($"别墅-{i+1}栋");
                row.CreateCell(2).SetCellValue("住宅");
            }


            return;

        }


      

        public void ExportExcelPlanBasicInfo()
        {

            var sheet = _workbook.GetSheet("项目规划指标");

            var rowIndex = 0;

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, MergeMaxColumnNum));
            var row = sheet.CreateRowWithHeightInPoints(rowIndex, (short)ExportExcelEnumRowPointHeight.Header);
            row.CreateCellWithCellStyle(0,CellStyleHeader).SetCellValue("基本信息");

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("总占地面积");
            row.CreateCellWithCellStyle(1, CellStyleArea).SetCellValue(132787.00);
            row.CreateCell(2).SetCellValue("建设占地面积");
            row.CreateCellWithCellStyle(3, CellStyleArea).SetCellValue(132787.00);
            row.CreateCell(4).SetCellValue("代征地的面积");
            row.CreateCellWithCellStyle(5, CellStyleArea).SetCellValue(0.00);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("建筑占地面积");
            row.CreateCellWithCellStyle(1, CellStyleArea).SetCellValue(51091.49);
            row.CreateCell(2).SetCellValue("容积率");
            row.CreateCellWithCellStyle(3, CellStyleArea).SetCellValue(1.07);
            row.CreateCell(4).SetCellValue("建筑密度");
            row.CreateCellWithCellStyle(5, CellStylePercent).SetCellValue(0.38);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("总建筑面积");
            row.CreateCellWithCellStyle(1, CellStyleArea).SetCellValue(184949.23);
            row.CreateCell(2).SetCellValue("地上建筑面积");
            row.CreateCellWithCellStyle(3, CellStyleArea).SetCellValue(141722.19);
            row.CreateCell(4).SetCellValue("地下建筑面积");
            row.CreateCellWithCellStyle(5, CellStyleArea).SetCellValue(432278.04);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("总可售面积");
            row.CreateCellWithCellStyle(1, CellStyleArea).SetCellValue(135620.14);
            row.CreateCell(2).SetCellValue("地上可售面积");
            row.CreateCellWithCellStyle(3, CellStyleArea).SetCellValue(135620.14);
            row.CreateCell(4).SetCellValue("地下可售面积");
            row.CreateCellWithCellStyle(5, CellStyleArea).SetCellValue(17252.73);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("总可租面积");
            row.CreateCellWithCellStyle(1, CellStyleArea).SetCellValue(0.00);
            row.CreateCell(2).SetCellValue("地上可租面积");
            row.CreateCellWithCellStyle(3, CellStyleArea).SetCellValue(0.00);
            row.CreateCell(4).SetCellValue("地下可租面积");
            row.CreateCellWithCellStyle(5, CellStyleArea).SetCellValue(0.00);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("总计容面积");
            row.CreateCellWithCellStyle(1, CellStyleArea).SetCellValue(0.00);
            row.CreateCell(2).SetCellValue("地上计容面积");
            row.CreateCellWithCellStyle(3, CellStyleArea).SetCellValue(0.00);
            row.CreateCell(4).SetCellValue("地下计容面积");
            row.CreateCellWithCellStyle(5, CellStyleArea).SetCellValue(0.00);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("总还建面积");
            row.CreateCellWithCellStyle(1, CellStyleArea).SetCellValue(0.00);
            row.CreateCell(2).SetCellValue("景观面积");
            row.CreateCellWithCellStyle(3, CellStyleArea).SetCellValue(39836.10);
            row.CreateCell(4).SetCellValue("用地红线外景观面积");
            row.CreateCellWithCellStyle(5, CellStyleArea).SetCellValue(0.00);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("用地红线内景观面积");
            row.CreateCellWithCellStyle(1, CellStyleArea).SetCellValue(0.00);
            row.CreateCell(2).SetCellValue("屋顶景观面积");
            row.CreateCellWithCellStyle(3, CellStyleArea).SetCellValue(0.00);
            row.CreateCell(4).SetCellValue("垂直绿化面积");
            row.CreateCellWithCellStyle(5, CellStyleArea).SetCellValue(0.00);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("车位");
            row.CreateCellWithCellStyle(1, CellStyleArea).SetCellValue(577.00);
            row.CreateCell(2).SetCellValue("地上车位");
            row.CreateCellWithCellStyle(3, CellStyleArea).SetCellValue(77.00);
            row.CreateCell(4).SetCellValue("地下人防车位");
            row.CreateCellWithCellStyle(5, CellStyleArea).SetCellValue(0.00);


            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("地下非人防车位");
            row.CreateCellWithCellStyle(1, CellStyleArea).SetCellValue(500.00);
            row.CreateCell(2).SetCellValue("标准车位面积");
            row.CreateCellWithCellStyle(3, CellStyleArea).SetCellValue(34.00);
            row.CreateCell(4).SetCellValue("是否回签");
            row.CreateCellWithCellStyle(5, CellStyleBody).SetCellValue("否");

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("是否还建");
            row.CreateCellWithCellStyle(1, CellStyleArea).SetCellValue("否");

            return;

        }

        public void ExportExcelPlanProjectInfo()
        {
            var sheet = _workbook.GetSheet("项目规划指标");

            var rowIndex = sheet.LastRowNum;
            rowIndex++;

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, MergeMaxColumnNum));
            var row = sheet.CreateRowWithHeightInPoints(rowIndex, (short) ExportExcelEnumRowPointHeight.Header);
            row.CreateCellWithCellStyle(0,CellStyleHeader).SetCellValue("产品规划指标");

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCellWithCellStyle(0, CellStyleTableHeader).SetCellValue("产品大类");
            row.CreateCellWithCellStyle(1, CellStyleTableHeader).SetCellValue("产品类型");
            row.CreateCellWithCellStyle(2, CellStyleTableHeader).SetCellValue("产品名称");
            row.CreateCellWithCellStyle(3, CellStyleTableHeader).SetCellValue("地上建筑面积(平米)");
            row.CreateCellWithCellStyle(4, CellStyleTableHeader).SetCellValue("地下建筑面积(平米)");
            row.CreateCellWithCellStyle(5, CellStyleTableHeader).SetCellValue("总建筑面积(平米)");
            row.CreateCellWithCellStyle(6, CellStyleTableHeader).SetCellValue("可售面积(平米)");
            row.CreateCellWithCellStyle(7, CellStyleTableHeader).SetCellValue("可租面积(平米)");
            row.CreateCellWithCellStyle(8, CellStyleTableHeader).SetCellValue("还建面积(平米)");
            row.CreateCellWithCellStyle(9, CellStyleTableHeader).SetCellValue("栋数(栋)");
            row.CreateCellWithCellStyle(10, CellStyleTableHeader).SetCellValue("层数(层)");
            row.CreateCellWithCellStyle(11, CellStyleTableHeader).SetCellValue("单元数(个)");
            row.CreateCellWithCellStyle(12, CellStyleTableHeader).SetCellValue("户数(户)");
            row.CreateCellWithCellStyle(13, CellStyleTableHeader).SetCellValue("层高-首/标(米)");
            row.CreateCellWithCellStyle(14, CellStyleTableHeader).SetCellValue("建筑面积(平米)");

            rowIndex++;

            var planRowCount = 5;
            var planColumnCount = 15;
            var cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex + planRowCount, 0, planColumnCount);

            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("车库/车房");
            row.CreateCell(1).SetCellValue("产权车位");
            row.CreateCell(2).SetCellValue("车位");
            row.CreateCell(3).SetCellValue(0);
            row.CreateCell(4).SetCellValue(22529);
            row.CreateCell(5).SetCellValue(0);
            row.CreateCell(6).SetCellValue(0);
            row.CreateCell(7).SetCellValue(0);
            row.CreateCell(8).SetCellValue(0);
            row.CreateCell(9).SetCellValue(1);
            row.CreateCell(10).SetCellValue(0);
            row.CreateCell(11).SetCellValue(1);
            row.CreateCell(12).SetCellValue(1);
            row.CreateCell(13).SetCellValue(4);
            row.CreateCell(14).SetCellValue(0);


            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("配套");
            row.CreateCell(1).SetCellValue("幼儿园");
            row.CreateCell(2).SetCellValue("配建");
            row.CreateCell(3).SetCellValue(3161);
            row.CreateCell(4).SetCellValue(0);
            row.CreateCell(5).SetCellValue(0);
            row.CreateCell(6).SetCellValue(0);
            row.CreateCell(7).SetCellValue(0);
            row.CreateCell(8).SetCellValue(0);
            row.CreateCell(9).SetCellValue(1);
            row.CreateCell(10).SetCellValue(0);
            row.CreateCell(11).SetCellValue(1);
            row.CreateCell(12).SetCellValue(1);
            row.CreateCell(13).SetCellValue(4);
            row.CreateCell(14).SetCellValue(0);



            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("商业");
            row.CreateCell(1).SetCellValue("独立商业");
            row.CreateCell(2).SetCellValue("商业");
            row.CreateCell(3).SetCellValue(12571);
            row.CreateCell(4).SetCellValue(1316);
            row.CreateCell(5).SetCellValue(0);
            row.CreateCell(6).SetCellValue(9360);
            row.CreateCell(7).SetCellValue(0);
            row.CreateCell(8).SetCellValue(0);
            row.CreateCell(9).SetCellValue(3);
            row.CreateCell(10).SetCellValue(0);
            row.CreateCell(11).SetCellValue(3);
            row.CreateCell(12).SetCellValue(3);
            row.CreateCell(13).SetCellValue(4);
            row.CreateCell(14).SetCellValue(0);


            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("住宅");
            row.CreateCell(1).SetCellValue("别墅");
            row.CreateCell(2).SetCellValue("住宅");
            row.CreateCell(3).SetCellValue(55581);
            row.CreateCell(4).SetCellValue(17253);
            row.CreateCell(5).SetCellValue(0);
            row.CreateCell(6).SetCellValue(55581);
            row.CreateCell(7).SetCellValue(0);
            row.CreateCell(8).SetCellValue(0);
            row.CreateCell(9).SetCellValue(49);
            row.CreateCell(10).SetCellValue(0);
            row.CreateCell(11).SetCellValue(190);
            row.CreateCell(12).SetCellValue(190);
            row.CreateCell(13).SetCellValue(3);
            row.CreateCell(14).SetCellValue(0);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("住宅");
            row.CreateCell(1).SetCellValue("高层");
            row.CreateCell(2).SetCellValue("住宅");
            row.CreateCell(3).SetCellValue(70409);
            row.CreateCell(4).SetCellValue(2129);
            row.CreateCell(5).SetCellValue(0);
            row.CreateCell(6).SetCellValue(70409);
            row.CreateCell(7).SetCellValue(0);
            row.CreateCell(8).SetCellValue(0);
            row.CreateCell(9).SetCellValue(5);
            row.CreateCell(10).SetCellValue(0);
            row.CreateCell(11).SetCellValue(5);
            row.CreateCell(12).SetCellValue(614);
            row.CreateCell(13).SetCellValue(3);
            row.CreateCell(14).SetCellValue(0);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("合计");
            row.CreateCell(1).SetCellValue("");
            row.CreateCell(2).SetCellValue("");
            row.CreateCell(3).SetCellFormula($"Sum(D{cellRangeAddress.FirstRow + 1}:D{cellRangeAddress.LastRow })");
            row.CreateCell(4).SetCellFormula($"Sum(E{cellRangeAddress.FirstRow + 1}:E{cellRangeAddress.LastRow })");
            row.CreateCell(5).SetCellFormula($"Sum(F{cellRangeAddress.FirstRow + 1}:F{cellRangeAddress.LastRow })");
            row.CreateCell(6).SetCellFormula($"Sum(G{cellRangeAddress.FirstRow + 1}:G{cellRangeAddress.LastRow })");
            row.CreateCell(7).SetCellFormula($"Sum(H{cellRangeAddress.FirstRow + 1}:H{cellRangeAddress.LastRow })");
            row.CreateCell(8).SetCellFormula($"Sum(I{cellRangeAddress.FirstRow + 1}:I{cellRangeAddress.LastRow })");
            row.CreateCell(9).SetCellFormula($"Sum(J{cellRangeAddress.FirstRow + 1}:J{cellRangeAddress.LastRow })");
            row.CreateCell(10).SetCellFormula($"Sum(K{cellRangeAddress.FirstRow + 1}:K{cellRangeAddress.LastRow })");
            row.CreateCell(11).SetCellFormula($"Sum(L{cellRangeAddress.FirstRow + 1}:L{cellRangeAddress.LastRow })");
            row.CreateCell(12).SetCellFormula($"Sum(M{cellRangeAddress.FirstRow + 1}:M{cellRangeAddress.LastRow })");
            row.CreateCell(13).SetCellFormula($"Sum(N{cellRangeAddress.FirstRow + 1}:N{cellRangeAddress.LastRow })");
            row.CreateCell(14).SetCellFormula($"Sum(O{cellRangeAddress.FirstRow + 1}:O{cellRangeAddress.LastRow })");

            return;

        }

        public void ExportExcelPlanSectionInfo()
        {
            var sheet = _workbook.GetSheet("项目规划指标");

            var rowIndex = sheet.LastRowNum;
            rowIndex++;

            sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, MergeMaxColumnNum));
            var row = sheet.CreateRowWithHeightInPoints(rowIndex, (short)ExportExcelEnumRowPointHeight.Header);
            row.CreateCellWithCellStyle(0, CellStyleHeader).SetCellValue("标段规划指标");

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCellWithCellStyle(0, CellStyleTableHeader).SetCellValue("标段名称");
            row.CreateCellWithCellStyle(1, CellStyleTableHeader).SetCellValue("产品大类");
            row.CreateCellWithCellStyle(2, CellStyleTableHeader).SetCellValue("产品类型");
            row.CreateCellWithCellStyle(3, CellStyleTableHeader).SetCellValue("产品名称");
            row.CreateCellWithCellStyle(4, CellStyleTableHeader).SetCellValue("建设用地面积(平米)");
            row.CreateCellWithCellStyle(5, CellStyleTableHeader).SetCellValue("地上建筑面积(平米)");
            row.CreateCellWithCellStyle(6, CellStyleTableHeader).SetCellValue("地下建筑面积(平米)");
            row.CreateCellWithCellStyle(7, CellStyleTableHeader).SetCellValue("总建筑面积(平米)");
            row.CreateCellWithCellStyle(8, CellStyleTableHeader).SetCellValue("可售面积(平米)");
            row.CreateCellWithCellStyle(9, CellStyleTableHeader).SetCellValue("可租面积(平米)");
            row.CreateCellWithCellStyle(10, CellStyleTableHeader).SetCellValue("户数(户)");
            row.CreateCellWithCellStyle(11, CellStyleTableHeader).SetCellValue("建筑占地面积(平米)");


            rowIndex++;
            var firstSectionBuildingAreaRowIndex = rowIndex;
            row = sheet.CreateRow(rowIndex);
            const int firstSectionRowCount = 5;
            const int firstSectionColCount = 12;
            var firstSectionCellRangeAddress = new CellRangeAddress(rowIndex+1, rowIndex+1+firstSectionRowCount, 
                0,firstSectionColCount);
            row.CreateCell(0).SetCellValue("一标段");
            row.CreateCell(1).SetCellValue("");
            row.CreateCell(2).SetCellValue("");
            row.CreateCell(3).SetCellValue("");
            row.CreateCellWithCellStyle(4, CellStyleArea).SetCellValue(0.00);
            row.CreateCell(5).SetCellFormula($"Sum(F{firstSectionCellRangeAddress.FirstRow + 1}:F{firstSectionCellRangeAddress.LastRow})");
            row.CreateCell(6).SetCellFormula($"Sum(G{firstSectionCellRangeAddress.FirstRow + 1}:G{firstSectionCellRangeAddress.LastRow})");
            row.CreateCell(7).SetCellFormula($"Sum(H{firstSectionCellRangeAddress.FirstRow + 1}:H{firstSectionCellRangeAddress.LastRow})");
            row.CreateCell(8).SetCellFormula($"Sum(I{firstSectionCellRangeAddress.FirstRow + 1}:I{firstSectionCellRangeAddress.LastRow})");
            row.CreateCell(9).SetCellFormula($"Sum(J{firstSectionCellRangeAddress.FirstRow + 1}:J{firstSectionCellRangeAddress.LastRow})");
            row.CreateCell(10).SetCellFormula($"Sum(K{firstSectionCellRangeAddress.FirstRow + 1}:K{firstSectionCellRangeAddress.LastRow})");
            row.CreateCell(11).SetCellFormula($"Sum(H{firstSectionCellRangeAddress.FirstRow + 1}:H{firstSectionCellRangeAddress.LastRow})");

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("一标段");
            row.CreateCell(1).SetCellValue("车库/车房");
            row.CreateCell(2).SetCellValue("产权车位");
            row.CreateCell(3).SetCellValue("车位");
            row.CreateCell(4).SetCellValue(0);
            row.CreateCell(5).SetCellValue(0);
            row.CreateCell(6).SetCellValue(22529);
            row.CreateCell(7).SetCellValue(0);
            row.CreateCell(8).SetCellValue(0);
            row.CreateCell(9).SetCellValue(0);
            row.CreateCell(10).SetCellValue(1);
            row.CreateCell(11).SetCellValue(22529);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("一标段");
            row.CreateCell(1).SetCellValue("配套");
            row.CreateCell(2).SetCellValue("幼儿园");
            row.CreateCell(3).SetCellValue("配建");
            row.CreateCell(4).SetCellValue(0);
            row.CreateCell(5).SetCellValue(3161);
            row.CreateCell(6).SetCellValue(0);
            row.CreateCell(7).SetCellValue(0);
            row.CreateCell(8).SetCellValue(0);
            row.CreateCell(9).SetCellValue(0);
            row.CreateCell(10).SetCellValue(1);
            row.CreateCell(11).SetCellValue(1054);


            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("一标段");
            row.CreateCell(1).SetCellValue("商业");
            row.CreateCell(2).SetCellValue("独立商业");
            row.CreateCell(3).SetCellValue("商业");
            row.CreateCell(4).SetCellValue(0);
            row.CreateCell(5).SetCellValue(12571);
            row.CreateCell(6).SetCellValue(1316);
            row.CreateCell(7).SetCellValue(0);
            row.CreateCell(8).SetCellValue(9630);
            row.CreateCell(9).SetCellValue(0);
            row.CreateCell(10).SetCellValue(3);
            row.CreateCell(11).SetCellValue(6286);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("一标段");
            row.CreateCell(1).SetCellValue("住宅");
            row.CreateCell(2).SetCellValue("别墅");
            row.CreateCell(3).SetCellValue("住宅");
            row.CreateCell(4).SetCellValue(0);
            row.CreateCell(5).SetCellValue(54621);
            row.CreateCell(6).SetCellValue(16990);
            row.CreateCell(7).SetCellValue(0);
            row.CreateCell(8).SetCellValue(54621);
            row.CreateCell(9).SetCellValue(0);
            row.CreateCell(10).SetCellValue(186);
            row.CreateCell(11).SetCellValue(18719);

            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("一标段");
            row.CreateCell(1).SetCellValue("住宅");
            row.CreateCell(2).SetCellValue("高层");
            row.CreateCell(3).SetCellValue("住宅");
            row.CreateCell(4).SetCellValue(0);
            row.CreateCell(5).SetCellValue(70409);
            row.CreateCell(6).SetCellValue(2129);
            row.CreateCell(7).SetCellValue(0);
            row.CreateCell(8).SetCellValue(70409);
            row.CreateCell(9).SetCellValue(0);
            row.CreateCell(10).SetCellValue(614);
            row.CreateCell(11).SetCellValue(2129);


            rowIndex++;

            const int secondSectionRowCount = 1;
            const int secondSectionColCount = 12;
            var secondSectionRowRangeAddress = new CellRangeAddress(rowIndex + 1, rowIndex + 1 + secondSectionRowCount,
                0, secondSectionColCount);
            
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("二标段");
            row.CreateCell(1).SetCellValue("");
            row.CreateCell(2).SetCellValue("");
            row.CreateCell(3).SetCellValue("");

            var secondSectionBuildingAreaRowIndex = rowIndex;
            row.CreateCellWithCellStyle(4, CellStyleArea).SetCellValue(0);
            row.CreateCell(5).SetCellFormula($"Sum(F{secondSectionRowRangeAddress.FirstRow + 1}:F{secondSectionRowRangeAddress.LastRow})");
            row.CreateCell(6).SetCellFormula($"Sum(G{secondSectionRowRangeAddress.FirstRow + 1}:G{secondSectionRowRangeAddress.LastRow})");
            row.CreateCell(7).SetCellFormula($"Sum(H{secondSectionRowRangeAddress.FirstRow + 1}:H{secondSectionRowRangeAddress.LastRow})");
            row.CreateCell(8).SetCellFormula($"Sum(I{secondSectionRowRangeAddress.FirstRow + 1}:I{secondSectionRowRangeAddress.LastRow})");
            row.CreateCell(9).SetCellFormula($"Sum(J{secondSectionRowRangeAddress.FirstRow + 1}:J{secondSectionRowRangeAddress.LastRow})");
            row.CreateCell(10).SetCellFormula($"Sum(K{secondSectionRowRangeAddress.FirstRow + 1}:K{secondSectionRowRangeAddress.LastRow})");
            row.CreateCell(11).SetCellFormula($"Sum(H{secondSectionRowRangeAddress.FirstRow + 1}:H{secondSectionRowRangeAddress.LastRow})");


            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("二标段");
            row.CreateCell(1).SetCellValue("住宅");
            row.CreateCell(2).SetCellValue("别墅");
            row.CreateCell(3).SetCellValue("住宅");
            row.CreateCell(4).SetCellValue(0);
            row.CreateCell(5).SetCellValue(960);
            row.CreateCell(6).SetCellValue(263);
            row.CreateCell(7).SetCellValue(0);
            row.CreateCell(8).SetCellValue(960);
            row.CreateCell(9).SetCellValue(0);
            row.CreateCell(10).SetCellValue(4);
            row.CreateCell(11).SetCellValue(320);



            rowIndex++;
            row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("合计");
            row.CreateCell(1).SetCellValue("");
            row.CreateCell(2).SetCellValue("");
            row.CreateCell(3).SetCellValue("");
            row.CreateCell(4).SetCellFormula($"Sum(E{firstSectionBuildingAreaRowIndex + 1},E{secondSectionBuildingAreaRowIndex + 1})");
            row.CreateCell(5).SetCellFormula($"Sum(F{firstSectionCellRangeAddress.FirstRow + 1}:F{firstSectionCellRangeAddress.LastRow}," +
                                             $" F{secondSectionRowRangeAddress.FirstRow + 1}:F{secondSectionRowRangeAddress.LastRow})");
            row.CreateCell(6).SetCellFormula($"Sum(G{firstSectionCellRangeAddress.FirstRow + 1}:G{firstSectionCellRangeAddress.LastRow}," +
                                             $" G{secondSectionRowRangeAddress.FirstRow + 1}:G{secondSectionRowRangeAddress.LastRow})");
            row.CreateCell(7).SetCellFormula($"Sum(H{firstSectionCellRangeAddress.FirstRow + 1}:H{firstSectionCellRangeAddress.LastRow}," +
                                             $" H{secondSectionRowRangeAddress.FirstRow + 1}:H{secondSectionRowRangeAddress.LastRow})");
            row.CreateCell(8).SetCellFormula($"Sum(I{firstSectionCellRangeAddress.FirstRow + 1}:I{firstSectionCellRangeAddress.LastRow}," +
                                             $" I{secondSectionRowRangeAddress.FirstRow + 1}:I{secondSectionRowRangeAddress.LastRow})");
            row.CreateCell(9).SetCellFormula($"Sum(J{firstSectionCellRangeAddress.FirstRow + 1}:J{firstSectionCellRangeAddress.LastRow}," +
                                             $" J{secondSectionRowRangeAddress.FirstRow + 1}:J{secondSectionRowRangeAddress.LastRow})");
            row.CreateCell(10).SetCellFormula($"Sum(K{firstSectionCellRangeAddress.FirstRow + 1}:K{firstSectionCellRangeAddress.LastRow}," +
                                             $" K{secondSectionRowRangeAddress.FirstRow + 1}:K{secondSectionRowRangeAddress.LastRow})");
            row.CreateCell(11).SetCellFormula($"Sum(L{firstSectionCellRangeAddress.FirstRow + 1}:L{firstSectionCellRangeAddress.LastRow}," +
                                             $" L{secondSectionRowRangeAddress.FirstRow + 1}:L{secondSectionRowRangeAddress.LastRow})");



            return;

        }
    }
}
