using System;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;


namespace GM.DocumentUtility  
{
    public class ExcelTemplate   
    {
       private HSSFWorkbook wb;  
       private static string fontName = "Tohama";  
       private static short fontHeight = 10;  
       private short colHeaderColor;
       private short colDetailColor;  
       private short colBorderColor;  
       private short colFooterColor;

       private ICellStyle headerLeftCellStyle;   
       private ICellStyle headerRightCellStyle;  
       private ICellStyle colHeadCellStyle;  
       private ICellStyle colLeftCellStyle;  
       private ICellStyle colRightCellStyle;
       private ICellStyle colCenterCellStyle;  
       private ICellStyle colNumberCellStyle;  
       private ICellStyle col1DecimalCellStyle;  
       private ICellStyle col2DecimalCellStyle;  
       private ICellStyle col4DecimalCellStyle;   
       private ICellStyle col6DecimalCellStyle;  
       private ICellStyle footerCellStyle;
       private ICellStyle footerRightCellStyle;
       private ICellStyle footerCenterCellStyle;
       private ICellStyle footerNumberCellStyle;
       private ICellStyle footer1DecimalCellStyle;
       private ICellStyle footer2DecimalCellStyle;
       private ICellStyle footer4DecimalCellStyle;
       private ICellStyle footer6DecimalCellStyle;

        private ICellStyle col2DecimalBorderBottomCellStyle;

        public ExcelTemplate(HSSFWorkbook workbook,
            short colHeaderColor = 48, short colDetailColor = 9,
            short colFooterColor = 44, short colBorderColor = 22)
        {
            wb = workbook;
            //colHeaderColor = SetColor(36, 160, 244);
            //colDetailColor = SetColor(232, 242, 254);
            //colBorderColor = SetColor(172, 185, 202);
            //colFooterColor = SetColor(47, 117, 181);

            //colHeaderColor = 48;
            //colDetailColor = 9;
            //colBorderColor = 22;
            //colFooterColor = 44;

            headerLeftCellStyle = HeaderCellStyle();  

            headerRightCellStyle = HeaderCellStyle(horizontalAlignment: HorizontalAlignment.Right);   

            colHeadCellStyle = CellStyle(fontBoldWeight: FontBoldWeight.Bold,  
              verticalAlignment: VerticalAlignment.center, horizontalAlignment: HorizontalAlignment.Center,  
              borderColor: colBorderColor, foregroundColor: colHeaderColor, fontColor: HSSFColor.White.Index);   

            colLeftCellStyle = CellStyle(borderColor: colBorderColor, foregroundColor: colDetailColor);   

            colRightCellStyle = CellStyle(horizontalAlignment: HorizontalAlignment.Right, borderColor: colBorderColor, foregroundColor: colDetailColor);      

            colCenterCellStyle = CellStyle(horizontalAlignment: HorizontalAlignment.Center, borderColor: colBorderColor, foregroundColor: colDetailColor);

            colNumberCellStyle = CellStyle(horizontalAlignment: HorizontalAlignment.Right, borderColor: colBorderColor, foregroundColor: colDetailColor,
                dataFormat: "#,##0");

            col1DecimalCellStyle = CellStyle(horizontalAlignment: HorizontalAlignment.Right, borderColor: colBorderColor, foregroundColor: colDetailColor,
                dataFormat: "#,##0.0");

            col2DecimalCellStyle = CellStyle(horizontalAlignment: HorizontalAlignment.Right, borderColor: colBorderColor, foregroundColor: colDetailColor,
                dataFormat: "#,##0.00");

            col4DecimalCellStyle = CellStyle(horizontalAlignment: HorizontalAlignment.Right, borderColor: colBorderColor, foregroundColor: colDetailColor,
                dataFormat: "#,##0.0000");

            col6DecimalCellStyle = CellStyle(horizontalAlignment: HorizontalAlignment.Right, borderColor: colBorderColor, foregroundColor: colDetailColor,
                dataFormat: "#,##0.000000");

            footerCellStyle = CellStyle(fontBoldWeight: FontBoldWeight.Bold, borderColor: colBorderColor, foregroundColor: colFooterColor);

            footerRightCellStyle = CellStyle(fontBoldWeight: FontBoldWeight.Bold, horizontalAlignment: HorizontalAlignment.Right, borderColor: colBorderColor, foregroundColor: colFooterColor);

            footerCenterCellStyle = CellStyle(fontBoldWeight: FontBoldWeight.Bold, horizontalAlignment: HorizontalAlignment.Center, borderColor: colBorderColor, foregroundColor: colFooterColor);

            footerNumberCellStyle = CellStyle(fontBoldWeight: FontBoldWeight.Bold, horizontalAlignment: HorizontalAlignment.Right, borderColor: colBorderColor, foregroundColor: colFooterColor,
                dataFormat: "#,##0");

            footer1DecimalCellStyle = CellStyle(fontBoldWeight: FontBoldWeight.Bold, horizontalAlignment: HorizontalAlignment.Right, borderColor: colBorderColor, foregroundColor: colFooterColor,
                dataFormat: "#,##0.0");

            footer2DecimalCellStyle = CellStyle(fontBoldWeight: FontBoldWeight.Bold, horizontalAlignment: HorizontalAlignment.Right, borderColor: colBorderColor, foregroundColor: colFooterColor,
                dataFormat: "#,##0.00");

            footer4DecimalCellStyle = CellStyle(fontBoldWeight: FontBoldWeight.Bold, horizontalAlignment: HorizontalAlignment.Right, borderColor: colBorderColor, foregroundColor: colFooterColor,
                dataFormat: "#,##0.0000");

            footer6DecimalCellStyle = CellStyle(fontBoldWeight: FontBoldWeight.Bold, horizontalAlignment: HorizontalAlignment.Right, borderColor: colBorderColor, foregroundColor: colFooterColor,
                dataFormat: "#,##0.000000");


        }

        #region CellStyle
        public ICellStyle HeaderCellStyle(FontBoldWeight fontBoldWeight = FontBoldWeight.Normal,
            HorizontalAlignment horizontalAlignment = HorizontalAlignment.Left, short fontHeight = 10)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            IFont font = wb.CreateFont();
            font.FontName = fontName;
            font.FontHeightInPoints = fontHeight;
            font.Boldweight = (short)fontBoldWeight;
            cellStyle.Alignment = horizontalAlignment;
            cellStyle.SetFont(font);
            return cellStyle;
        }

        public ICellStyle CellStyle(FontBoldWeight fontBoldWeight = FontBoldWeight.Normal,
            VerticalAlignment verticalAlignment = VerticalAlignment.Bottom, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Left,
            string dataFormat = "", short borderColor = HSSFColor.COLOR_NORMAL, short foregroundColor = HSSFColor.COLOR_NORMAL,
            short fontColor = HSSFColor.Black.Index, FontUnderlineType underline = FontUnderlineType.None)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            IFont font = wb.CreateFont();
            font.Color = fontColor;
            font.FontName = fontName;
            font.FontHeightInPoints = fontHeight;
            font.Boldweight = (short)fontBoldWeight;
            font.Underline = underline;
            cellStyle.VerticalAlignment = verticalAlignment;
            cellStyle.Alignment = horizontalAlignment;
            cellStyle.FillForegroundColor = foregroundColor;
            cellStyle.FillPattern = FillPattern.SolidForeground;
            cellStyle.RightBorderColor = borderColor;
            cellStyle.LeftBorderColor = borderColor;
            cellStyle.TopBorderColor = borderColor;
            cellStyle.BottomBorderColor = borderColor;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            if (!string.IsNullOrEmpty(dataFormat))
            {
                cellStyle.DataFormat = wb.CreateDataFormat().GetFormat(dataFormat);
            }
            cellStyle.SetFont(font);
            return cellStyle;
        }

        #endregion

        #region CreateCell
        public void CreateCellHeaderLeft(IRow row, int col, string value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = headerLeftCellStyle;
        }

        public void CreateCellHeaderRight(IRow row, int col, string value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = headerRightCellStyle;
        }

        public void CreateCellColHead(IRow row, int col, string value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = colHeadCellStyle;
        }

        public void CreateCellColLeft(IRow row, int col, string value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = colLeftCellStyle;
        }

        public void CreateCellColLeft(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = colLeftCellStyle;
        }

        public void CreateCellColRight(IRow row, int col, string value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = colRightCellStyle;
        }

        public void CreateCellColRight(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = colRightCellStyle;
        }

        public void CreateCellColCenter(IRow row, int col, string value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = colCenterCellStyle;
        }

        public void CreateCellColCenter(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = colCenterCellStyle;
        }

        public void CreateCellColNumber(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = colNumberCellStyle;
        }

        public void CreateCellCol1Decimal(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = col1DecimalCellStyle;
        }

        public void CreateCellCol2Decimal(IRow row, int col, string value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = col2DecimalCellStyle;
        }

        public void CreateCellCol2Decimal(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = col2DecimalCellStyle;
        }

        public void CreateCellCol4Decimal(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = col4DecimalCellStyle;
        }

        public void CreateCellCol6Decimal(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = col6DecimalCellStyle;
        }

        public void CreateCellFooter(IRow row, int col, string value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = footerCellStyle;
        }

        public void CreateCellFooterCenter(IRow row, int col, string value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = footerCenterCellStyle;
        }

        public void CreateCellFooter(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = footerCellStyle;
        }

        public void CreateCellFooterRight(IRow row, int col, string value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = footerRightCellStyle;
        }

        public void CreateCellFooterRight(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = footerRightCellStyle;
        }

        public void CreateCellFooterNumber(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = footerNumberCellStyle;
        }

        public void CreateCellFooter2Decimal(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = footer2DecimalCellStyle;
        }

        public void CreateCellFooter6Decimal(IRow row, int col, double value)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = footer6DecimalCellStyle;
        }

        public void CreateCellCustomStyle(IRow row, int col, string value, ICellStyle customStyle)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = customStyle;
        }

        public void CreateCellCustomStyle(IRow row, int col, double value, ICellStyle customStyle)
        {
            ICell cell = row.CreateCell(col);
            cell.SetCellValue(value);
            cell.CellStyle = customStyle;
        }

        public void CreateCellByRowNumCenterHeader(ISheet sheet, int rownum, int cellcreate, string value)
        {
            sheet.GetRow(rownum).CreateCell(cellcreate).SetCellValue(value);
            sheet.GetRow(rownum).GetCell(cellcreate).CellStyle = colHeadCellStyle;
        }

        public void CreateCellByRowNumCenter(ISheet sheet, int rownum, int cellcreate, string value)
        {
            sheet.GetRow(rownum).CreateCell(cellcreate).SetCellValue(value);
            sheet.GetRow(rownum).GetCell(cellcreate).CellStyle = colCenterCellStyle;
        }

        public void CreateCellByRowNum2Decimal(ISheet sheet, int rownum, int cellcreate, double value)
        {
            sheet.GetRow(rownum).CreateCell(cellcreate).SetCellValue(value);
            sheet.GetRow(rownum).GetCell(cellcreate).CellStyle = col2DecimalCellStyle;
        }

        #endregion

        private short SetColor(int r, int g, int b)
        {
            HSSFPalette palette = wb.GetCustomPalette();
            try
            {
                HSSFColor hssfColor = palette.FindColor((byte)r, (byte)g, (byte)b);
                if (hssfColor == null)
                {
                    hssfColor = palette.FindSimilarColor((byte)r, (byte)g, (byte)b);
                    return hssfColor.Indexed;
                }
                else
                {
                    return hssfColor.Indexed;
                }
            }
            catch { }
            return HSSFColor.COLOR_NORMAL;
        }

        public string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
