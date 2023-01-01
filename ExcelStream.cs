using System;
using System.Data;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace GM.DocumentUtility
{
   public class ExcelStream 
   {
      public bool ExportExcelXLSX(ref ExcelEntity ExcelEnt, DataTable Dt_Export)  
      {
            try 
            {
                  XSSFWorkbook Xssfworkbook = new XSSFWorkbook();  
                  int iCol = 0;  
                  ISheet st = Xssfworkbook.CreateSheet(ExcelEntity.SheetName);   
                  ICellStyle csty = Xssfworkbook.CreateCellStyle();   
                  IFont f = Xssfworkbook.CreateFont();  
                  f.Boldweight = (short)FontBoldWeight.Bold;  
                  csty.Alignment = HorizontalAlignment.Center;  
                  csty.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;   
                  csty.SetFont(f);   
                  IRow r = st.CreateRow(0);     

                  foreach (DataColumn Column in Dt_Export.Columns)   
                  {
                        ICell c = r.CreateCell(iCol);  
                        c.CellStyle = csty;  
                        c.SetCellValue(Column.ColumnName);  
                        iCol = iCol + 1;   
                  }  
                  for (int j = 0; j <= iCol - 1; j++)   
                        st.AutoSizeColumn(j);  
                  for (int i = 0, i <= Dt_Export.Rows.Count - 1; i++)  
                  {
                        r = st.CreateRow(i + 1);   
                        for (int j = 0; j <= Dt_Export.Columns.Count - 1; j++)
                              r.CreateCell(j).SetCellValue(Dt_Export.Rows[i][j].ToString());
                  }

                  if (File.Exists(ExcelEnt.PathDes + @"\\" + ExcelEnt.FileName))
                        File.Delete(ExcelEnt.PathDes + @"\\" + ExcelEnt.FileName);
                  FileStream FileData = new FileStream(ExcelEnt.PathDes + @"\\" + ExcelEnt.FileName, FileMode.Create);
                  Xssfworkbook.Write(FileData);
                  FileData.Close();

                  ExcelEnt.ReturnCode = "0";
                  ExcelEnt.Msg = "Success";
            }
            catch (Exception Ex)
            {
                ExcelEnt.ReturnCode = "-999";
                ExcelEnt.Msg = Ex.Message;
                return false;
            }

            return true;
        }

        public bool ExportExcelXLS(ref ExcelEntity ExcelEnt, DataTable Dt_Export)
        {
            try
            {
                HSSFWorkbook Hssfworkbook = new HSSFWorkbook();
                int iCol = 0;
                ISheet st = Hssfworkbook.CreateSheet(ExcelEnt.SheetName);
                ICellStyle csty = Hssfworkbook.CreateCellStyle();
                IFont f = Hssfworkbook.CreateFont();
                f.Boldweight = (short)FontBoldWeight.Bold;
                csty.Alignment = HorizontalAlignment.Center;
                csty.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
                csty.FillPattern = FillPattern.SolidForeground;
                csty.SetFont(f);
                IRow r = st.CreateRow(0);
                foreach (DataColumn Column in Dt_Export.Columns)
                {
                    ICell c = r.CreateCell(iCol);
                    c.CellStyle = csty;
                    c.SetCellValue(Column.ColumnName);
                    iCol = iCol + 1;
                }
                for (int j = 0; j <= iCol - 1; j++)
                    st.AutoSizeColumn(j);
                for (int i = 0; i <= Dt_Export.Rows.Count - 1; i++)
                {
                    r = st.CreateRow(i + 1);
                    for (int j = 0; j <= Dt_Export.Columns.Count - 1; j++)
                        r.CreateCell(j).SetCellValue(Dt_Export.Rows[i][j].ToString());
                }

                if (File.Exists(ExcelEnt.PathDes + @"\\" + ExcelEnt.FileName))
                    File.Delete(ExcelEnt.PathDes + @"\\" + ExcelEnt.FileName);
                FileStream FileData = new FileStream(ExcelEnt.PathDes + @"\\" + ExcelEnt.FileName, FileMode.Create);
                Hssfworkbook.Write(FileData);
                FileData.Close();

                ExcelEnt.ReturnCode = "0";
                ExcelEnt.Msg = "Success";
            }
            catch (Exception Ex)
            {
                ExcelEnt.ReturnCode = "-999";
                ExcelEnt.Msg = Ex.Message;
                return false;
            }

            return true;
        }
    }
}