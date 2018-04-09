using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

using DocumentFormat.OpenXml.Spreadsheet;

namespace excelUpload
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Code Start

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(@"D:\EmployeeTemplate.xlsm", false))

            {

                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                WorksheetPart worksheetPart = workbookPart.WorksheetParts.Last();

                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().Last();

                Worksheet worksheet = worksheetPart.Worksheet;

                uint i = 0;

                int j = 1;

                StringBuilder SCTXml = new StringBuilder();



                SCTXml.Append("<ROOT>");



                int Ln_Count = 0;



                for (i = 2; i <= sheetData.Elements<Row>().Count(); i++)

                {

                    j = j + 1;



                    Cell cell = GetCell(worksheet, "A", i);

                    string colVal = "";



                    colVal = getValue(cell, spreadsheetDocument.WorkbookPart, "text");



                    if (colVal == "")

                    { break; }

                    //M_Employee_UI EMP = new M_Employee_UI();

                    //Common objCommon = new Common();

                    //EMP = objCommon.ValidateUsersFromAD(colVal, System.Web.Configuration.WebConfigurationManager.AppSettings["DomainAD"].ToString());

                    SCTXml.Append("<tbl_Employee");

                    //if (EMP.Eid != null)

                    //{

                    //    SCTXml.Append(" Empid ='" + EMP.Eid + "' ");

                    //    SCTXml.Append(" First_Nm ='" + EMP.First_Nm + "' ");

                    //    SCTXml.Append(" Last_Nm ='" + EMP.Last_Nm + "' ");



                        cell = GetCell(worksheet, "B", i);



                        SCTXml.Append(" RoleNm ='" + getValue(cell, spreadsheetDocument.WorkbookPart, "text") + "' ");



                        cell = GetCell(worksheet, "C", i);



                        SCTXml.Append(" Location_nm ='" + getValue(cell, spreadsheetDocument.WorkbookPart, "text") + "' ");



                        cell = GetCell(worksheet, "D", i);



                        SCTXml.Append(" Client_nm ='" + getValue(cell, spreadsheetDocument.WorkbookPart, "text") + "' ");



                        cell = GetCell(worksheet, "E", i);



                        SCTXml.Append(" Product_nm ='" + getValue(cell, spreadsheetDocument.WorkbookPart, "text") + "' ");



                        cell = GetCell(worksheet, "F", i);



                        SCTXml.Append(" SubProduct_nm ='" + getValue(cell, spreadsheetDocument.WorkbookPart, "text") + "' ");

                    //}

                    SCTXml.Append(" />");

                    Ln_Count = Ln_Count + 1;

                }

                SCTXml.Append("</ROOT>");

                string strXML = SCTXml.ToString();

                //msg = objDA.InsertBulkEmp(strXML.ToString(), Eid);

                //return msg;

            }

        }

        #endregion
    

    private static string getValue(Cell cell, WorkbookPart workbookPart, string type)
        {
            if (cell != null)
            {
                var cellValue = cell.CellValue;
                var text = (cellValue == null) ? cell.InnerText : cellValue.Text;
                if ((cell.DataType != null) && (cell.DataType == "s"))
                {
                    text = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(
                            Convert.ToInt32(cell.CellValue.Text)).InnerText;
                }
                return text;
            }
            else {
                return string.Empty;
            }
        }

        private static Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>().Where(c => string.Compare
                   (c.CellReference.Value, columnName +
                   rowIndex, true) == 0).FirstOrDefault();
        }

        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
    }

}
