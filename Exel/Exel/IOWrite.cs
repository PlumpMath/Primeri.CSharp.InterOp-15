using System;
using InteropExel=Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
namespace Exel
{
    public class IOWrite
    {
        private DataStruct data;
        private InteropExel.Application excel;
       public IOWrite(DataStruct ndata)
        {
            data = ndata;
        }
        public bool ExportTable()
        {
            try
            {
                //podgotovka
                excel =new  InteropExel.Application();
                if (excel == null) return false;
                excel.Visible = false;

                InteropExel.Workbook workbook = excel.Workbooks.Add();
                if (workbook == null) return false;

                InteropExel.Worksheet sheet = (InteropExel.Worksheet) workbook.Worksheets[1];
                sheet.Name = "Table 1";

                //filling of table
                int i = 1;
                addrow(new DataRow("First Name","Last Name","Age"),i++);
                foreach(DataRow row in data.table)
                {
                    addrow(row, i++);
                }


                //memorise and close
                workbook.SaveCopyAs(GetPath()); //memorise woorkbook
                excel.DisplayAlerts = false; //exclude all alerts of Exel
                workbook.Close();
                excel.Quit();

                //Clear memory from excel !!!needed - using System.Runtime.InteropServices;
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (excel != null) Marshal.ReleaseComObject(excel);
                workbook = null;
                sheet = null;
                excel = null;
                GC.Collect();

                return true;
            } catch { }
            return false;
        }


        private string GetPath()
        {
          return System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "table.xlsx");
        }


        public void RunFile()
        {
            try
            {

                System.Diagnostics.Process.Start(GetPath());
            } catch { }
        }


        public void addrow(DataRow ndataRow, int nindexRow)
        {
            try
            {//zapis na 1 red
                InteropExel.Range range;
                range = excel.Range["A" + nindexRow.ToString(), "A" + nindexRow.ToString()];
                range.Value2 = ndataRow.FirstName;

                range = excel.Range["B" + nindexRow.ToString(), "B" + nindexRow.ToString()];
                range.Value2 = ndataRow.LastName;

                range = excel.Range["C" + nindexRow.ToString(), "C" + nindexRow.ToString()];
                range.Value2 = ndataRow.Age;
            }
            catch { }
        }
    }
}
