using System;
using InteropExel=Microsoft.Office.Interop.Excel;
namespace Exel
{
    public class IOWrite
    {
        private DataStruct data;
        private InteropExel.Application exel;
       public IOWrite(DataStruct ndata)
        {
            data = ndata;
        }
        public bool ExportTable()
        {
            try
            {
                //podgotovka
                exel =new  InteropExel.Application();
                if (exel == null) return false;
                exel.Visible = false;

                InteropExel.Workbook workbook = exel.Workbooks.Add();
                if (workbook == null) return false;

                InteropExel.Worksheet sheet = (InteropExel.Worksheet) workbook.Worksheets[1];
                sheet.Name = "Table 1";

                //filling of table



                //memorise and close
                workbook.SaveCopyAs(GetPath()); //memorise woorkbook
                exel.DisplayAlerts = false; //exclude all alerts of Exel
                workbook.Close();
                exel.Quit();


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


        public void addrow(DataRow nrow)
        {
            try
            {

            }catch { }
        }
    }
}
