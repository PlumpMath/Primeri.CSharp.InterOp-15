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
                //mejdinni proverki


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
