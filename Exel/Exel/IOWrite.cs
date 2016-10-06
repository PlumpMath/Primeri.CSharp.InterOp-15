using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exel
{
    public class IOWrite
    {
        private DataStruct data;
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
