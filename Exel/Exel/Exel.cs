using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exel
{
    class Exel
    {
        static void Main(string[] args)
        {
            DataStruct data = new DataStruct();
            IOWrite write = new IOWrite(data);

            //zapisvane na danni v osnovnata table
            data.AddRow("Iliana", "Nestorova", "54");
            data.AddRow("Gabriela", "Nestorova", "28");
            data.AddRow("Angel", "Nestorov", "57");

            data.PrintTable();

            write.ExportTable();
            write.RunFile();
        }
    }
}
