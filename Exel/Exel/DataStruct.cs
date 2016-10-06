using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exel
{
    public   class DataStruct
    {
        public List<DataRow> table = new List<DataRow>();
        public DataStruct()
        {

        }
        public void AddRow(string fname, string lname, string age)
        {
            table.Add(new DataRow(fname, lname, age));
        }
        public void PrintTable()
        {
            try
            {
                foreach(DataRow row in table)
                {
                    Console.WriteLine(row.FirstName+" "+row.LastName+", "+row.Age);
                }

            }
            catch { }
        }
    }
    public class DataRow
    {
        private string firstName = "";
        private string lastName = "";
        private string age = "";

        public DataRow(string nfirstName, string nlastName, string nage)
        {
            firstName = nfirstName;
            lastName = nlastName;
            age = nage;
        }
        public string FirstName
        {
            set { value = firstName; }
            get { return firstName; }
        }
        public string LastName
        {
            set { value = lastName; }
            get { return lastName; }
        }
        public string Age
        {
            set { value = age; }
            get { return age; }
        }
    }
}
