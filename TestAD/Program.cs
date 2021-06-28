using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AD_Employee;

namespace TestAD
{
    class Program
    {
        static void Main(string[] args)
        {
            AD_Employee.Employee e = new Employee();
            DateTime dt = Convert.ToDateTime("5-31-21");
            e.CreateNewUser("hgr12345", "test guy jr",  "test.guy@rangerenergy.com", "804-Midland","13-Wireline Production","x0001","test title",8,"Hourly",5,dt,"smith,john-hgr012345","77-Operations");

        }
    }
}
