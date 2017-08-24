using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DynamicExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Type excelType = Type.GetTypeFromProgID("Excel.Application", true);
            dynamic excel = Activator.CreateInstance(excelType);
            excel.Workbooks.Add();

            dynamic defaultWorksheet = excel.ActiveSheet;

            defaultWorksheet.Cells[1, "A"] = "This is the column name";
            defaultWorksheet.Columns[1].AutoFit();
            defaultWorksheet.Name = "Test";

            excel.Visible = true;

        }
    }
}
