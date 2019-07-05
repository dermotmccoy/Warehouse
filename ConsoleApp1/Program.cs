using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
    class Program
    {
        public static void Main() {


            var warehouseItems = new List<warehouseTransact> { new warehouseTransact { ProductCode="70007926259"
, Location="071309-WH",PalletSize="1"},new warehouseTransact { ProductCode="7000798888"
, Location="078888-WH",PalletSize="2"} };

            openExcel();
            //FirstMacro_CodedStep();

            //displayInExcel(warehouseItems);
        }

        /// <summary>
        /// https://stackoverflow.com/questions/14248592/running-an-excel-macro-via-c-run-a-macro-from-one-workbook-on-another
        /// </summary>
        public static void openExcel()
        {
            // Object for missing (or optional) arguments.
            object oMissing = System.Reflection.Missing.Value;

            // Create an instance of Microsoft Excel
            Excel.ApplicationClass oExcel = new Excel.ApplicationClass();

            // Make it visible
            oExcel.Visible = true;

            // Open Worksheet01.xlsm
            Excel.Workbooks oBooks = oExcel.Workbooks;
            Excel._Workbook oBook = null;
            oBook = oBooks.Open("C:\\devel\\warehouse\\WarehouseDatabase.xlsm", oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);


            ((Excel.Worksheet)oExcel.ActiveWorkbook.Sheets[7]).Select();

            Excel._Worksheet workSheet = (Excel.Worksheet)oExcel.ActiveSheet;
            workSheet.Cells[1, "B"] = "70007926259";// "Product value";
            workSheet.Cells[2, "B"] = "1";//"Pallet value";
            workSheet.Cells[2, "F"] = "071309-WH";//"Location value";

            //oExcel.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, oExcel, new object[] { "C:\\devel\\warehouse\\WarehouseDatabase.xlsm!SearchIn" });
            oExcel.Run("SearchIn",false);


        }

        public static void FirstMacro_CodedStep()
        {
            // Create an instance of Microsoft Excel
            Excel.ApplicationClass oExcel = new Excel.ApplicationClass();
            Console.WriteLine("ApplicationClass: " + oExcel);

            Excel._Worksheet workSheet = (Excel.Worksheet)oExcel.ActiveSheet;
            workSheet.Cells[1, "B"] = "Product value";
            workSheet.Cells[2, "B"] = "Pallet value";
            workSheet.Cells[2, "F"] = "Location value";


            // Run the macro, "First_Macro"
            RunMacro(oExcel, new Object[] { "Worksheet01.xlsm!First_Macro" });

            //Garbage collection
            GC.Collect();
        }

        private static void RunMacro(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, oApp, oRunArgs);
        }

        /// <summary>
        /// https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/interop/how-to-access-office-onterop-objects
        /// </summary>
        /// <param name="warehouseitems"></param>
        static void displayInExcel(IEnumerable<warehouseTransact> warehouseitems) {

            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Visible = true;

            // Create a new, empty workbook and add it to the collection returned 
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template. 
            // Because no argument is sent in this example, Add creates a new workbook. 
            excelApp.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is
            // removed in a later procedure.
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            // Establish column headings in cells A1 and B1.
            workSheet.Cells[1, "A"] = "Product Code";
            workSheet.Cells[1, "B"] = "Pallet Size";
            workSheet.Cells[1, "C"] = "Location";

            var row = 1;
            foreach (var item in warehouseitems)
            {
                row++;
                workSheet.Cells[row, "A"] = item.ProductCode;
                workSheet.Cells[row, "B"] = item.PalletSize;
                workSheet.Cells[row, "C"] = item.Location;
            }

            ((Excel.Range)workSheet.Columns[1]).AutoFit();
            ((Excel.Range)workSheet.Columns[2]).AutoFit();
            ((Excel.Range)workSheet.Columns[3]).AutoFit();
        }
    }

     public class warehouseTransact
    {
        private string productCode = "";
        private string palletSize = "";
        private string location = "";

        public string ProductCode { get => productCode; set => productCode = value; }

        
        public string PalletSize { get => palletSize; set => palletSize = value; }

        public string Location { get => location; set => location = value; }


    }
}
