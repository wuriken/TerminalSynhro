using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TerminalSynhro;

namespace TestSynchro
{
    class Program
    {
        static void Main(string[] args)
        {


            //MConvert.ConvertProductsToCsvFile(@"C:\Users\user\Desktop\Projects\DatalogicScorpio\Exchange\OUT\Products.xlsx");
            //MConvert.ConvertToExcelDocuments(
            //    @"C:\Users\user\Desktop\Projects\DatalogicScorpio\Exchange\IN\Arrival_10_14.csv");
            //MConvert.ConvertToExcelDocuments(
            //    @"C:\Users\user\Desktop\Projects\DatalogicScorpio\Exchange\IN\Inventory_12_23.csv");
            //MConvert.ConvertToExcelDocuments(
            //    @"C:\Users\user\Desktop\Projects\DatalogicScorpio\Exchange\IN\Production_12_23.csv");
            //MConvert.CopyFromDevice();
            MConvert.GetInvoicesFromTerminal();
            //MConvert.LoadDataToTerminal();
           // bool ccc = MConvert.CheckTerminal();
            Console.WriteLine(); 
            Console.ReadKey();
        }
    }
}
