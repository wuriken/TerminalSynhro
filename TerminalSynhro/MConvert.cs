using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using OpenNETCF.Desktop.Communication;

namespace TerminalSynhro
{
    public static class MConvert
    {
        private static string PathToExchangeFolder = @"C:\Users\user\Desktop\Projects\DatalogicScorpio\Exchange\";
        private static string PathToInvoiceTerminalFolder = @"\Program Files\DatalogicScorpio\Invoices\";
        public static string PathToRootTerminalFolder = @"\Program Files\DatalogicScorpio\";

        public static void ItalicBoldFontEnable(Worksheet sheet, string cellFrom, string cellTo)
        {
            sheet.Range[cellFrom, cellTo].Font.Italic = true;
            sheet.Range[cellFrom, cellTo].Font.Bold = true;
        }

        public static object CellsValueGet(Worksheet sheet, string cellFrom, string cellTo)
        {
            bool stop = false;
            while (!stop)
            {
                try
                {
                    return sheet.Range[cellFrom, cellTo].Cells.Value ?? string.Empty;
                }
                catch (Exception)
                {
                    // ignored
                }
            }
            return string.Empty;
        }

        public static void CellsValuesSet(Worksheet sheet, string cellFrom, string cellTo, object value)
        {
            bool stop = false;
            while (!stop)
            {
                try
                {
                    sheet.Range[cellFrom, cellTo].Cells.Value = value;
                    stop = true;
                }
                catch (Exception)
                {
                    //ignores
                }
            }
        }

        public static DirectoryInfo ExchangeDirectoryInfo()
        {
            return  new DirectoryInfo(PathToExchangeFolder);
        }

        public static DirectoryInfo TerminalDirectoryInvoiceInfo()
        {
            return new DirectoryInfo(PathToInvoiceTerminalFolder);
        }

        public static DirectoryInfo TerminalRootFolderInfo()
        {
            return new DirectoryInfo(PathToRootTerminalFolder);
        }

        public static bool GetInvoicesFromTerminal()
        {
            CopyFromDevice();
            if (!ConvertInInvoicesToXlsx(PathToExchangeFolder + @"IN\")) return false;
            return true; //DirectoriesCopy();
        }

        public static void CopyToDevice(string pathFolder, string pathDevice)
        {
            RAPI rap = new RAPI();
            rap.Connect();
            if (rap.Connected)
            {
                rap.CopyFileToDevice(pathFolder, pathDevice, true);
            }
            rap.Disconnect();
        }

        public static void CopyFromDevice()
        {
            RAPI rap = new RAPI();
            rap.Connect();
            if (rap.Connected)
            {
                IEnumerable<FileInformation> inf = rap.EnumerateFiles(PathToInvoiceTerminalFolder + "*");
                foreach (var item in inf)
                {
                    if (item.FileName.Contains(".csv")) continue;
                    IEnumerable<FileInformation> informations =
                        rap.EnumerateFiles(PathToInvoiceTerminalFolder + @"\" + item.FileName + @"\*");
                    foreach (var file in informations)
                    {
                        rap.CopyFileFromDevice(PathToExchangeFolder + @"IN\" + file.FileName, 
                            PathToInvoiceTerminalFolder + @"\" + item.FileName + @"\" + file.FileName);
                    }
                }
               
            }
            rap.Disconnect();
        }

        private static bool EnterDataInTerminalDelete()
        {
            try
            {
                FileInfo[] fileInfos = ExchangeDirectoryInfo().GetFiles();
                foreach (FileInfo item in fileInfos)
                {
                    File.Delete(item.FullName);
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLineToFile("EnterDataInTerminalDelete" + ex.Message, @"C:\Users\Public\dll.log");
                return false;
            }
        }

        public static bool LoadDataToTerminal()
        {
            if (!EnterDataInTerminalDelete())
                return false;
            if (ConvertProductsToCsvFile(PathToExchangeFolder + @"OUT\" + "Products.xlsx") &
                ConvertDataToCsvFile(new FileInfo(PathToExchangeFolder + @"OUT\" + "Contractors.xlsx")) &
                ConvertDataToCsvFile(new FileInfo(PathToExchangeFolder + @"OUT\" + "ProductsGroup.xlsx")) &
                ConvertDataToCsvFile(new FileInfo(PathToExchangeFolder + @"OUT\" + "ProductsType.xlsx")) &
                ConvertDataToCsvFile(new FileInfo(PathToExchangeFolder + @"OUT\" + "Storage.xlsx")))
            {
                FileInfo[] fileinF = ExchangeDirectoryInfo().GetFiles();
                foreach (FileInfo item in fileinF)
                {
                    CopyToDevice(item.FullName, PathToInvoiceTerminalFolder + item.Name);
                }
            }
            return true;

        }

        private static bool ConvertInInvoicesToXlsx(string path)
        {
            try
            {
                FileInfo[] filesInfo = new DirectoryInfo(path).GetFiles();
                foreach (var item in filesInfo)
                {
                    if (item.Name.Contains(".csv"))
                    {
                        ConvertToExcelDocuments(item);
                        item.Delete();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLineToFile("ConvertToXlSx" + ex.Message, @"C:\Users\Public\dll.log");
                return false;
            }  
        }

        public static bool ConvertToExcelDocuments(FileInfo file)
        {
            try
            {
                Application excelApp = new Application
                {
                    Visible = false,
                    SheetsInNewWorkbook = 1
                };
                excelApp.Workbooks.Add(Type.Missing);
                excelApp.Worksheets.Add();
                Workbook workBook = excelApp.Workbooks[1];
                Worksheet excelWorkSheet = (Worksheet)excelApp.Worksheets.Item[1];
                List<string[]> list = InvoiceCsvFileRead(file.FullName);
                int cellNum = 1;
                foreach (string[] item in list)
                {
                    CellsValuesSet(excelWorkSheet, "A" + cellNum, "M" + cellNum, item);
                    cellNum++;
                }

                excelApp.DefaultSaveFormat = XlFileFormat.xlExcel12;
                workBook.Saved = false;
                workBook.SaveAs(PathToExchangeFolder + @"IN\" + file.Name.Split
                    (new [] {".csv"}, StringSplitOptions.None)[0] + ".xlsx", Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                WriteLineToFile("File " + PathToExchangeFolder + @"IN\" + file.Name.Split(new [] {".csv"}, StringSplitOptions.None)[0] 
                    + " converted and copied.", @"C:\Users\Public\dll.log");
                excelApp.Quit();
                return true;
            }
            catch (Exception ex)
            {
                WriteLineToFile(ex.Message, @"C:\Users\Public\dll.log");
                return false;
            }
        }

        public static bool ConvertProductsToCsvFile(string path)
        {
            Application excelApp = new Application { Visible = false };
            try
            {
                excelApp.Workbooks.Open(path);
                Worksheet excelWorkSheet = (Worksheet)excelApp.Worksheets.Item[1];
                bool stop = false;
                int cellNumber = 1;
                List<string> resultList = new List<string>();
                while (!stop)
                {
                    object[,] res = (object[,])CellsValueGet(excelWorkSheet, "A" + cellNumber, "K" + cellNumber);
                    var tempStr = string.Empty;
                    for (int i = 1; i < res.Length; i++)
                    {
                        if (res[1, i] == null) res[1, i] = string.Empty;
                    }
                    if (res[1, 1].ToString() == string.Empty)
                    {
                        stop = true;
                    }
                    else
                    {
                        for (int i = 1; i <= res.GetLength(1); i++)
                        {
                            if (i == res.GetLength(1))
                            {
                                tempStr += res[1, i].ToString();
                            }
                            else
                            {
                                tempStr += res[1, i] + ";";
                            }
                        }
                        resultList.Add(tempStr);
                        cellNumber++;
                    }
                }
                excelApp.Quit();
                foreach (string item in resultList)
                {
                    WriteLineToFile(item, PathToExchangeFolder + "Products.csv");
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLineToFile(ex.Message, @"C:\Users\Public\dll.log");
                return false;
            }
        }

        public static bool ConvertDataToCsvFile(FileInfo fileInf)
        {
            Application excelApp = new Application { Visible = false };
            try
            {
                excelApp.Workbooks.Open(fileInf.FullName);
                Worksheet excelWorkSheet = (Worksheet)excelApp.Worksheets.Item[1];
                bool stop = false;
                int cellNumber = 1;
                List<string> resultList = new List<string>();
                while (!stop)
                {
                    object res = CellsValueGet(excelWorkSheet, "A" + cellNumber, "A" + cellNumber) ?? string.Empty;
                    if (res.ToString() == string.Empty)
                    {
                        stop = true;
                    }
                    else
                    {
                        resultList.Add(res.ToString());
                        cellNumber++;
                    }
                }
                excelApp.Quit();
                foreach (string item in resultList)
                {
                    WriteLineToFile(item, PathToExchangeFolder + 
                        fileInf.Name.Split(new []{".xls"}, StringSplitOptions.None)[0]  + ".csv");
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLineToFile(ex.Message, @"C:\Users\Public\dll.log");
                return false;
            }
        }



        public static List<string> OpenFileWithConfig()
        {
            List<string> resultList = new List<string>();
            try
            {
                StreamReader stream = new StreamReader(@"C:\Users\Public\synchropath.ini", Encoding.GetEncoding(1251));
                string result;
                while ((result = stream.ReadLine()) != null)
                {
                    WriteLineToFile(result, @"C:\Users\Public\dll.log");
                   resultList.Add(result);
                }
                stream.Close();
                stream.Dispose();
            }
            catch (Exception ex)
            {
                WriteLineToFile(ex.Message, @"C:\Users\Public\dll.log");
                return new List<string>();
            }
            return resultList;
        }


        public static List<string[]> InvoiceCsvFileRead(string path)
        {
            List<string[]> resultList = new List<string[]>();
            try
            {
                StreamReader stream = new StreamReader(path, Encoding.GetEncoding(1251));
                string result;
                while ((result = stream.ReadLine()) != null)
                {
                    string[] tempArr = result.Split(';');
                    if (tempArr.Length == 13)
                    {
                        resultList.Add(tempArr);
                    }
                }
                stream.Close();
                stream.Dispose();
            }
            catch (Exception ex)
            {
                WriteLineToFile(ex.Message, @"C:\Users\Public\dll.log");
                return new List<string[]>();
            }
            return resultList;

        }

        public static void WriteLineToFile(string product, string fileName)
        {
            try
            {
                FileStream fs = new FileStream(fileName, FileMode.Append);
                StreamWriter sw = new StreamWriter(fs, Encoding.GetEncoding(1251));
                sw.WriteLine(product);
                sw.Flush();
                sw.Close();
                fs.Close();
            }
            catch (Exception ex)
            {
                WriteLineToFile(ex.Message, @"C:\Users\Public\dll.log");
            }
        }

        public static bool DirectoriesCopy()
        {
            if (TerminalDirectoryInvoiceInfo() != null)
            {
                string path = TerminalDirectoryInvoiceInfo().Parent.FullName + @"\Archives";
                if (!Directory.Exists(path))
                {
                    try
                    {
                        Directory.CreateDirectory(path);
                    }
                    catch (Exception ex)
                    {
                        WriteLineToFile(ex.Message, @"C:\Users\Public\dll.log");
                        return false;
                    }
                }
                try
                {
                    DirectoryInfo[] dirs = TerminalDirectoryInvoiceInfo().GetDirectories();
                    foreach (DirectoryInfo item in dirs)
                    {
                        Directory.CreateDirectory(PathToRootTerminalFolder + @"\Archives\" + item.Name);
                        FileInfo[] fileInf = item.GetFiles();
                        foreach (FileInfo file in fileInf)
                        {
                            File.Move(PathToInvoiceTerminalFolder + item.Name + @"\" + file.Name,
                                PathToRootTerminalFolder + @"\Archives\" + item.Name + @"\" + file.Name);
                            WriteLineToFile("File: " + item.Name + @"\" + file.Name + " moved.", @"C:\Users\Public\dll.log");
                        }
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    WriteLineToFile(ex.Message, @"C:\Users\Public\dll.log");
                    return false;
                }
            }
            return false;

        }

    }
}
