/* Excel Data Extractor */

// Right Click on the Solution Name -> Add -> COM Reference ->
// -> "Microsoft Excel xx.x Object Library"

using System;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelDataExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
        FileAgain:
            Console.Write("Excel Filename to Read: ");

            string basepath;
            string filepath;
            int dot;

            if (Console.ReadLine() is not string filename)
            {
                Console.WriteLine("\nInput is Null!\n");
                goto FileAgain;
            }

            if (filename.LastIndexOf('\\') == -1)           // If The Input is Only A Filename and Not A Full Path of A File
            {
                basepath = AppDomain.CurrentDomain.BaseDirectory;      // basepath is the Working Directory of This Solution
                filepath = basepath + filename;                                     // filepath = Full Path of the Word File
                dot = filename.LastIndexOf('.');      // Detect If the Filename is Valid by Checking the Extension Existence
            }
            else         // If A Full Path of the Word File is Entered or the File is Dragged over the Command Prompt Window
            {
                filepath = filename;                                             // Directly Assign filename to the filepath
                dot = filename.Substring(filename.LastIndexOf('\\')).LastIndexOf(".");    // Detect Dot Char in the Filename
            }

            Console.WriteLine(filepath);

            if (dot == -1)                          // If There is No Dot in the Filename, which Means There is No Extension
            {
                Console.WriteLine("Filename is Invalid!\n");                        // Assume That it is An Invalid Filename
                goto FileAgain;
            }

            if (!File.Exists(filepath))
            {
                Console.WriteLine("File Not Found!\n");
                goto FileAgain;
            }


            int k = filepath.LastIndexOf(".");           // After Accepting A Valid Filename with A Dot Char, Find the Last Dot

            string[] splitname = { filepath.Substring(0, k), filepath.Substring(k) };     // {basepath/onlyfilename, extension}
            if (splitname[1] != ".xls" && splitname[1] != ".xlsx")
            {
                Console.WriteLine("This is NOT an Excel File!\n");
                goto FileAgain;
            }
            string savepath = splitname[0] + ".txt";    // savepath = basepath/onlyfilename.txt (doc/docx will Be Saved to txt)


            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelBook = excelApp.Workbooks.Open(filepath);

            Console.WriteLine("Worksheets in the Excel File:\n");       // Listing the Sheets That arwe Found in the Excel File
            List<string> sheets = new List<string>();
            foreach (Excel.Worksheet item in excelBook.Worksheets)
            {
                sheets.Add(item.Name);
                Console.WriteLine(sheets.Count + ". " + item.Name);

            }
            Console.WriteLine();


            Excel.Worksheet excelSheet;

        WorkAgain:
            Console.Write("Number of Worksheet to Read: ");              // Asking User to Enter the Number of Sheet to Process
            try
            {
                excelSheet = excelBook.Worksheets[sheets.ElementAt(Convert.ToInt32(Console.ReadLine()) - 1)];
            }
            catch (Exception)
            {
                Console.WriteLine("Worksheet Not Found!\n");
                goto WorkAgain;
            }


            Excel.Range excelRange = excelSheet.UsedRange;
            object[,] excelArray = (object[,])excelRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

            Console.WriteLine();

            using (StreamWriter textFile = new StreamWriter(savepath))
            {
                for (int i = 1; i <= excelArray.GetLength(0); i++)
                {
                    for (int j = 1; j <= excelArray.GetLength(1); j++)
                    {
                        Console.Write(excelArray[i, j]);
                        textFile.Write(excelArray[i, j]);
                        if (j != excelArray.GetLength(1))
                        {
                            Console.Write("\t");
                            textFile.Write("\t");
                        }
                    }
                    Console.WriteLine();
                    textFile.WriteLine();
                }
            }


            Console.WriteLine("\n\n----------");
            Console.WriteLine("Written in: " + savepath);
            excelBook.Close(false, filepath, null);
            excelApp.Quit();

            Console.WriteLine();
            Console.Write("Press any key to continue . . . ");
            Console.ReadKey();
            return;
        }
    }
}
