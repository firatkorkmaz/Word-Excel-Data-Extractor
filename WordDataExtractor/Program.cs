/* Word Data Extractor */

// Right Click on the Solution Name -> Add -> COM Reference ->
// -> "Microsoft Word xx.x Object Library"

using System;
using Microsoft.Office.Interop.Word;


namespace WordDataExtractor
{
    class Program
    {
        static void Main()
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

            if (filename.LastIndexOf('\\') == -1)         // If the Input is Only A Filename and Not A Full Path of A File
            {
                basepath = AppDomain.CurrentDomain.BaseDirectory;    // basepath is the Working Directory of This Solution
                filepath = basepath + filename;             		 // filepath = Full Path of the Word File
                dot = filename.LastIndexOf('.');    // Detect if the Filename is Valid by Checking the Extension Existence
            }
            else       // If A Full Path of the Word File is Entered or the File is Dragged over the Command Prompt Window
            {
                filepath = filename;    			// Directly Assign filename to the filepath.
                dot = filename.Substring(filename.LastIndexOf('\\')).LastIndexOf(".");  // Detect dot char in the Filename
            }

            Console.WriteLine(filepath);

            if (dot == -1)                        // If There is No Dot in the Filename, which Means There is No Extension
            {
                Console.WriteLine("Filename is Invalid!\n");    		          // Assume That It is An Invalid Filename
                goto FileAgain;
            }

            if (!File.Exists(filepath))
            {
                Console.WriteLine("File Not Found!\n");
                goto FileAgain;
            }


            int k = filepath.LastIndexOf(".");           // After Accepting A Valid Filename with A Dot Char, Find the Last Dot

            string[] splitname = { filepath.Substring(0, k), filepath.Substring(k) };     // {basepath/onlyfilename, extension}
            if (splitname[1] != ".doc" && splitname[1] != ".docx")
            {
                Console.WriteLine("This is NOT a Word File!\n");
                goto FileAgain;
            }
            string savepath = splitname[0] + ".txt";    // savepath = basepath/onlyfilename.txt (doc/docx will Be Saved to txt)


            Application application = new Application();
            Document document = application.Documents.Open(filepath);

            Console.WriteLine();
            using (StreamWriter textFile = new StreamWriter(savepath))  // Write the Content of the Word File to A New Textfile
            {
                int count = document.Paragraphs.Count;
                for (int i = 1; i <= count; i++)                        // For Each Paragraph, Get the Text Data
                {
                    string text = document.Paragraphs[i].Range.Text;
                    if (text == "")
                        continue;
                    Console.WriteLine(text);                            // Print Each Paragraph on Screen
                    textFile.Write(text);                               // Write Each Paragraph in the onlyfilename.txt
                }
            }


            Console.WriteLine("\n\n----------");
            Console.WriteLine("Written in: " + savepath);
            application.Quit();

            Console.WriteLine();
            Console.Write("Press any key to continue . . . ");
            Console.ReadKey();
            return;
        }
    }
}