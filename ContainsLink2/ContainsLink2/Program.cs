using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace ContainsLink
{
    class Program
    {
        static void Main(string[] args)
        {
            //get user's desktop path for log file, and add log file name
            string logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\containsLink_log.txt";

            //retrive file names from user's specified directory and sub directories
            Console.WriteLine("Enter directory to search:  ");
            string path = Console.ReadLine();
            string[] myFiles = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories);

            //process each file
            foreach (string filename in myFiles)
            {
                //skip file if just temp file
                if (filename.Substring(0, 1) != "~")
                {
                    //write filename to log file
                    System.IO.File.AppendAllText(logFilePath, filename);
                    System.IO.File.AppendAllText(logFilePath, Environment.NewLine);
                    Console.WriteLine(filename);

                    //make list to store file's links
                    List<string> links = new List<string>();

                    //find file's extension (type)
                    int extLoc = filename.LastIndexOf(".");
                    int endLoc = filename.Length - extLoc - 1;
                    string fileExt = filename.Substring(extLoc + 1, endLoc);

                    //process file according to its type, extracting the file's links
                    if (fileExt == "pdf")
                    {
                        //Setup some variables to be used later
                        PdfReader reader = default(PdfReader);
                        int pageCount = 0;
                        PdfDictionary pageDictionary = default(PdfDictionary);
                        PdfArray annots = default(PdfArray);

                        //Open our reader
                        reader = new PdfReader(filename);
                        //Get the page cont
                        pageCount = reader.NumberOfPages;

                        //Loop through each page
                        for (int i = 1; i <= pageCount; i++)
                        {
                            //Get the current page
                            pageDictionary = reader.GetPageN(i);
                            //Get all of the annotations for the current page
                            annots = pageDictionary.GetAsArray(PdfName.ANNOTS);
                            //Make sure we have something
                            if ((annots == null) || (annots.Length == 0))
                            {
                                Console.WriteLine("nothing");
                            }
                            //Loop through each annotation
                            if (annots != null)
                            {
                                //add page number to list
                                links.Add("Page " + i);

                                foreach (PdfObject A in annots.ArrayList)
                                {
                                    //Convert the itext-specific object as a generic PDF object
                                    PdfDictionary AnnotationDictionary =
                                        (PdfDictionary)PdfReader.GetPdfObject(A);
                                    //Make sure this annotation has a link
                                    if (!AnnotationDictionary.Get(PdfName.SUBTYPE).Equals(PdfName.LINK))
                                        continue;
                                    //Make sure this annotation has an ACTION
                                    if (AnnotationDictionary.Get(PdfName.A) == null)
                                        continue;
                                    //Get the ACTION for the current annotation
                                    PdfDictionary AnnotationAction =
                                        AnnotationDictionary.GetAsDict(PdfName.A);
                                    // Test if it is a URI action
                                    if (AnnotationAction.Get(PdfName.S).Equals(PdfName.URI))
                                    {
                                        PdfString Destination = AnnotationAction.GetAsString(PdfName.URI);
                                        string url = Destination.ToString();

                                        //add url to list
                                        links.Add(url);
                                    }
                                }
                            }
                        }
                    }
                    else if ((fileExt == "xls") || (fileExt == "xlsx"))
                    {
                        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                        excelApp.Visible = false;

                        //open excel file
                        Workbook excelWorkbook = excelApp.Workbooks.Open(filename, 0, true, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                        Sheets excelSheets = excelWorkbook.Worksheets;

                        //loop through worksheets
                        for (int i = 1; i <= excelSheets.Count; i++)
                        {
                            Worksheet sheet = (Worksheet)excelSheets.Item[i];

                            //extract links
                            if (sheet.Hyperlinks.Count > 0)
                            {
                                links.Add("Page " + i);

                                for (int j = 1; j <= sheet.Hyperlinks.Count; j++)
                                {
                                    string address = sheet.Hyperlinks[i].Address;
                                    links.Add(address);
                                }
                            }
                        }
                    }
                    else if ((fileExt == "doc") || (fileExt == "docx"))
                    {
                        Microsoft.Office.Interop.Word.Application applicationObject = new Microsoft.Office.Interop.Word.Application();
                        Document aDDoc = applicationObject.Documents.Open(FileName: filename, ReadOnly: true);
                        Microsoft.Office.Interop.Word.Hyperlinks hyperlinks = aDDoc.Hyperlinks;

                        if (hyperlinks.Count > 0)
                        {
                            //links.Add("Page: ");

                            foreach (var hyperlink in hyperlinks)
                            {
                                string address = ((Microsoft.Office.Interop.Word.Hyperlink)hyperlink).Address;
                                links.Add(address);
                            }
                        }
                    }
                    else if ((fileExt == "csv") || (fileExt == "txt"))
                    {
                        string[] lines = System.IO.File.ReadAllLines(filename);

                        foreach (string line in lines)
                        {
                            foreach (Match item in Regex.Matches(line, @"(http(s)?://)?([\w\-_]+(?:(?:\.[\w\-_]+)+))([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?"))
                            {
                                string address = item.Value;
                                links.Add(address);
                            }
                        }
                    }
                    else
                    {

                    }

                    //write the file's link count to the log file
                    //have to subtract the page number markers from the links count
                    int fileLinkCount = 0;
                    foreach (string line in links)
                    {
                        if (line.Substring(0, 4) != "Page")
                        {
                            fileLinkCount++;
                        }
                    }
                    System.IO.File.AppendAllText(logFilePath, "Link count: " + fileLinkCount + Environment.NewLine);
                    
                    //write the file's links to the log file
                    if (links.Count > 0)
                    {
                        using (TextWriter tw = new StreamWriter(logFilePath, append: true))
                        {
                            foreach (String s in links)
                                tw.WriteLine(s);
                        }
                    }

                    System.IO.File.AppendAllText(logFilePath, "----------" + Environment.NewLine);
                }
            }
        }
    }
}
