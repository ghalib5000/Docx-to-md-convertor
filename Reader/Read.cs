using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Threading.Tasks;
using System.Data;
using System.Text.RegularExpressions;

namespace Reader
{
    class Read : IRead
    {

        private static string FileLocation = @"D:\convertor\test - Copy\word\document.xml";
        private static string FileLocation2 = @"D:\convertor\test.docx";
        private static string write_loc = @"D:\";
        private static string totalRawText = "";
        protected static string[] fintext;
        private string file_docx = "";
        private string file_md = "";
        private string response = "";


        /// <summary>
        /// reads the contents of the file
        /// </summary>
        public string File()
        {
            try
            {

                Console.Write("enter the location of the file: ");
                FileLocation = Console.ReadLine();
                string t = Path.GetExtension(FileLocation);
                Console.WriteLine(t);
                if (file_docx == "docx" || file_md == "md" || file_docx == ".xml")
                {
                }
                else
                {
                    throw new Exception("Wrong type!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return FileLocation;
        }
        /// <summary>
        /// writes the contents of the file to a new location
        /// </summary>
        public string Write()
        {
            try
            {
                string[] s = System.IO.File.ReadAllLines(FileLocation);
                if (this.file_md == "md")
                {
                    string response = "";
                    Console.Write("Convert to .docx file?: ");
                    response = Console.ReadLine();
                    if (response == "yes" || response == "y" || response == "Y")
                    {
                        Console.Write("enter the location for saving the file:");
                        write_loc = Console.ReadLine();
                        if (!Directory.Exists(write_loc))
                        {
                            Console.Write("Directory does not exsits....create one?: ");
                            response = Console.ReadLine();
                            if (response == "yes" || response == "y" || response == "Y")
                            {
                                Directory.CreateDirectory(write_loc);
                            }
                        }
                        write_loc += "out.docx";
                        //File.WriteAllLines(write_loc, s);
                        /* for(int i=0;i<s.Length;i++)
                         {
                           Console.WriteLine((i / s.Length) * 100);
                         }*/
                        //   Convert(FileLocation, write_loc, WdSaveFormat.wdFormatXMLDocument);
                    }
                    else
                    {
                        Console.WriteLine("Exiting...");
                    }
                }
                else if (this.file_docx == "docx" || this.file_docx == ".xml")
                {

                    Console.Write("Convert to .md file?: ");
                    response = Console.ReadLine();
                    if (response == "yes" || response == "y" || response == "Y")
                    {
                        Console.Write("enter the location for saving the file:");
                        write_loc = Console.ReadLine();
                        if (!Directory.Exists(write_loc))
                        {
                            Console.Write("Directory does not exsits....create one?: ");
                            response = Console.ReadLine();
                            if (response == "yes" || response == "y" || response == "Y")
                            {
                                Directory.CreateDirectory(write_loc);
                            }
                        }
                        write_loc += "out.md";
                        /* for (float i = 0; i < s.Length; i++)
                          {
                              Console.WriteLine((i / s.Length) * 100);
                           File.WriteAllLines (write_loc, s);
                          }*/
                        //  Convert(FileLocation, write_loc, WdSaveFormat.wdFormatRTF);//wdFormatXML);// WdSaveFormat.wdFormatUnicodeText);
                    }
                    else
                    {
                        Console.WriteLine("Exiting...");
                    }
                }
            }
            catch (DirectoryNotFoundException dirntfnd)
            {
                Console.WriteLine(dirntfnd);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);

            }
            return write_loc;
        }

        public void func()
        {
            XmlTextReader textReader = new XmlTextReader(FileLocation);
            textReader.Read();
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(textReader);
            try
            {
                //the raw text is extracted from the xml onto the string
                totalRawText = xdoc.DocumentElement.InnerXml;
                //                Console.WriteLine(totalRawText);
                recursivefunc(totalRawText, "</w:p>");
               /* foreach (string t in fintext)
                {
                    if (t.Contains("<w:b />"))
                    {
                        recursivefunc(t, "<w:b />");
                    }
                    else if (t.Contains("<w:i />"))
                    {
                        recursivefunc(t, "<w:i />");
                    }
                }
                */
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        public static void recursivefunc(string text, string pattern)
        {

            fintext = totalRawText.Split(pattern);
            //iterating through each index of the total text lines
            foreach (string t in fintext)
            {
               
                //    System.IO.File.AppendAllText(write_loc + "output.md", t + Environment.NewLine);
                Console.WriteLine(t);
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine();

            }
        }







        // Convert a Word 2008 .docx to Word 2003 .doc
        public void Convert(string input, string output)
        {
            WdSaveFormat format = WdSaveFormat.wdFormatUnicodeText;
            // Create an instance of Word.exe
            Word._Application oWord = new Word.Application();

            if (this.file_md == "md")
            {
                format = WdSaveFormat.wdFormatStrictOpenXMLDocument;
            }
            else if (this.file_docx == "docx" || this.file_docx == ".xml")
            {
                format = WdSaveFormat.wdFormatUnicodeText;
            }

            // Interop requires objects.
            object oMissing = System.Reflection.Missing.Value;
            object isVisible = true;
            object readOnly = false;
            object oInput = input;
            object oOutput = output;
            object oFormat = format;

            try
            {

                // Make this instance of word invisible (Can still see it in the taskmgr).
                oWord.Visible = false;

                // Load a document into our instance of word.exe
                Word._Document oDoc = oWord.Documents.Open(ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                // Make this document the active document.
                oDoc.Activate();

                // Save this document in Word 2003 format.
                oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                // Always close Word.exe.
                oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
            }
        }








































        /* void IRead.Convert(string input, string output)
         {
             WdSaveFormat format = WdSaveFormat.wdFormatUnicodeText;
             string[] s = File.ReadAllLines(FileLocation);
             // Create an instance of Word.exe
             Word._Application oWord = new Word.Application();

             // Make this instance of word invisible (Can still see it in the taskmgr).
             oWord.Visible = false;
             // Interop requires objects.
             object oMissing = System.Reflection.Missing.Value;
             object isVisible = true;
             object readOnly = false;
             object oInput = input;
             object oOutput = output;
             object oFormat = format;
             try
             {
                 if (file_docx == "docx")
                 {
                     format = WdSaveFormat.wdFormatUnicodeText;
                 }
                 if (file_md == "md")
                 {
                     format = WdSaveFormat.wdFormatDocument;
                 }

                 // Load a document into our instance of word.exe
                 Word._Document oDoc = oWord.Documents.Open(ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                 // Make this document the active document.
                 oDoc.Activate();

                 // Save this document in Word 2003 format.
                 oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                 // Always close Word.exe.
                 oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
             }
             catch (Exception ex)
             {
                 Console.WriteLine(ex);
                 oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
             }
         }*/
    }
}

