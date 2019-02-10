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
        private static string[] LinedText;
        private static string[] WordedText;
        private static string[] finalText;
        private static string bold_parameters   = "<w:b /></w:rPr><w:t";
        private static string italic_parameters ="<w:i /></w:rPr><w:t";
        private static string bold_italic = "<w:b /><w:i /></w:rPr><w:t";
        private static string italic_bold = "<w:i /><w:b /></w:rPr><w:t";
        int bold = 0, italic = 0;
        private string response = "";


        /// <summary>
        /// reads the contents of the file
        /// </summary>
        public string File()
        {
            try
            {
             //   Console.Write("enter the location of the file: ");
              //  FileLocation = Console.ReadLine();
                string t = Path.GetExtension(FileLocation);
                // Console.WriteLine(t);
                Console.WriteLine();
                if (t == ".docx" || t == ".xml")
                {
                   // createWriteLocation(FileLocation);
                    Reader(FileLocation);
                    //Writer(LinedText, write_loc);
                    //bold italic is 1,italic bold is 2,bold is 3,italic is 4 
                    styleChecker(LinedText,bold_italic,italic_bold,bold_parameters, italic_parameters);
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
        public void styleChecker(string[] text, string st1, string st2, string st3, string st4)
        {
            int i = 0;
            foreach(string line in text)
            {
                lineSplitter(line);
                foreach (string words in WordedText)
                {
                    //starters
                    {
                        //bold italic starter
                        if (words.Contains(st1) && bold == 0 && italic == 0)
                        {
                            finalText[i] += "**_" + words;
                            bold = 1;
                            italic = 1;
                        }
                        //italic bold starter
                        else if (words.Contains(st2) && bold == 0 && italic == 0)
                        {
                            finalText[i] +=  "_**" + words;
                            bold = 1;
                            italic = 1;
                        }
                        //bold starter
                        else if (words.Contains(st3) && bold == 0)
                        {
                            finalText[i] += "**" + words;
                            bold = 1;
                        }
                        //italic starter
                        else if (words.Contains(st4) && italic == 0)
                        {
                            finalText[i] += "_" + words;
                            italic = 1;
                        }
                    }
                    //when style is started
                    {
                        //bold italic 
                        if (words.Contains(st1) && bold == 0 && italic == 1)
                        {
                            finalText[i] += "**" + words;
                            bold = 1;
                        }
                        //italic bold
                        else if (words.Contains(st2) && bold == 0 && italic == 1)
                        {
                            finalText[i] += "**" + words;
                            bold = 1;
                        }
                        //bold italic
                        else if (words.Contains(st1) && bold == 1 && italic == 0)
                        {
                            finalText[i] += "_" + words;
                            italic = 1;
                        }
                        //italic bold
                        else if (words.Contains(st2) && bold == 1 && italic == 0)
                        {
                            finalText[i] += "_" + words;
                            italic = 1;
                        }
                        //bold 
                        else if (words.Contains(st3) && bold == 1)
                        {
                            finalText[i] += words;
                        }
                        //italic
                        else if (words.Contains(st4) && italic == 1)
                        {
                            finalText[i] += words;
                        }
                    }
                }
                //Writer(WordedText, write_loc);
                i++;
            }
        }
        public void lineSplitter(string lines)
        {
                WordedText = lines.Split("</w:t></w:r>");
        }

        /// <summary>
        /// Creates a new location to write the contents of file to
        /// </summary>
        public string createWriteLocation(string FileLocation)
        {
            try
            {
                Console.Write("Convert to .md file?: ");
                response = Console.ReadLine();
                response = response.ToLower();
                if (response[0] == 'y')
                {
                    Console.Write("enter the location for saving the file:");
                    write_loc = Console.ReadLine();
                    if (!Directory.Exists(write_loc))
                    {
                        Console.Write("Directory does not exsits....create one?: ");
                        response = Console.ReadLine();
                        response = response.ToLower();
                        if (response[0] =='y')
                        {
                            Directory.CreateDirectory(write_loc);
                        }
                    }
                    write_loc += "out.md";
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

        /// <summary>
        /// Reads and splits the xml text into each seperate line
        /// </summary>
        public void Reader(string fileloc)
        {
            XmlTextReader textReader = new XmlTextReader(fileloc);
            textReader.Read();
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(textReader);
            try
            {
                //the raw text is extracted from the xml onto the string
                totalRawText = xdoc.DocumentElement.InnerXml;
                //                Console.WriteLine(totalRawText);
                Splitter(totalRawText, "</w:p>");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        /// <summary>
        /// splits the xml text with the given pattern
        /// </summary>
        /// <param name="text"></param>
        /// <param name="pattern"></param>
        public static void Splitter(string text, string pattern)
        {
            //spliting each line of text into an array
            LinedText = totalRawText.Split(pattern);
            //iterating through each index of the total text lines
          /*  foreach (string t in LinedText)
            {
                Console.WriteLine(t);
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine();
            }*/
        }
        /// <summary>
        /// writes the contents of string array to a new file in each line
        /// </summary>
        /// <param name="text"></param>
        /// <param name="outputLocation"></param>
        public void Writer(string[] text,string outputLocation)
        {
            foreach(string t in text)
            {
                System.IO.File.AppendAllText(outputLocation, t);
            }
        }
    }
}

