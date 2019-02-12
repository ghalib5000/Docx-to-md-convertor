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
        private static string[] StandbyText = new string[3];
        private static string[] filedata = new string[15];
        private static string[] finalText = new string[15];
        private static string bold ="b" ;
        private static string italic = "i";
        int boldstart = 0, italicstart = 0, boldlines = 0, italiclines = 0, typestart = 0, i = 0;
        string styleidentifier = "",copytext = "",textchecker="";
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
                    createWriteLocation();
                    function();
                    ender(i);
                    dispOnConsole(finalText);
                    Writer(finalText, write_loc);
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
        public void ender(int i)
        {
            if (boldstart == 1)
            {
                finalText[i - 2] += "**";
            }
            if (italicstart == 1)
            {
                finalText[i - 2] += "_";
            }
        }
        public void function()
        {   // new xdoc instance 
            XmlDocument xDoc = new XmlDocument();
            //load up the xml from the location 
            xDoc.Load(FileLocation);
            i = 0;
            //important logic that can help solve the problem
            foreach (XmlNode node in xDoc.DocumentElement.ChildNodes)
            {
                foreach (XmlNode locNode in node)
                {
                    foreach (XmlNode text in locNode)
                    {
                         foreach (XmlNode locNode3 in text)
                        {
                            foreach (XmlNode style in locNode3)
                            {
                                XmlNodeType type = style.NodeType;
                                if (type == XmlNodeType.Element)
                                {
                                    textchecker = text.LocalName;
                                    //the type of style
                                    styleidentifier = style.LocalName;
                                    //the text of the above style to copy
                                    copytext = text.InnerText;
                                    //the final checker
                                    if (textchecker == "r")
                                    {
                                    typestart = 1;
                                        checker(styleidentifier, bold, i);
                                    }
                                }
                                else if (type == XmlNodeType.Text&&typestart==0)
                                {
                                    styleidentifier = "";
                                    checker(styleidentifier, bold, i);
                                    copytext = text.InnerText;
                                    finalText[i] += copytext;
                                    typestart = 0;
                                }
                                else
                                {
                                    typestart = 0;
                                }
                                if (type == XmlNodeType.SignificantWhitespace)
                                {
                                    styleidentifier = "";
                                    checker(styleidentifier, bold, i);
                                    finalText[i] += " ";
                                    typestart = 0;
                                }
                            }
                        }
                    }
                   // Console.WriteLine(finalText[i]);
                    i++;
                }
            }
        }
        public void dispOnConsole(string[] text)
        {
           foreach(string t in text)
            {
                Console.WriteLine(t);
            }
        }
        public void checker(string style,string bold,int i)
        {
            //for bold data
            {
                //for bold starting
                if (style == bold && boldstart == 0)
                {
                    finalText[i] += "**" + copytext;
                    boldstart = 1;
                }
                //when bold is initialized
                else if (style == bold && boldstart == 1)
                {
                    boldlines++;
                    finalText[i] += copytext;
                }
                //when bold ends
                else if (!(style == bold) && boldstart == 1)
                {
                    // boldlines++;
                    finalText[i - (boldlines)] += "**";
                    boldstart = 0;
                }
            }
            //for italic data
            {
                //for italic starting
                if (style == italic && italicstart == 0)
                {
                    finalText[i] += "_" + copytext;
                    italicstart = 1;
                }
                //when italic is initialized
                else if (style == italic && italicstart == 1)
                {
                    italiclines++;
                    finalText[i] += copytext;
                }
                //when italic ends
                else if (!(style == italic) && italicstart == 1)
                {
                    // boldlines++;
                    finalText[i - italiclines+1] += "_";
                    italicstart = 0;
                }
                /*
                else if ((style != italic) && (style != bold) && (italicstart == 0) && (boldstart == 0))
                {
                    finalText[i] += copytext;
                }
                */
            }

        }
        /// <summary>
        /// Creates a new location to write the contents of file to
        /// </summary>
        public string createWriteLocation()
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
                        if (response[0] == 'y')
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
        /// writes the contents of string array to a new file in each line
        /// </summary>
        /// <param name="text"></param>
        /// <param name="outputLocation"></param>
        public void Writer(string[] text, string outputLocation)
        {
            foreach (string t in text)
            {
                System.IO.File.AppendAllText(outputLocation, t);
            }
        }
    }
}

