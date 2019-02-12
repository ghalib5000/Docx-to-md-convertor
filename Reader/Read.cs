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

        private static string FileLocation2 = @"D:\convertor\test - Copy - Copy\word\document.xml";
        private static string FileLocation = @"D:\convertor\test.docx";
        private static string write_loc = @"D:\";
        private static string[] StandbyText = new string[3];
        private static string[] filedata = new string[15];
        private static string[] finalText = new string[15];
        private static string bold ="b" ;
        private static string italic = "i";
        int boldstart = 0, italicstart = 0, boldlines = 0, italiclines = 0, typestart = 0, i = 0, errcnt = 0;
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
                    mover(FileLocation);
                   // createWriteLocation();
                    //function();
                    //ender(i);
                    //dispOnConsole(finalText);
                   // Writer(finalText, write_loc);
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
        public void mover(string loc)
        {
            System.IO.File.Move(loc, @"C:\Windows\Temp\temp.zip");
        }
        public void extractor(string loc)
        {

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
                                if (errcnt == 0)
                                {
                                    //the type of style
                                    styleidentifier = style.LocalName;
                                }
                                //the text of the above style to copy
                                copytext = text.InnerText;
                                //if they are not equal then there are 2 style that are being used
                                if (locNode3.FirstChild==locNode3.LastChild)
                                {
                                    XmlNodeType type = style.NodeType;
                                    if (type == XmlNodeType.Element)
                                    {
                                        textchecker = text.LocalName;
                                        //the final checker
                                        if (textchecker == "r")
                                        {
                                            typestart = 1;
                                            checker(styleidentifier, i);
                                        }
                                    }
                                    else if (type == XmlNodeType.Text && typestart == 0&&(styleidentifier == "bi"|| styleidentifier == "ib"))
                                    {
                                        //styleidentifier = "";
                                        checker(styleidentifier, i);
                                        //copytext = text.InnerText;
                                        //finalText[i] += copytext;
                                        typestart = 0;
                                    }
                                    else if (type == XmlNodeType.Text && typestart == 0)
                                    {
                                        //styleidentifier = "";
                                        checker(styleidentifier, i);
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
                                        // styleidentifier = "";
                                        checker(styleidentifier, i);
                                        finalText[i] += " ";
                                        typestart = 0;
                                    }
                                    errcnt = 0;
                                }
                                else
                                {
                                    errcnt = 1;
                                    if(styleidentifier=="b")
                                    {
                                        styleidentifier += "i";
                                    }
                                    else if (styleidentifier == "i")
                                    {
                                        styleidentifier += "b";
                                    }
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
        public void checker(string style, int i)
        {
            //for bold italic or italic bold
            {
                //if none of them are started
                if (style == "bi" && boldstart == 0 && italicstart == 0)
                {
                    finalText[i] += "**_" + copytext;
                    boldstart = 1;
                    italicstart = 1;
                }
                //if none of them are started
                else if (style == "ib" && boldstart == 0 && italicstart == 0)
                {
                    finalText[i] += "_**" + copytext;
                    boldstart = 1;
                    italicstart = 1;
                }
                //if none of them are started
                else if (style == "bi" && boldstart == 1 && italicstart == 0)
                {
                    finalText[i] += "_" + copytext;
                    boldlines = 1;
                    italicstart = 1;
                }
                //if none of them are started
                else if (style == "ib" && boldstart == 0 && italicstart == 1)
                {
                    finalText[i] += "**" + copytext;
                    boldstart = 1;
                    italiclines = 1;
                }
                //if none of them are started
                else if (style == "bi" && boldstart == 1 && italicstart == 1)
                {
                    finalText[i] += copytext;
                    boldlines = 1;
                    italiclines = 1;
                }
                //if none of them are started
                else if (style == "ib" && boldstart == 1 && italicstart == 1)
                {
                    finalText[i] += "**_" + copytext;
                    boldlines = 1;
                    italiclines = 1;
                }
                //for bold data
                //for bold starting
                else if (style == bold && boldstart == 0)
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
                    finalText[i] += "**";
                    boldstart = 0;
                }
                //for italic data
                //for italic starting
                else if (style == italic && italicstart == 0)
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
                    finalText[i] += "_";
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

