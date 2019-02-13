using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Threading.Tasks;
using System.Data;

namespace Reader
{
    class Read : Logger, IRead
    {

        private static string FileLocation = @"C:\Windows\Temp\temp\word\document.xml";
        private static string FileLocation2 = @"D:\convertor\test.docx";
        private static string extract_path = @"C:\Windows\Temp\temp\";
        private static string copyPath = @"C:\Windows\Temp\temp.zip";
        private static string write_loc = @"D:\";
        private static string[] StandbyText = new string[3];
       // private static string[] filedata = new string[10];
        private static string[] finalText = new string[1000];
        private static string bold = "b";
        private static string italic = "i";
        private int boldstart = 0, italicstart = 0, boldlines = 0, italiclines = 0, typestart = 0, i = 0, errcnt = 0;
        private string styleidentifier = "", copytext = "", textchecker = "";
        private string response = "";
        Logger loging = new Logger();
        /// <summary>
        /// reads the contents of the file and does stuff to it :3
        /// </summary>
        public void File()
        {

            try
            {
                   Console.Write("Enter the location of the file: ");
                   FileLocation2 = Console.ReadLine();
                string t = Path.GetExtension(FileLocation2);
                if (t == ".docx" || t == ".xml")
                {
                    mover(FileLocation2, copyPath);
                    extractor(copyPath, extract_path);
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
                loging.Error(ex);
                // Console.WriteLine(ex);
            }
        }
        /// <summary>
        /// moves the loc file to the path destination
        /// </summary>
        /// <param name="loc"></param>
        /// <param name="path"></param>
        public void mover(string loc, string path)
        {
            loging.Information("Started mover method");
            System.IO.File.Copy(loc, path, true);
        }
        /// <summary>
        /// extracts the loc file to dest destination
        /// </summary>
        /// <param name="loc"></param>
        /// <param name="dest"></param>
        public void extractor(string loc, string dest)
        {
            loging.Information("Started extracting file");
            System.IO.Compression.ZipFile.ExtractToDirectory(loc, dest, true);
        }
        public void ender(int i)
        {
            loging.Information("started ender method");
            if (boldstart == 1)
            {
                finalText[i - 2] += "**";
            }
            if (italicstart == 1)
            {
                finalText[i - 2] += "_";
            }
        }
        /// <summary>
        /// a function which does all the important things
        /// </summary>
        public void function()
        {
            loging.Information("Started main function");
            // new xdoc instance 
            XmlDocument xDoc = new XmlDocument();
            //load up the xml from the location 
            xDoc.Load(FileLocation);
            i = 0;
            //important logic that helps solve the problem
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
                                if (locNode3.FirstChild == locNode3.LastChild)
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
                                    else if (type == XmlNodeType.Text && typestart == 0 && (styleidentifier == "bi" || styleidentifier == "ib"))
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
                                    if (styleidentifier == "b")
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
        /// <summary>
        /// displays the text on the console
        /// </summary>
        /// <param name="text"></param>
        public void dispOnConsole(string[] text)
        {
            loging.Information("Started dispOnConsole method");
            foreach (string t in text)
            {
                if(t!=null)
                Console.WriteLine(t);
            }
        }
        /// <summary>
        /// checks for styles in text
        /// </summary>
        /// <param name="style"></param>
        /// <param name="i"></param>
        public void checker(string style, int i)
        {
            try
            {
                loging.Information("Started Checker Method");
                //for bold italic or italic bold

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
                    finalText[i] += "_";
                    italicstart = 0;
                }
            }
            catch (Exception ex)
            {
                loging.Error(ex);
            }
        }
        /// <summary>
        /// Creates a new location to write the contents of file to
        /// </summary>
        public void createWriteLocation()
        {
            try
            {
                loging.Information("Started createWriteLocation Method");
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
                loging.Error(dirntfnd);
                //Console.WriteLine(dirntfnd);
            }
            catch (Exception ex)
            {
                loging.Error(ex);
                //Console.WriteLine(ex);
            }
        }
        /// <summary>
        /// writes the contents of string array to a new file in each line
        /// </summary>
        /// <param name="text"></param>
        /// <param name="outputLocation"></param>
        public void Writer(string[] text, string outputLocation)
        {
            loging.Information("Started Writer Method");
            foreach (string t in text)
            {
                if (t != null)
                {
                    loging.Information("Copied " + t + " to file at location " + write_loc);
                    System.IO.File.AppendAllText(outputLocation, t + Environment.NewLine);
                }
            }
        }
    }
}

