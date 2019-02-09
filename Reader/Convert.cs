using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.IO.Packaging;
using System.Linq;
using System.Threading.Tasks;
using System.Data;
using System.Text.RegularExpressions;

namespace Reader
{
    class Convertor:Read
    {

        private static string FileLocation = @"D:\convertor\test - Copy\word\document.xml";
        private static string FileLocation2 = @"D:\convertor\test.docx";
        private static string write_loc = @"D:\";
        private const string boldfinder = "<w:b />";
        private const string italicfinder = "<w:i />";
        private const string spaces = "xml:space= \"preserve\"";
        /*static XmlNode node = null;
        static XmlDocument xDoc = null;
        static XmlNode locNode = null;*/
        private static int boldstart = 0, italicstart = 0, boldlines = 0, italiclines = 0, linecount = 0, k = 0;
        private static string[] filedata = new string[linecount];
        private static string[] totdata = new string[linecount];
        private string standby = "";
        public void read()
        {
           
            try
            {
                //  Console.Write("enter the location of the file: ");
                //  FileLocation = Console.ReadLine();
                string t = Path.GetExtension(FileLocation);
                if (t == ".docx" || t == ".xml")
                {
                    XmlTextReader textReader = new XmlTextReader(FileLocation);
                    textReader.Read();
                    XmlDocument xdoc = new XmlDocument();
                    xdoc.Load(textReader);
                    /*foreach (XmlNode node in xdoc.DocumentElement.ChildNodes)
                    {
                        string t2 = node.OuterXml;  //the formatting style of text
                        string text = node.InnerText;  //main text
                    }*/
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
        }

        public void writer()
        {
            int i = 0;
            string dirLoc = FileLocation;
            try
            {
                // new xdoc instance 
                XmlDocument xDoc = new XmlDocument();
                //load up the xml from the location 
                xDoc.Load(dirLoc);
                Convertor text = new Convertor();
                // Convertor italic= new Convertor();

               // foreach (XmlNode node in xDoc.DocumentElement.ChildNodes)
                {
                   // foreach (XmlNode locNode in node)
                    {
                        linecount++;
                    }
                    filedata = new string[15];
                    totdata = new string[15];
                   // foreach (XmlNode locNode in node)
                   foreach(string t in fintext)
                    {
                        // get the content of the loc node 
                        string maintext = t; //locNode.InnerXml;
                        //  Match mc = Regex.Match(loc, @"(?</w:rPr><w:t>).+(?=</w:t>)");

                        //  Console.WriteLine(mc);
                        string[] seperator = new string[] { "<w:t>", "</w:t>", "<w:t xml:space=\"preserve\">", "<w:t xml:space=\"preserve\"> " };
                        string[] fintext = maintext.Split(seperator, StringSplitOptions.None);
                        for (int j = 1; j < fintext.Length - 1; j += 2, k++)
                        {
                            //  filedata[k] = fintext[j];
                            //    Console.WriteLine(fintext[j]);
                            text.boldchecker(fintext[j - 1], fintext[j], k);
                            if (standby == "")
                            {
                                standby = filedata[k];
                                totdata[i] += standby;
                            }
                            else if (standby == filedata[k - 1])
                            {
                                standby = filedata[k];
                                totdata[i] += standby;
                            }
                            else if (standby != filedata[k - 1])
                            {
                                totdata[i - 1] = totdata[i - 1].Replace(standby, filedata[k - 1]);
                                standby = filedata[k];
                                totdata[i] += standby;
                                //totdata[i-1]  += filedata[k-1];
                            }
                            else if (standby != filedata[k] && standby != null)
                            {
                                standby = filedata[k];
                                totdata[i] += standby;
                            }
                            else
                            {
                                totdata[i] += standby;
                            }
                            // filedata[k];
                        }
                        //an array of string that will store the final output of line
                        //foreach(string t in filedata)
                        {
                            //   totdata[i] += t;
                        }
                        i++;
                    }
                    if (boldstart == 1)
                    {
                        totdata[i-2] += "**";
                    }
                    if (italicstart == 1)
                    {
                        totdata[i-2] += "_";
                    }
                }

                //writelocationchecker();

                foreach (string t in totdata)
                {
                    System.IO.File.AppendAllText(write_loc + "output.md", t + Environment.NewLine);
                }

                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine();
                Dispstring(filedata);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw new Exception();
            }
            Console.WriteLine("All Done :-)");
        }


        public static void writelocationchecker()
        {
            string response = "";
            //ask to convert the file into md
            Console.Write("Convert to .md file?: ");
            response = Console.ReadLine();
            if (response == "yes" || response == "y" || response == "Y")
            {
                //asking for location to save the file
                Console.Write("enter the location for saving the file:");
                write_loc = Console.ReadLine();
                // write_loc += ".md";
                if (!Directory.Exists(write_loc))
                {
                    //asking for directory to be created
                    Console.Write("Directory does not exsits....create one?: ");
                    response = Console.ReadLine();
                    if (response == "yes" || response == "y" || response == "Y")
                    {
                        Directory.CreateDirectory(write_loc);
                    }
                    else
                    {
                        throw new Exception("Exiting...");
                    }
                }
            }
        }
        public void boldchecker(string OuterText, string InnerText, int i)
        {
            try
            {
                //bold data  
                {
                    filedata[i] = InnerText;
                    //for bold starting
                    if ((Regex.IsMatch(OuterText, boldfinder)) && boldstart == 0)
                    {
                        boldstart = 1;
                        filedata[i] = "**" + InnerText;
                        Console.WriteLine(filedata[i]);
                    }
                    /*
                    //if italic is already started but bold is not
                    else if ((Regex.IsMatch(OuterText, italicfinder)) && italicstart == 1 && boldstart == 0)
                    {
                        boldstart = 1;
                        filedata[i] += " **" + InnerText;
                        //    Console.WriteLine(filedata[i]);
                    }
                    */
                    //when bold is initialized
                    else if ((Regex.IsMatch(OuterText, boldfinder)) && boldstart == 1)
                    {
                        boldlines++;
                        filedata[i] = InnerText;
                        Console.WriteLine(filedata[i]);
                    }
                    //when bold ends
                    else if (!(Regex.IsMatch(OuterText, boldfinder)) && boldstart == 1)
                    {
                        // boldlines++;
                        filedata[i - 1] += "**";
                        Console.WriteLine(filedata[i]);
                        boldstart = 0;
                    }
                    //for bold italic
                    if ((Regex.IsMatch(OuterText, boldfinder)) && (Regex.IsMatch(OuterText, italicfinder)))
                    {
                        //  Console.WriteLine(filedata[i]);
                    }
                }
                //italic data
                {
                    //for italic starting
                    if ((Regex.IsMatch(OuterText, italicfinder)) && italicstart == 0)
                    {
                        italicstart = 1;
                        filedata[i] = "_" + InnerText;
                        Console.WriteLine(filedata[i]);
                    }/*
                    //if bold is already started but italic is not
                    else if((Regex.IsMatch(OuterText, italicfinder)) && italicstart == 0 && boldstart == 1)
                        {
                            italicstart = 1;
                            filedata[i] += " _" + InnerText;
                            //    Console.WriteLine(filedata[i]);
                        }
                        */
                     //when italic is initialized
                    else if ((Regex.IsMatch(OuterText, italicfinder)) && italicstart == 1)
                    {
                        italiclines++;
                        filedata[i] = InnerText;
                        Console.WriteLine(filedata[i]);
                    }
                    //when italic ends
                    else if (!(Regex.IsMatch(OuterText, italicfinder)) && italicstart == 1)
                    {
                        // italiclines++;
                        filedata[i - 1] += "_";
                        Console.WriteLine(filedata[i]);
                        italicstart = 0;
                    }
                    //for bold italic
                    if ((Regex.IsMatch(OuterText, italicfinder)) && (Regex.IsMatch(OuterText, boldfinder)))
                    {
                        //   Console.WriteLine(filedata[i]);
                    }
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            Console.WriteLine("Done :-)");
        }

        public static void Dispstring(string[] data)
        {
            foreach (string t in data)
            {
                Console.WriteLine(t);
            }
        }























    }
}

/*public void read2()
{
    int ws = 0;
    int pi = 0;
    int dc = 0;
    int cc = 0;
    int ac = 0;
    int et = 0;
    int el = 0;
    int xd = 0;
    // Read a document  
    XmlTextReader textReader = new XmlTextReader(FileLocation);
    // Read until end of file  
    while (textReader.Read())
    {
        XmlNodeType nType = textReader.NodeType;
        // If node type us a declaration  
        if (nType == XmlNodeType.XmlDeclaration)
        {
            Console.WriteLine("Declaration:" + textReader.Name.ToString());
            xd = xd + 1;
        }
        // if node type is a comment  
        if (nType == XmlNodeType.Comment)
        {
            Console.WriteLine("Comment:" + textReader.Name.ToString());
            cc = cc + 1;
        }
        // if node type us an attribute  
        if (nType == XmlNodeType.Attribute)
        {
            Console.WriteLine("Attribute:" + textReader.Name.ToString());
            ac = ac + 1;
        }
        // if node type is an element  
        if (nType == XmlNodeType.Element)
        {
            Console.WriteLine("Element:" + textReader.Name.ToString());
            el = el + 1;
        }
        // if node type is an entity\  
        if (nType == XmlNodeType.Entity)
        {
            Console.WriteLine("Entity:" + textReader.Name.ToString());
            et = et + 1;
        }
        // if node type is a Process Instruction  
        if (nType == XmlNodeType.Entity)
        {
            Console.WriteLine("Entity:" + textReader.Name.ToString());
            pi = pi + 1;
        }
        // if node type a document  
        if (nType == XmlNodeType.DocumentType)
        {
            Console.WriteLine("Document:" + textReader.Name.ToString());
            dc = dc + 1;
        }
        // if node type is white space  
        if (nType == XmlNodeType.Whitespace)
        {
            Console.WriteLine("WhiteSpace:" + textReader.Name.ToString());
            ws = ws + 1;
        }
    }
    // Write the summary  
    Console.WriteLine("Total Comments:" + cc.ToString());
    Console.WriteLine("Total Attributes:" + ac.ToString());
    Console.WriteLine("Total Elements:" + el.ToString());
    Console.WriteLine("Total Entity:" + et.ToString());
    Console.WriteLine("Total Process Instructions:" + pi.ToString());
    Console.WriteLine("Total Declaration:" + xd.ToString());
    Console.WriteLine("Total DocumentType:" + dc.ToString());
    Console.WriteLine("Total WhiteSpaces:" + ws.ToString());
}*/
/*    public void write()
    {
        // Create a new file in D:\\ dir  
        XmlTextWriter textWriter = new XmlTextWriter("D:\\myXmFile.xml", null);
        // Opens the document  
        textWriter.WriteStartDocument();
        // Write comments  
        textWriter.WriteComment("First Comment XmlTextWriter Sample Example");
        textWriter.WriteComment("myXmlFile.xml in root dir");
        // Write first element  
        textWriter.WriteStartElement("Student");
        textWriter.WriteStartElement("r", "RECORD", "urn:record");
        // Write next element  
        textWriter.WriteStartElement("Name", "");
        textWriter.WriteString("Student");
        textWriter.WriteEndElement();
        // Write one more element  
        textWriter.WriteStartElement("Address", "");
        textWriter.WriteString("Colony");
        textWriter.WriteEndElement();
        // WriteChars  
        char[] ch = new char[3];
        ch[0] = 'a';
        ch[1] = 'r';
        ch[2] = 'c';
        textWriter.WriteStartElement("Char");
        textWriter.WriteChars(ch, 0, ch.Length);
        textWriter.WriteEndElement();
        // Ends the document.  
        textWriter.WriteEndDocument();
        // close writer  
        textWriter.Close();
    }*/





