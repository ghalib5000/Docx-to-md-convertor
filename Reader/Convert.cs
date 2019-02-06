using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.IO.Packaging;
using System.Linq;
using System.Threading.Tasks;
using System.Data;

namespace Reader
{
    class Convertor
    {

        private static string FileLocation = "D:\\convertor\\test - Copy\\word\\document.xml";
        private static string FileLocation2 = "D:\\convertor\\test.docx";
        private static string write_loc = "D:\\out.md";
        private static string[] filedata = new string[30];
        static XmlNode node = null;
        static XmlDocument xDoc = null;
        static XmlNode locNode = null;
         int linecount = 0;
         int textstyle = 0;
         int totalLines = 0;
        public void read()
        {
            //try
            {
                //  Console.Write("enter the location of the file: ");
                //  FileLocation = Console.ReadLine();
                XmlTextReader textReader = new XmlTextReader(FileLocation);
                textReader.Read();
                XmlDocument xdoc = new XmlDocument();
                xdoc.Load(textReader);
                foreach (XmlNode node in xdoc.DocumentElement.ChildNodes)
                {
                    string t2 = node.OuterXml;
                    string text = node.InnerText;
                    // File.Create(write_loc);
                    File.AppendAllText(write_loc, text);
                    File.AppendAllText(write_loc, t2);
                    Console.WriteLine(text);
                    Console.WriteLine(t2);
                    //or loop through its children as well
                }
            }
        }
        public void disp()
        {
            XmlTextReader textReader = new XmlTextReader(FileLocation);
            textReader.Read();
            // If the node has value  
            while (textReader.Read())
            {

                // Move to fist element  
                textReader.MoveToElement();
                Console.WriteLine("XmlTextReader Properties Test");
                Console.WriteLine("===================");
                // Read this element's properties and display them on console  
                Console.WriteLine("Name:" + textReader.Name);
                Console.WriteLine("Base URI:" + textReader.BaseURI);
                Console.WriteLine("Local Name:" + textReader.LocalName);
                Console.WriteLine("Attribute Count:" + textReader.AttributeCount.ToString());
                Console.WriteLine("Depth:" + textReader.Depth.ToString());
                Console.WriteLine("Line Number:" + textReader.LineNumber.ToString());
                Console.WriteLine("Node Type:" + textReader.NodeType.ToString());
                Console.WriteLine("Attribute Count:" + textReader.Value.ToString());
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
                Convertor bold = new Convertor();
                Convertor italic= new Convertor();

                foreach (XmlNode node in xDoc.DocumentElement.ChildNodes)
                {
                    foreach (XmlNode locNode in node)
                    {
                        // get the content of the loc node 
                        string loc = locNode.InnerText;
                       bold.boldchecker('b', "**", node, locNode, i);
                        italic.boldchecker('i', "_", node, locNode, i);
                        i++;
                    }
                }
                Dispstring(filedata);
            }
            catch { }
            Console.WriteLine("All Done :-)");
        }


        public  void boldchecker(char type, string style, XmlNode node, XmlNode locNode, int i)
        {
           

            string dirLoc = FileLocation;
            try
            {

                // get the content of the loc node 
                string loc = locNode.InnerText;
                if (locNode.OuterXml.Contains("<w:pPr><w:rPr><w:" + type + " /></w:rPr></w:pPr><w:r><w:rPr><w:" + type + " /></w:rPr>") && textstyle == 0 && linecount == 0)
                {
                    //bold starts here
                    filedata[i] = style + locNode.InnerText;
                    Console.WriteLine("FOUUUUUND IT!!!!!!!" + style);
                    //   Console.WriteLine(locNode.OuterXml);
                    textstyle = 1;
                    totalLines++;
                    //linecount++;
                }
                else if (locNode.OuterXml.Contains("<w:pPr><w:rPr><w:" + type + " /></w:rPr></w:pPr><w:r><w:rPr><w:" + type + " /></w:rPr>") && textstyle == 1)
                {
                    filedata[i] = locNode.InnerText;
                    totalLines++;
                    linecount++;
                }
                else if (!locNode.OuterXml.Contains("<w:pPr><w:rPr><w:" + type + " /></w:rPr></w:pPr><w:r><w:rPr><w:" + type + " /></w:rPr>") && textstyle == 1 && linecount >= 1)
                {
                    Console.WriteLine("STYLE ENDED!");
                    filedata[i-1] += style;
                    filedata[i] = locNode.InnerText;
                    textstyle = 0;
                    linecount = 0;
                }
                //bold style terminator
                else if (linecount == 0 && textstyle == 1)
                {
                    Console.WriteLine("STYLE ENDED!");
                    filedata[i - 1] += style;
                    filedata[i] = locNode.InnerText;
                    textstyle = 0;
                    linecount = 0;

                }
                else if (locNode.OuterXml.Contains("<w:r><w:t>") && linecount == 1)
                {
                    filedata[i] = locNode.InnerText;
                    linecount = 0;
                }
                
                //foreach (string t in filedata)
                {
                //    Console.WriteLine(t);
                }
            }
            catch { }
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





 