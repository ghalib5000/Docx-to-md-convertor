using System;
using System.IO;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;


namespace Reader
{
    class Program
    {
        static void Main(string[] args)
        {
            IRead document = new Read();
             document.File();
           }
    }
}
