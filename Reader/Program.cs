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
           // Convertor c = new Convertor();
         //   c.read();
          //   c.disp();
          //   c.writer();
             string docin,docout;
             Read document = new Read();
             docin =document.File();
             docout=document.Write();
             document.Convert(docin, docout);
        }
    }
}
