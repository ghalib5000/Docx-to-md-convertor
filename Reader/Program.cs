using System;
using System.IO;
using System.Xml;


namespace Reader 
{
    class Program: Logger
    {
        static void Main(string[] args)
        {
            
            IRead document = new Read();
              document.File();
        }
    }
}
