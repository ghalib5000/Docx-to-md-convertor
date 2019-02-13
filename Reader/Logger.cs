using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
namespace Reader
{
    class Logger
    {
        static string date = DateTime.Now.ToString("yyyy_MM_dd");
        private readonly string logloc = @"C:\Windows\Temp\Log_"+date+".log";
        public Logger()
        {
            string t = "Logging started at " + DateTime.Now+Environment.NewLine;
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(t);
            File.AppendAllText(logloc, t);
        }
        public void Information(string info)
        {
            string data = "IMFORMATION: Work " + info + " Done at " + DateTime.Now + Environment.NewLine;
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(data);
            File.AppendAllText(logloc, data);
        }
        public void Error(Exception error)
        {
            string data = "ERROR: Error "+ error.ToString()+ " found  at " + DateTime.Now + Environment.NewLine;
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine(data);
            File.AppendAllText(logloc, data);
            throw (error);
        }
    }
}
