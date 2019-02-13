using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;

namespace Reader
{
    interface IRead
    {
        void File();
        void mover(string FileLocation2, string copyPath);
        void extractor(string copyPath, string extract_path);
        void createWriteLocation();
        void function();
        void ender(int i);
        void dispOnConsole(string[] finalText);
        void Writer(string[]finalText, string write_loc);
    }
}
