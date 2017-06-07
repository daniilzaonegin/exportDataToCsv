using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace exportDataToXlsx
{
    class ConsoleWriter
    {
        public static void Write(string message)
        {
            Console.WriteLine("["+DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:fff")+ "]"+ message);
        }
    }
}
