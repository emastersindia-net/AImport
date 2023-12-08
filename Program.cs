using AImport.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AImport
{
    internal class Program
    {
        static public void Main()
        {
            FileImport obj = new FileImport();
            //obj.GetMISDataSearch(null);

            obj.ProcessFile();
            //Console.ReadKey();
        }
    }
}
