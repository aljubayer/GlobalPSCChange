using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GlobalPSC
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputFile = Directory.GetCurrentDirectory() + @"\HuaweiPSCChangeWOTemplate.xlsx";
                string templateFile = Directory.GetCurrentDirectory() + @"\3G Radio Network Planning Data Template.xlsm";
                HuaweiPSCByCME aCme = new HuaweiPSCByCME();
                string message = aCme.ProcessTemplate(inputFile, templateFile);
                Console.WriteLine(message);
                Console.ReadKey();
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception: " + exception.Message);
                Console.ReadKey();
            }

        }
    }
}
