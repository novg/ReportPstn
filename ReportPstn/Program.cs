using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportPstn
{
    sealed class Program
    {
        public static void Main(string[] argv)
        {
            string fileName = argv[0];
            Report report = new Report(fileName);
            report.Create();
        }
    }
}
