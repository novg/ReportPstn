using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportPstn
{
    public class ReportEventArgs : EventArgs
    {
        public string DefaultFilePath { get; internal set; }
    }
}
