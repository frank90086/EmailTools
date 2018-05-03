using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Curves.SendEmail.Tools.Models
{
    public class ExcelModel
    {
        public string Email { get; set; }
        public string Name { get; set; }
        public string C { get; set; }
        public int M { get; set; }
        public int Box { get; set; }
        public int B { get; set; }
        public int In { get; set; }
        public int N { get; set; }
        public int All { get; set; }
        public string IsSend { get; set; }
    }
}