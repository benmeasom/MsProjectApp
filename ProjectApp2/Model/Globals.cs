using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectApp2.Model
{
    static class Globals
    {
        private static ExcelMapFilePaths _excelMapFilePaths = null;

        public static ExcelMapFilePaths ExcelMapFilePaths
        {
            get { return _excelMapFilePaths; }
            set { _excelMapFilePaths = value; }
        }
    }
}
