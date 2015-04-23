using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GemBox.Spreadsheet;

namespace ExcelAddIn3
{
    class ExcelInformation
    {
    }

    // Good source if more are required: https://docs.google.com/spreadsheet/ccc?key=0AsHau4_IeCfwdG45c0VhcjBBUGNreVo2ZzNNRU1BT0E#gid=0
    // And http://superuser.com/questions/440330/list-of-all-excel-functions-by-version
    public class ExcelFunction
    {
        public string Name { get; private set; }

        /*
        public List<object> DefaultArgs { get; private set; }
        public int NumArgs { get { return DefaultArgs.Count; } }
        public int NumRequiredArgs {  get { return DefaultArgs.Count(v => v == null); } }
        public int NumDefaultArgs { get { return NumArgs - NumRequiredArgs; } }
        public bool HasVarArgs { get; private set; }
        */
        /// <summary>True if parameters are array as argument, e.g. SUM.</summary>
        public bool ParamsAsArray { get; private set; }
        

        /// <summary>Version this function first appears in. Might be innacurate for functions introduced earlier than 11.</summary>
        /// <example>IFERROR was introduced in 2007, which is v12.0, so this should be 12 for that function.</example>
        public int Since { get; private set; }

        private static readonly Dictionary<String, ExcelFunction> known = new Dictionary<string, ExcelFunction>(); 
        public static readonly IReadOnlyDictionary<String, ExcelFunction> Known = new ReadOnlyDictionary<string, ExcelFunction>(known);
        public static readonly IEnumerable<ExcelFunction> Known2 = known.Values; 

        private ExcelFunction(string name)
        {
            Name = name;
            known.Add(name, this);

            // Defaults:
            //DefaultArgs = new List<object>();
            //HasVarArgs = false;
            ParamsAsArray = false;
            Since = 11; // Functions only explicity need to add this if it's later than 2003
        }

        public static readonly ExcelFunction SUM = new ExcelFunction("SUM")
        {
            //DefaultArgs = {null},
            //HasVarArgs = true,
            ParamsAsArray = true
        };
    }
}
