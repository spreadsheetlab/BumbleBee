using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BumbleBee.Refactorings.Util
{
    public class ParserSheetReference
    {
        public string Worksheet { get; }
        public string FileName { get; }
        public string WorksheetClean => Worksheet.Remove(Worksheet.Length - 1);

        public ParserSheetReference(string filename, string worksheet)
        {
            if (!worksheet.EndsWith("!"))
            {
                worksheet += "!";
            }
            Worksheet = worksheet;
            FileName = filename;
        }

        public override bool Equals(object o)
        {
            var other = o as ParserSheetReference;
            if (ReferenceEquals(other, this)) return true;
            if (ReferenceEquals(other, null)) return false;

            return Worksheet == other.Worksheet && FileName == other.FileName;
        }

        public override int GetHashCode()
        {
            int hash = 17;
            unchecked
            {
                hash = (hash * 3) + Worksheet.GetHashCode();
                hash = (hash * 3) + FileName.GetHashCode();
            }
            return hash;
        }

        public string toString()
        {
            return FileName + Worksheet;
        }
    }
}
