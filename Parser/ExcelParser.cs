using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Parser
{
    internal class ExcelParser : Parser
    {
        public ExcelParser()
        {
        }

        public List<Data.Piece> ParseFile(string fileToParse)
        {
            return new List<Data.Piece>();
        }

        public Dictionary<string, string> GetHeader()
        {
            return new Dictionary<string, string>();
        }
    }
}
