using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Parser
{
    internal interface Parser
    {
        List<Data.Piece> ParseFile(String fileToParse);
        Dictionary<string, string> GetHeader();
    }
}
