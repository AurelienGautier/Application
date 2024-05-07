using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Parser
{
    internal abstract class Parser
    {
        protected List<Data.Piece>? dataParsed;

        /*-------------------------------------------------------------------------*/

        public abstract List<Data.Piece> ParseFile(String fileToParse);

        /*-------------------------------------------------------------------------*/
    
        public abstract String GetFileExtension();

        /*-------------------------------------------------------------------------*/
    }
}
