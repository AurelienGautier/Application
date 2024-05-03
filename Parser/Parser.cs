using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Parser
{
    internal abstract class Parser
    {
        protected Dictionary<string, string>? header;
        protected List<Data.Piece>? dataParsed;

        /*-------------------------------------------------------------------------*/

        public abstract List<Data.Piece> ParseFile(String fileToParse);

        /*-------------------------------------------------------------------------*/

        public Dictionary<string, string> GetHeader()
        {
            if(this.header == null)
            {
                this.header = new Dictionary<string, string>();
            }
            
            return this.header;
        }

        /*-------------------------------------------------------------------------*/
    
        public abstract String GetFileExtension();

        /*-------------------------------------------------------------------------*/
    }
}
