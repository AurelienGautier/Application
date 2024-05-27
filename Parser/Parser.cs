namespace Application.Parser
{
    public abstract class Parser
    {
        protected List<Data.Piece>? dataParsed;

        /*-------------------------------------------------------------------------*/

        public abstract List<Data.Piece> ParseFile(String fileToParse);

        /*-------------------------------------------------------------------------*/
    
        public abstract String GetFileExtension();

        /*-------------------------------------------------------------------------*/
    }
}
