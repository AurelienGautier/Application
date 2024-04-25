using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Exceptions
{
    public class MeasureTypeNotFoundException : Exception
    {
        public MeasureTypeNotFoundException(string measureType, string file, int line) :
            base("Type de mesure \"" + measureType + "\" non identifié dans le fichier \"" + file + "\" à la ligne " + line)
        {
        }

        public MeasureTypeNotFoundException(string measureType, string file, string cell) : 
            base("Type de mesure \"" + measureType + "\" non identifié dans le fichier \"" + file + "\" dans la cellule \""+ cell + "\"")
        {
        }
    }
}
