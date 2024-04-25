using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Exceptions
{
    public class ExcelFileAlreadyInUseException : Exception
    {
        public ExcelFileAlreadyInUseException(String fileName) : 
            base("Le fichier excel \"" + fileName + "\" est déjà en cours d'utilisation. Veuillez le fermer puis réessayer.") { }
    }
}
