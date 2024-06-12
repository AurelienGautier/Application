using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Exceptions
{
    public class IncorrectFormatException : Exception
    {
        public IncorrectFormatException() : base("Le format du fichier est incorrect.") { }
        public IncorrectFormatException(string message) : base(message) { }
    }
}
