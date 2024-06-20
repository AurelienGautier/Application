using Newtonsoft.Json;
using System.IO;

namespace Application.Facade
{
    public class JsonLibraryLink
    {
        /// <summary>
        /// Replace the content of a json file with the data passed as parameter.
        /// </summary>
        /// <param name="filePath">The path of the file to modify</param>
        /// <param name="data">The data to write in the file</param>
        public void WriteJsonFile(String filePath, Object data)
        {
            String json = JsonConvert.SerializeObject(data, Formatting.Indented);

            Directory.CreateDirectory(Environment.CurrentDirectory + "\\conf");
            File.Create(filePath).Close();
            StreamWriter writer = new StreamWriter(filePath);
            writer.Write(json);
            writer.Close();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the content of a file.
        /// </summary>
        /// <param name="filePath">The path of the file to read.</param>
        /// <returns>The content of the file or null if it cannot be read.</returns>
        public string GetFileContent(string filePath)
        {
            String content = "";

            try
            {
                StreamReader reader = new StreamReader(filePath);
                content = reader.ReadToEnd();
                reader.Close();
            }
            catch
            {
                throw new Exceptions.ConfigDataException("Le fichier " + filePath + " a été supprimé ou déplacé");
            }

            return content;
        }

        /*-------------------------------------------------------------------------*/

        public T GetJsonFilecontent<T>(string filePath) where T : class
        {
            String content = this.GetFileContent(filePath);

            T? data = JsonConvert.DeserializeObject<T>(content);

            if(data == null)
            {
                throw new Exceptions.ConfigDataException("Le contenu du fichier " + filePath + " est invalide");
            }

            return data;
        }

        /*-------------------------------------------------------------------------*/
    }
}
