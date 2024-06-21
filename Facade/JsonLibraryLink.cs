using Newtonsoft.Json;
using System.IO;

namespace Application.Facade
{
    public static class JsonLibraryLink
    {
        /// <summary>
        /// Replace the content of a json file with the data passed as parameter.
        /// </summary>
        /// <param name="filePath">The path of the file to modify</param>
        /// <param name="data">The data to write in the file</param>
        public static void WriteJsonFile(String filePath, Object data)
        {
            String json = JsonConvert.SerializeObject(data, Formatting.Indented);

            Directory.CreateDirectory(Environment.CurrentDirectory + "\\conf");
            File.Create(filePath).Close();
            StreamWriter writer = new (filePath);
            writer.Write(json);
            writer.Close();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the content of a file.
        /// </summary>
        /// <param name="filePath">The path of the file to read.</param>
        /// <returns>The content of the specified file.</returns>
        public static string GetFileContent(string filePath)
        {
            String content;

            try
            {
                StreamReader reader = new (filePath);
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

        /// <summary>
        /// Get the content of a json file and parse it to the specified type.
        /// </summary>
        /// <typeparam name="T">The type of the data to return</typeparam>
        /// <param name="filePath">The json file with the content to get</param>
        /// <returns>The json file to parse</returns>
        /// <exception cref="Exceptions.ConfigDataException"></exception>
        public static T GetJsonFilecontent<T>(string filePath) where T : class
        {
            String content = GetFileContent(filePath);

            T? data = JsonConvert.DeserializeObject<T>(content);

            return data ?? throw new Exceptions.ConfigDataException("Le contenu du fichier " + filePath + " est invalide");
        }

        /*-------------------------------------------------------------------------*/
    }
}
