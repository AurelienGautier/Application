namespace Application.Data
{
    public class Piece
    {
        // Il y a une liste de données pour chaque plan de mesure pour les données de la pièce
        private readonly List<List<Data>> pieceData;

        private readonly List<String> measurePlans;

        private readonly Header header;

        /*-------------------------------------------------------------------------*/

        public Piece() 
        {
            this.pieceData = new List<List<Data>>();
            this.measurePlans = new List<String>();

            // Création de l'en-tête vide
            this.header = new Header();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * GetLinesToWriteNumber
         * 
         * Retourne le nombre de lignes à prévoir dans le formulaire excel pour cette pièce
         * return : int - Nombre de lignes à écrire
         * 
         */
        public int GetLinesToWriteNumber()
        {
            int lineNb = 0;

            for(int i = 0; i < this.pieceData.Count; i++) 
            {
                lineNb++;

                lineNb += this.pieceData[i].Count;
            }

            return lineNb;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * AddMeasurePlan
         * 
         * Ajoute un plan de mesure à la pièce
         * measurePlan : String - Plan de mesure à ajouter
         * 
         */
        public void AddMeasurePlan(String measurePlan)
        {
            this.measurePlans.Add(measurePlan);
            this.pieceData.Add(new List<Data>());
        }

        /*-------------------------------------------------------------------------*/

        /**
         * AddData
         * 
         * Ajoute une donnée à la pièce
         * data : Data.Data - Donnée à ajouter
         * 
         */
        public void AddData(Data data)
        {
            if (this.pieceData.Count == 0)
            {
                this.AddMeasurePlan("");
            }

            this.pieceData[this.pieceData.Count - 1].Add(data);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * GetMeasurePlans
         * 
         * Retourne la liste des plans de mesure utilisés pour mesurer la pièce
         * return : List<String> - Liste des plans de mesure
         * 
         */
        public List<String> GetMeasurePlans()
        {
            return this.measurePlans;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * GetData
         * 
         * Retourne la liste des valeurs de mesure de la pièce
         * return : List<List<Data.Data>> - Liste des données
         * 
         */
        public List<List<Data>> GetData()
        {
            return this.pieceData;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Crée l'en-tête à partir d'une chaîne de caractères
         * text : String - L'en-tête récupérée de manière brute depuis le fichier à analyser
         *
         */
        public void CreateHeader(string text)
        {
            string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            Dictionary<string, string> rawHeader = new Dictionary<string, string>();

            foreach (var line in lines)
            {
                string[] parts = line.Split(new[] { ':' }, 3);

                string key = parts[0].Trim();
                string value = parts[2].Trim();

                rawHeader[key] = value;
            }

            this.header.FillHeader(rawHeader);
        }

        /*-------------------------------------------------------------------------*/

        public Header GetHeader()
        {
            return this.header;
        }

        /*-------------------------------------------------------------------------*/
    }
}
