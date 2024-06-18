namespace Application.Data
{
    public class Piece
    {
        // Il y a une liste de données pour chaque plan de mesure pour les données de la pièce
        private readonly List<MeasurePlan> measurePlans;

        private readonly Header header;

        /*-------------------------------------------------------------------------*/

        public Piece() 
        {
            this.measurePlans = new List<MeasurePlan>();

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

            foreach (var plan in this.measurePlans)
            {
                lineNb += plan.GetLinesToWriteNumber();
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
            this.measurePlans.Add(new MeasurePlan(measurePlan));
        }

        /*-------------------------------------------------------------------------*/

        /**
         * AddData
         * 
         * Ajoute une donnée à la pièce
         * data : Data.Data - Donnée à ajouter
         * 
         */
        public void AddData(Measure data)
        {
            if (this.measurePlans.Count == 0)
            {
                this.AddMeasurePlan("");
            }

            this.measurePlans[this.measurePlans.Count - 1].AddMeasure(data);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * GetMeasurePlans
         * 
         * Retourne la liste des plans de mesure utilisés pour mesurer la pièce
         * return : List<String> - Liste des plans de mesure
         * 
         */
        public List<MeasurePlan> GetMeasurePlans()
        {
            return this.measurePlans;
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
