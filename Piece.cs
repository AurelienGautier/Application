namespace Application
{
    internal class Piece
    {
        // Il y a une liste de données pour chaque plan de mesure pour les données de la pièce
        private readonly List<List<Data.Data>> pieceData;

        private readonly List<String> measurePlans;

        /*-------------------------------------------------------------------------*/

        public Piece() 
        {
            this.pieceData = new List<List<Data.Data>>();
            this.measurePlans = new List<String>();
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
            this.pieceData.Add(new List<Data.Data>());
        }

        /*-------------------------------------------------------------------------*/

        /**
         * AddData
         * 
         * Ajoute une donnée à la pièce
         * data : Data.Data - Donnée à ajouter
         * 
         */
        public void AddData(Data.Data data)
        {
            if (this.pieceData.Count == 0)
            {
                this.AddMeasurePlan("");
            }

            this.pieceData[this.pieceData.Count - 1].Add(data);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * SetValues
         * 
         * Définit les valeurs de mesure de la pièce
         * values : List<double> - Liste des valeurs
         * 
         */
        public void SetValues(List<double> values)
        {
            int i = pieceData.Count - 1;
            int j = this.pieceData[i].Count - 1;

            this.pieceData[i][j].SetValues(values);
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
        public List<List<Data.Data>> GetData()
        {
            return this.pieceData;
        }

        /*-------------------------------------------------------------------------*/
    }
}
