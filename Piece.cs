namespace Application
{
    internal class Piece
    {
        private readonly List<List<Data.Data>> pieceData;
        private readonly List<String> measurePlans;

        /*-------------------------------------------------------------------------*/

        public Piece() 
        {
            this.pieceData = new List<List<Data.Data>>();
            this.measurePlans = new List<String>();
        }

        /*-------------------------------------------------------------------------*/

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

        public void AddMeasurePlan(String measurePlan)
        {
            this.measurePlans.Add(measurePlan);
            this.pieceData.Add(new List<Data.Data>());
        }

        /*-------------------------------------------------------------------------*/

        public void AddData(Data.Data data)
        {
            if (this.pieceData.Count == 0)
            {
                this.AddMeasurePlan("");
            }

            this.pieceData[this.pieceData.Count - 1].Add(data);
        }

        /*-------------------------------------------------------------------------*/

        public void SetValues(List<double> values)
        {
            int i = pieceData.Count - 1;
            int j = this.pieceData[i].Count - 1;

            this.pieceData[i][j].SetValues(values);
        }

        /*-------------------------------------------------------------------------*/

        public List<String> GetMeasurePlans()
        {
            return this.measurePlans;
        }

        /*-------------------------------------------------------------------------*/

        public List<List<Data.Data>> GetData()
        {
            return this.pieceData;
        }

        /*-------------------------------------------------------------------------*/
    }
}
