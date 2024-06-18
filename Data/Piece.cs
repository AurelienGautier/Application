namespace Application.Data
{
    /// <summary>
    /// Represents a piece got from a source file.
    /// </summary>
    public class Piece
    {
        private readonly List<MeasurePlan> measurePlans;
        private readonly Header header;

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Initializes a new instance of the <see cref="Piece"/> class.
        /// </summary>
        public Piece()
        {
            this.measurePlans = new List<MeasurePlan>();
            this.header = new Header();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the number of lines to write in the Excel form for this piece.
        /// </summary>
        /// <returns>The number of lines to write.</returns>
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

        /// <summary>
        /// Adds a measure plan to the piece.
        /// </summary>
        /// <param name="measurePlan">The measure plan to add.</param>
        public void AddMeasurePlan(string measurePlan)
        {
            this.measurePlans.Add(new MeasurePlan(measurePlan));
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Adds data to the piece.
        /// </summary>
        /// <param name="data">The data to add.</param>
        public void AddData(Measure data)
        {
            if (this.measurePlans.Count == 0)
            {
                this.AddMeasurePlan("");
            }

            this.measurePlans[this.measurePlans.Count - 1].AddMeasure(data);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the list of measure plans used to measure the piece.
        /// </summary>
        /// <returns>The list of measure plans.</returns>
        public List<MeasurePlan> GetMeasurePlans()
        {
            return this.measurePlans;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Creates the header from a raw text.
        /// </summary>
        /// <param name="text">The raw header text.</param>
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

        /// <summary>
        /// Gets the header of the piece.
        /// </summary>
        /// <returns>The header.</returns>
        public Header GetHeader()
        {
            return this.header;
        }

        /*-------------------------------------------------------------------------*/
    }
}
