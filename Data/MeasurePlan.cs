using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Data
{
    public class MeasurePlan
    {
        readonly String name;
        readonly List<Measure> measures;

        public MeasurePlan(String measurePlanName)
        {
            name = measurePlanName;
            measures = new List<Measure>();
        }

        public void AddMeasure(Measure measure) 
        { 
            measures.Add(measure); 
        }

        public int GetLinesToWriteNumber() { return measures.Count + 1; }

        public String GetName() { return this.name; }

        public List<Measure> GetMeasures() { return this.measures; }
    }
}
