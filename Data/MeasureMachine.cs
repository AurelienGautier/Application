namespace Application.Data
{
    public class MeasureMachine
    {
        public String Name { get; }
        public Parser.Parser Parser { get; }
        public List<Form> PossiblesForms { get; set; }

        public MeasureMachine(string name, Parser.Parser parser)
        {
            this.Name = name;
            this.Parser = parser;
            this.PossiblesForms = new List<Form>();
        }

        public void setPossibleForms(List<Form> forms)
        {
            this.PossiblesForms = forms;
        }
    }
}
