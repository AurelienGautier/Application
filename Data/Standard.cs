namespace Application.Data
{
    public class Standard
    {
        public String Code { get; set; }
        public String Name { get; set; }
        public String Raccordement { get; set; }
        public String Validity { get; set; }

        public Standard(String code, String name, String raccordement, String validity)
        {
            this.Code = code;
            this.Name = name;
            this.Raccordement = raccordement;
            this.Validity = validity;
        }
    }
}
