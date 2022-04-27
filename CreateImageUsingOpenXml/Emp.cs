namespace CreateImageUsingOpenXml
{
    class Emp
    {
       public string Name { get; set; }
        public string Age { get; set; }
        public string Gender { get; set; }

        public override bool Equals(object obj)
        {
            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public override string ToString()
        {
            return $" Name = {Name}, Age = {Age}, Gender {Gender}";
        }
    }
}
