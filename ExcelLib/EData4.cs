namespace ExcelLib
{
    public class EData4
    {
        public EData4()
        {
            Index = 0;
            Input = new Contact();
            Output = new Contact();
        }

        private int _index;
        public Contact Input;
        public Contact Output;

        public int Index
        {
            get { return _index; }
            set { _index = value; }
        }
    }
}
