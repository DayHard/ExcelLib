namespace ExcelLib
{
    public class EData4
    {
        public EData4()
        {
            Index = 0;
            Input = new In();
            Output = new Out();
        }

        private int _index;
        public In Input;
        public Out Output;

        public int Index
        {
            get { return _index; }
            set { _index = value; }
        }
    }
}
