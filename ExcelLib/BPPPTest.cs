namespace ExcelLib
{
    /// <summary>
    /// Один тест соотвествующий каждой строке таблицы Excel файла. Для BPPP.
    /// </summary>
    public class BPPPTest
    {
        public BPPPTest()
        {
            Index = 0;
            Input = new Contact();
            Output = new Contact();
        }

        private int _index;
        private double _min;
        private double _max;
        private double _value;
        private string _comment;
        public Contact Input;
        public Contact Output;

        public int Index
        {
            get { return _index; }
            set { _index = value; }
        }

        public double Min
        {
            get { return _min; }
            set { _min = value; }
        }

        public double Max
        {
            get { return _max; }
            set { _max = value; }
        }

        public double Value
        {
            get { return _value; }
            set { _value = value; }
        }

        public string Comment
        {
            get { return _comment; }
            set { _comment = value; }
        }
    }
}
