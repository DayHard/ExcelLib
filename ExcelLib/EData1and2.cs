using System;

namespace ExcelLib
{
    public class EData1And2
    {
        private int _index;
        private string _comment;
        public In Input;
        public Out Output;

        public EData1And2()
        {
            _index = 0;
            _comment = String.Empty;
            In input = new In();
            Out output = new Out();
            Input = input;
            Output = output;
        }
        public int Index
        {
            get { return _index; }
            set { _index = value; }
        }
        public string Comment
        {
            get { return _comment; }
            set { _comment = value; }
        }
    }

    public class In
    {
        private int _channel;
        private string _device;

        public In()
        {
            _channel = 0;
            _device = String.Empty;
        }
        public int Channel
        {
            get { return _channel; }
            set { _channel = value; }
        }

        public string Device
        {
            get { return _device; }
            set { _device = value; }
        }
    }

    public class Out
    {
        private int _channel;
        private string _device;

        public Out()
        {
            Channel = 0;
            Device = String.Empty;
        }

        public int Channel
        {
            get { return _channel; }
            set { _channel = value; }
        }

        public string Device
        {
            get { return _device; }
            set { _device = value; }
        }
    }
}
