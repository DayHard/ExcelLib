using System;

namespace ExcelLib
{
    public class ExcelData
    {
        private int _index;
        private string _comment;
        public In Input;
        public Out Output;

        public ExcelData()
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
        private string _channel;
        private string _device;

        public In()
        {
            _channel = String.Empty;
            _device = String.Empty;
        }
        public string Channel
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
        private string _channel;
        private string _device;

        public Out()
        {
            Channel = String.Empty;
            Device = String.Empty;
        }

        public string Channel
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
