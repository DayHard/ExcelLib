using System;

namespace ExcelLib
{
    public class DAQTest
    {
        private int _index;
        private string _comment;
        public Contact Input;
        public Contact Output;
        private string _result;

        public DAQTest()
        {
            _index = 0;
            _comment = String.Empty;
            Input = new Contact();
            Output = new Contact();
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

        public string Result
        {
            get { return _result; }
            set { _result = value; }
        }
    }

    public class Contact
    {
        private int _channel;
        private string _device;

        public Contact()
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

    //public class Out
    //{
    //    private int _channel;
    //    private string _device;

    //    public Out()
    //    {
    //        Channel = 0;
    //        Device = String.Empty;
    //    }

    //    public int Channel
    //    {
    //        get { return _channel; }
    //        set { _channel = value; }
    //    }

    //    public string Device
    //    {
    //        get { return _device; }
    //        set { _device = value; }
    //    }
    //}
}
