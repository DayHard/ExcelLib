﻿using System;

namespace ExcelLib
{
    public class DAQTest
    {
        private int _index;
        private string _comment;
        public Contact Input;
        public Contact Output;
        private string _result;
        private string _value;
        private string _errordescription;

        public DAQTest()
        {
            _index = 0;
            _comment = string.Empty;
            _errordescription = string.Empty;
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

        public string Value
        {
            get { return _value; }
            set { _value = value; }
        }

        public string ErrorDescription
        {
            get { return _errordescription; }
            set { _errordescription = value; }
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
}
