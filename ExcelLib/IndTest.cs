using System;
using System.Diagnostics.CodeAnalysis;

namespace ExcelLib
{
    public enum Control
    {
        Напряжение,
        Сопротивление,
        Индикация,
        ПадениеНапряженияБк,
        ПадениеНапряженияБэ,
        ПадениеНапряженияКб,
        ПадениеНапряженияЭб,
        ПадениеНапряженияЭк,
    }
    public enum MultMode
    {
        DiodeTest,
        Resistance,
        [SuppressMessage("ReSharper", "InconsistentNaming")]
        DCVoltage
    }
    public class IndTest
    {
        public IndTest()
        {
            _index = 0;
            _comment = string.Empty;
            _valMin = 0;
            _valMax = 0;
            _valUnit = string.Empty;
            _errordescription = string.Empty;
            Input = new Contact[7];
            VoltSupply = new VoltSupply();
        }
        public IndTest(int inputCount, int inputCount2)
        {
            _index = 0;
            _comment = string.Empty;
            _valMin = 0;
            _valMax = 0;
            _valUnit = string.Empty;
            _errordescription = string.Empty;
            VoltSupply = new VoltSupply();
            Input = new Contact[inputCount];
            for (int i = 0; i < Input.Length; i++)
            {
                Input[i] = new Contact();
            }
            CurrSource = new currSource[inputCount2];
            for (int i = 0; i < CurrSource.Length; i++)
            {
                CurrSource[i] = new currSource();
            }
        }

        private int _index;
        private string _comment;
        private double _valMin;
        private double _valMax;
        private string _valUnit;
        private string _errordescription;
        public Contact[] Input;
        public MultMode MultMode;
        public Control Control;
        public VoltSupply VoltSupply;
        public currSource[] CurrSource;

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

        public double ValMin
        {
            get { return _valMin; }
            set { _valMin = value; }
        }

        public double ValMax
        {
            get { return _valMax; }
            set { _valMax = value; }
        }

        public string ValUnit
        {
            get { return _valUnit; }
            set { _valUnit = value; }
        }

        public string ErrorDescription
        {
            get { return _errordescription; }
            set { _errordescription = value; }
        }
    }

    public class VoltSupply
    {
        private int _v1;
        private int _v2;

        public VoltSupply()
        {
            _v1 = 0;
            _v2 = 0;
        }
        public int V1
        {
            get { return _v1; }
            set { _v1 = value; }
        }
        public int V2
        {
            get { return _v2; }
            set { _v2 = value; }
        }
    }

    public class currSource
    {
        private int _currSource;

        public currSource()
        {
            _currSource = 0;
        }

        public int CurrSource
        {
            get { return _currSource; }
            set { _currSource = value; }
        }
    }
}
