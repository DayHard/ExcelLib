using System;
using System.Diagnostics.CodeAnalysis;

namespace ExcelLib
{
    public enum control
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
    public enum multMode
    {
        DiodeTest,
        Resistance,
        [SuppressMessage("ReSharper", "InconsistentNaming")]
        DCVoltage
    }
    public class EData3
    {
        public multMode MultMode;
        public control Control;
        public VoltSource VoltSource;
        public EData3()
        {
            _index = 0;
            _currSource = 0;
            _comment = String.Empty;
            _valMin = 0;
            _valMax = 0;
            _valUnit = String.Empty;
            Input = new In[7];
            VoltSource = new VoltSource();
        }
        public EData3(int inputCount)
        {
            _index = 0;
            _currSource = 0;
            _comment = String.Empty;
            _valMin = 0;
            _valMax = 0;
            _valUnit = String.Empty;
            Input = new In[inputCount];
            VoltSource = new VoltSource();
        }

        private int _index;
        public string gsdfgsdfsdfsdfsdfsdf;
        private short _currSource;
        //private short _voltSource;
        private string _comment;
        private int _valMin;
        private int _valMax;
        private string _valUnit;
        public In[] Input;
        public int Index
        {
            get { return _index; }
            set { _index = value; }
        }

        public short CurrSource
        {
            get { return _currSource; }
            set { _currSource = value; }
        }

        public string Comment
        {
            get { return _comment; }
            set { _comment = value; }
        }

        public int ValMin
        {
            get { return _valMin; }
            set { _valMin = value; }
        }

        public int ValMax
        {
            get { return _valMax; }
            set { _valMax = value; }
        }

        public string ValUnit
        {
            get { return _valUnit; }
            set { _valUnit = value; }
        }
    }

    public class VoltSource
    {
        private int _power1;
        private int _power2;

        public VoltSource()
        {
            _power1 = 0;
            _power2 = 0;
        }
        public int Power1
        {
            get { return _power1; }
            set { _power1 = value; }
        }
        public int Power2
        {
            get { return _power2; }
            set { _power2 = value; }
        }
    }
}
