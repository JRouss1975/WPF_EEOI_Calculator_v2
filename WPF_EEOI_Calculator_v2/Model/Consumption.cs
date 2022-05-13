using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPF_EEOI_Calculator_v2
{
    [Serializable]
    public class Consumption : Observable
    {
        private FuelTypes _fuelType;
        public FuelTypes FuelType
        {
            get { return _fuelType; }
            set
            {
                if (value == _fuelType) { return; }
                _fuelType = value;
                NotifyChange("");
            }
        }

        private double _fC;
        public double FC
        {
            get { return _fC; }
            set
            {
                if (value == _fC) { return; }
                _fC = value;
                NotifyChange("");
            }
        }

        private string _remarks;
        public string Remarks
        {
            get { return _remarks; }
            set
            {
                if (value == _remarks) { return; }
                _remarks = value;
                NotifyChange("Remarks");
            }
        }

        public double Cf
        {
            get
            {
                switch (FuelType)
                {
                    case FuelTypes.DGO:
                        return 3.206;
                    case FuelTypes.LFO:
                        return 3.15104;
                    case FuelTypes.HFO:
                        return 3.1144;
                    case FuelTypes.LPG_P:
                        return 3.000;
                    case FuelTypes.LPG_B:
                        return 3.03;
                    case FuelTypes.LNG:
                        return 2.75;
                    default:
                        return 0;
                }
            }
        }

        public double Emission
        {
            get
            {
                return Cf * FC;
            }
        }
    }
}