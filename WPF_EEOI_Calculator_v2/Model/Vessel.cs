using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using OxyPlot;

namespace WPF_EEOI_Calculator_v2
{
    [Serializable]
    public class Vessel : Observable
    {
        private string _vesselName;
        public string VesselName
        {
            get { return _vesselName; }
            set
            {
                if (value == _vesselName) { return; }
                _vesselName = value;
                NotifyChange("VesselName");
            }
        }

        private VesselTypes _vesselType;
        public VesselTypes VesselType
        {
            get { return _vesselType; }
            set
            {
                if (value == _vesselType) { return; }
                _vesselType = value;
                NotifyChange("VesselType");
            }
        }

        public string Flag { get; set; }

        private string _iMO;
        public string IMO
        {
            get { return _iMO; }
            set
            {
                if (value == _iMO) { return; }
                _iMO = value;
                NotifyChange("IMO");
            }
        }

        public string Tonnage { get; set; }

        private bool _isSelected;
        public bool IsSelected
        {
            get
            {
                return _isSelected;
            }
            set
            {
                if (value == _isSelected) { return; }
                _isSelected = value;
                NotifyChange("IsSelected");
            }
        }

        public ObservableCollection<Voyage> Voyages { get; set; } = new ObservableCollection<Voyage>();

        [XmlIgnoreAttribute]
        public double VesselEEOI
        {
            get
            {
                double emmisions = 0;
                double cargoDistanceProduct = 0;
                foreach (Voyage v in Voyages)
                {
                    if (v.IsEnabled)
                    {
                        emmisions += v.VoyageEmissions;
                        cargoDistanceProduct += (v.CargoMass * v.Distance);
                    }
                }
                _vesselEEOI = Math.Pow(10, 6) * (emmisions / cargoDistanceProduct);
                return _vesselEEOI;
            }
            set
            {
                if (value == _vesselEEOI) { return; }
                _vesselEEOI = value;
                NotifyChange("VesselEEOI");
            }
        }
        [NonSerialized]
        private double _vesselEEOI;

        [XmlIgnoreAttribute]
        public double VesselEmissions
        {
            get
            {
                double emmisions = 0;
                foreach (Voyage v in Voyages)
                {
                    if (v.IsEnabled)
                    {
                        emmisions = emmisions + v.VoyageEmissions;
                    }
                }
                _vesselEmissions = emmisions;
                return _vesselEmissions;
            }
            set
            {
                if (value == _vesselEmissions) { return; }
                _vesselEmissions = value;
                NotifyChange("VesselEmissions");
            }

        }
        [NonSerialized]
        private double _vesselEmissions;

        //List of EEOIs for enabled voyages  
        [XmlIgnoreAttribute]
        public ObservableCollection<DataPoint> VoyagesEEOIs
        {
            get
            {
                _voyagesEEOIs.Clear();
                int n = 1;
                foreach (Voyage v in Voyages)
                {
                    if (v.IsEnabled)
                    {
                        if (v.VoyageEEOI.HasValue)
                        {
                            DataPoint p = new DataPoint(n++, v.VoyageEEOI.Value);
                            _voyagesEEOIs.Add(p);
                        }
                    }
                }
                return _voyagesEEOIs;
            }
            set
            {
                if (value == _voyagesEEOIs) { return; }
                _voyagesEEOIs = value;
                NotifyChange("VoyagesEEOIs");
            }
        }
        [NonSerialized]
        private ObservableCollection<DataPoint> _voyagesEEOIs = new ObservableCollection<DataPoint>();

        [XmlIgnoreAttribute]
        public ObservableCollection<DataPoint> VoyagesEEOIsRA
        {
            get
            {
                _voyagesEEOIsRA.Clear();

                List<double> lstEEOIs = new List<double>();

                foreach (Voyage v in Voyages)
                {
                    if (v.IsEnabled)
                    {
                        if (v.VoyageEEOI.HasValue)
                        {
                            lstEEOIs.Add(v.VoyageEEOI.Value);
                        }
                    }
                }

                if (lstEEOIs.Count > PeriodLength)
                {
                    var temp = Enumerable.Range(0, (lstEEOIs.Count + 1) - PeriodLength)
                               .Select(x => lstEEOIs.Skip(x).Take(PeriodLength).Average())
                               .ToList();
                    int n = 1;
                    foreach (double d in temp)
                    {
                        DataPoint p = new DataPoint(n++, d);
                        _voyagesEEOIsRA.Add(p);
                    }
                }
                return _voyagesEEOIsRA;
            }
            set
            {
                if (value == _voyagesEEOIsRA) { return; }
                _voyagesEEOIsRA = value;
                NotifyChange("VoyagesEEOIsRA");
            }
        }
        [NonSerialized]
        private ObservableCollection<DataPoint> _voyagesEEOIsRA = new ObservableCollection<DataPoint>();

        [XmlIgnoreAttribute]
        public static int PeriodLength
        {
            get { return _periodLength; }
            set
            {
                if (value == _periodLength) { return; }
                if (value < 1)
                    _periodLength = 1;
                else
                    _periodLength = value;
            }
        }
        [NonSerialized]
        private static int _periodLength = 3;

        [XmlIgnoreAttribute]
        public int NumberOfEnabledVoyages
        {
            get
            {
                if (Voyages.Count > 0)
                {
                    _numberOfEnabledVoyages = Voyages.Count(x => x.IsEnabled);
                }
                return _numberOfEnabledVoyages;
            }
            set
            {
                if (value == _numberOfEnabledVoyages) { return; }
                _numberOfEnabledVoyages = value;
                NotifyChange("NumberOfEnabledVoyages");
            }
        }
        [NonSerialized]
        private int _numberOfEnabledVoyages;

        private bool _hasReg;
        public bool HasReg
        {
            get { return _hasReg; }
            set
            {
                if (value == _hasReg) { return; }
                _hasReg = value;
                NotifyChange("HasReg");
            }
        }
    }

    internal class VesselIMOComparer : IEqualityComparer<Vessel>
    {
        public bool Equals(Vessel x, Vessel y)
        {
            if (string.Equals(x.IMO, y.IMO, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            return false;
        }

        public int GetHashCode(Vessel obj)
        {
            return obj.IMO.GetHashCode();
        }
    }


}
