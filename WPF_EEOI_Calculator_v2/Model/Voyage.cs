using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace WPF_EEOI_Calculator_v2
{
    [Serializable]
    public class Voyage : Observable
    {
        private string _id;
        public string Id
        {
            get { return _id; }
            set
            {
                if (value == _id) { return; }
                _id = value;
                NotifyChange("Id");
            }
        }

        private bool _isEnabled = true;
        public bool IsEnabled
        {
            get { return _isEnabled; }
            set
            {
                if (value == _isEnabled) { return; }
                _isEnabled = value;
                NotifyChange("IsEnabled");
            }
        }

        private string _departurePort;
        public string DeparturePort
        {
            get { return _departurePort; }
            set
            {
                if (value == _departurePort) { return; }
                _departurePort = value;
                NotifyChange("DeparturePort");
            }
        }

        [XmlIgnore]
        public DateTime CompletedDate { get; set; } = DateTime.Now;

        [XmlElement("CompletedDate")]
        public string _CompletedDateString
        {
            get { return CompletedDate.ToString("dd/MM/yyyy"); }
            set
            {
                if (value == _CompletedDateString) { return; }
                CompletedDate = DateTime.Parse(value);
                NotifyChange("CompletedDate");
            }
        }

        private VoyageTypes _voyageType;
        public VoyageTypes VoyageType
        {
            get { return _voyageType; }
            set
            {
                if (value == _voyageType) { return; }
                _voyageType = value;
                if (value == VoyageTypes.Ballast)
                    CargoMass = 0;
                NotifyChange("VoyageType");
            }
        }

        private double _cargoMass;
        public double CargoMass
        {
            get { return _cargoMass; }
            set
            {
                if (value == _cargoMass) { return; }
                if (VoyageType == VoyageTypes.Cargo)
                    _cargoMass = value;
                else
                    _cargoMass = 0;
                NotifyChange("CargoMass");
            }
        }

        private double _distance;
        public double Distance
        {
            get { return _distance; }
            set
            {
                if (value == _distance) { return; }
                _distance = value;
                NotifyChange("Distance");
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

        public ObservableCollection<Consumption> Consumptions { get; set; } = new ObservableCollection<Consumption>();

        public double VoyageEmissions
        {
            get
            {
                return Consumptions.Sum(x => x.Emission);
            }
        }

        public double? VoyageEEOI
        {
            get
            {
                if (VoyageType == VoyageTypes.Cargo)
                    return Math.Pow(10, 6) * VoyageEmissions / (CargoMass * Distance);
                return null;
            }
        }

        public override bool Equals(object obj)
        {
            if (obj == null) return false;

            Voyage voyage = obj as Voyage;

            return (voyage != null)
                && (this.DeparturePort == voyage.DeparturePort)
                && (this.CompletedDate == voyage.CompletedDate);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

    }
}