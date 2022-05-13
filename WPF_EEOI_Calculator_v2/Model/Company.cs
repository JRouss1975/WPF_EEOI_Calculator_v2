using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPF_EEOI_Calculator_v2
{
    [Serializable]
    public class Company : Observable
    {
        public string CompanyName { get; set; }

        private string _sEEMPCycle;

        public string SEEMPCycle
        {
            get { return _sEEMPCycle; }
            set
            {
                if (_sEEMPCycle == value) { return; }
                _sEEMPCycle = value;
                NotifyChange("SEEMPCycle");
            }
        }

        private ObservableCollection<Vessel> _vessels = new ObservableCollection<Vessel>();
        public ObservableCollection<Vessel> Vessels
        {
            get { return _vessels; }
            set
            {
                if (_vessels == value) { return; }
                _vessels = value;
                NotifyChange("Vessels");
            }
        }
    }
}
