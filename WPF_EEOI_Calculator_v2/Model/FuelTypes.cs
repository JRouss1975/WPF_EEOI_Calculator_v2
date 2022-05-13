using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPF_EEOI_Calculator_v2
{
    [Serializable]
    [TypeConverter(typeof(EnumDescriptionTypeConverter))]
    public enum FuelTypes
    {
        [Description("Diesel/Gas Oil")]
        DGO,
        [Description("Light Fuel Oil (LFO)")]
        LFO,
        [Description("Heavy Fuel Oil (HFO)")]
        HFO,
        [Description("LPG Propane")]
        LPG_P,
        [Description("LPG Butane")]
        LPG_B,
        [Description("Liq. Natural Gas")]
        LNG
    }
}
