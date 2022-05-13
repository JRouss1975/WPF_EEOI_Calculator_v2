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
    public enum VesselTypes
    {
        [Description("Dry Cargo Carrier")]
        DryCargo,
        Tanker,
        [Description("Gas Tanker")]
        GasTanker,
        [Description("Containership")]
        Container,
        [Description("Ro-Ro Cargo Ship")]
        RoRoShip,
        [Description("General Cargo Ship")]
        GeneralCargo,
        [Description("Passenger or ro-ro passenger ship")]
        Passenger
    }
}
