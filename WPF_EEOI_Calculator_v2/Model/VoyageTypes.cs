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
    public enum VoyageTypes
    {
        Cargo,
        Ballast
    }
}
