using System;
using System.Globalization;
using System.Windows.Data;

namespace WPF_EEOI_Calculator_v2
{
    public class UnitConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string checkValue = value.ToString();
            string paramValue = parameter.ToString();

            if (checkValue.Equals("Metric"))
            {
                if (paramValue.Equals("Speed"))
                    return "[kmh]";
                return "[m]";
            }
            else
            {
                if (paramValue.Equals("Speed"))
                    return "[mph]";
                return "[ft]";
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class VesselTypeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {

            if (value is VesselTypes)
            {
                VesselTypes checkValue = (VesselTypes)value;


                switch (checkValue)
                {
                    case VesselTypes.DryCargo:
                        return "Dry Cargo";
                    case VesselTypes.Tanker:
                        return "Tanker";
                    case VesselTypes.GasTanker:
                        return "Gas Tanker";
                    case VesselTypes.Container:
                        return "Containership";
                    case VesselTypes.RoRoShip:
                        return "Ro-Ro Cargo Ship";
                    case VesselTypes.GeneralCargo:
                        return "General Cargo Ship";
                    case VesselTypes.Passenger:
                        return "Passenger or ro-ro passenger ship";
                    default:
                        return null;
                }
            }
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
