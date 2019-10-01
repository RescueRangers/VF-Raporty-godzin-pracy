using System;
using System.Globalization;
using System.Windows.Data;

namespace CM.Reports.Converters
{
    class HalfConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double dValue)
            {
                return dValue / 2 -25;
            }

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double dValue)
            {
                return dValue * 2;
            }

            return null;
        }
    }
}
