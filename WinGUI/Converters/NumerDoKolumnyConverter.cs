using System;
using System.Globalization;
using System.Windows.Data;

namespace WinGUI.Converters
{
    class NumerDoKolumnyConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            var numerKolumny = (int)value;
            var nazwaKolumny = string.Empty;

            while (numerKolumny > 0)
            {
                var modulo = (numerKolumny - 1) % 26;
                nazwaKolumny = System.Convert.ToChar(65 + modulo).ToString() + nazwaKolumny;
                numerKolumny = ((numerKolumny - modulo) / 26);
            }

            return nazwaKolumny;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));

            var nazwaKolumny = value.ToString().ToUpperInvariant();
            var suma = 0;

            foreach (var litera in nazwaKolumny)
            {
                suma *= 26;
                suma += (litera - 'A' + 1);
            }

            return suma;
        }
    }
}
