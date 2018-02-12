using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace VF_Raporty_Godzin_Pracy
{
    public class SerializacjaTlumaczen
    {
        private readonly string _appDataVF = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Vest-Fiber\Raporty\";
        private readonly string _nazwaPliku = "Tlumaczenia.xml";

        private readonly string _pelnaSciezka;

        private readonly XmlSerializer _serializer = new XmlSerializer(typeof(List<Tlumaczenie>));

        public SerializacjaTlumaczen()
        {
            _pelnaSciezka = _appDataVF + _nazwaPliku;
            Directory.CreateDirectory(_appDataVF);
        }

        public void SerializujTlumaczenia(List<Tlumaczenie> tlumaczenia)
        {
            using (FileStream strumien = new FileStream(_pelnaSciezka, FileMode.Create))
            {
                _serializer.Serialize(strumien, tlumaczenia);
            }
        }

        public List<Tlumaczenie> DeserializujTlumaczenia()
        {
            List<Tlumaczenie> tlumaczenia;

            using (FileStream strumien = new FileStream(_pelnaSciezka,FileMode.OpenOrCreate))
            {
                tlumaczenia = (List<Tlumaczenie>)_serializer.Deserialize(strumien);
            }

            return tlumaczenia;
        }
    }
}
