using System;
using System.Collections.Generic;
using System.Xml;

namespace VF_Raporty_Godzin_Pracy
{
    public static class Tlumacz
    {
        public static Dictionary<string, string> LadujTlumaczenia()
        {
            var dictionaryTlumaczenia = new Dictionary<string, string>();
            var tlumaczenia = AppDomain.CurrentDomain.BaseDirectory + $@"\Assets\TlumaczoneNaglowki.xml";
            var xmlDokument = new XmlDocument();
            xmlDokument.Load(tlumaczenia);
            var nodes = xmlDokument.DocumentElement.SelectNodes("/naglowki/naglowek");

            foreach (XmlNode node in nodes)
            {
                dictionaryTlumaczenia.Add(node.SelectSingleNode("klucz").InnerText.ToLower(), node.SelectSingleNode("wartosc").InnerText);
            }
            return dictionaryTlumaczenia;
        }

        public static void EdytujTlumaczenia(Dictionary<string, string> naglowki)
        {
            var tlumaczenia = AppDomain.CurrentDomain.BaseDirectory + $@"\Assets\TlumaczoneNaglowki.xml";
            var xmlDokument = new XmlDocument();
            xmlDokument.Load(tlumaczenia);
            var nodes = xmlDokument.DocumentElement.SelectNodes("/naglowki/naglowek");

            foreach (XmlNode node in nodes)
            {
                foreach (var naglowek in naglowki)
                {
                    if (node.SelectSingleNode("klucz").InnerText.ToLower() == naglowek.Key)
                    {
                        node.SelectSingleNode("wartosc").InnerText = naglowek.Value;
                    }
                }
            }
            xmlDokument.Save(tlumaczenia);
        }
    }
}
