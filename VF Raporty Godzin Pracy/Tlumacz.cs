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

        public static void UsunTlumaczenia(Dictionary<string, string> naglowki)
        {
            var tlumaczenia = AppDomain.CurrentDomain.BaseDirectory + $@"\Assets\TlumaczoneNaglowki.xml";
            var xmlDokument = new XmlDocument();
            xmlDokument.Load(tlumaczenia);
            var nodes = xmlDokument.DocumentElement.SelectNodes("/naglowki/naglowek");

            foreach (var naglowek in naglowki)
            {
                var usunTlumaczenie = xmlDokument.SelectSingleNode($"/naglowki/naglowek/wartosc[text()='{naglowek.Value}']");
                usunTlumaczenie.ParentNode.ParentNode.RemoveChild(usunTlumaczenie.ParentNode);
            }

            xmlDokument.Save(tlumaczenia);
        }

        public static void DodajTlumaczenia(Dictionary<string, string> tlumaczoneNaglowki)
        {
            var tlumaczenia = AppDomain.CurrentDomain.BaseDirectory + $@"\Assets\TlumaczoneNaglowki.xml";
            var xmlDokument = new XmlDocument();
            xmlDokument.Load(tlumaczenia);
            var xmlNamespace = xmlDokument.DocumentElement.NamespaceURI;

            var node = xmlDokument.DocumentElement.SelectSingleNode("/naglowki");

            foreach (var tlumaczenie in tlumaczoneNaglowki)
            {
                var naglowek = xmlDokument.CreateNode(XmlNodeType.Element, "naglowek", xmlNamespace);

                var klucz = xmlDokument.CreateNode(XmlNodeType.Element, "klucz", xmlNamespace);
                klucz.InnerText = tlumaczenie.Key;
                naglowek.AppendChild(klucz);

                var wartosc = xmlDokument.CreateNode(XmlNodeType.Element, "wartosc", xmlNamespace);
                wartosc.InnerText = tlumaczenie.Value;
                naglowek.AppendChild(wartosc);

                node.AppendChild(naglowek);
            }
            xmlDokument.Save(tlumaczenia);
        }
    }
}
