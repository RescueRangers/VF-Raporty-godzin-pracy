using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace DAL
{
    public static class TranslationSerialization
    {
        private static readonly string AppDataVf = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Vest-Fiber\Raporty\";
        private static readonly string FileName = "Tlumaczenia.xml";

        private static readonly string FullPath = AppDataVf + FileName;

        private static readonly XmlSerializer Serializer = new XmlSerializer(typeof(List<Translation>));

        public static void SerializeTranslations(List<Translation> translations)
        {
            Directory.CreateDirectory(AppDataVf);
            using (var stream = new FileStream(FullPath, FileMode.Create))
            {
                Serializer.Serialize(stream, translations);
            }
        }

        public static List<Translation> DeserializeTranslations()
        {
            List<Translation> translations;

            using (var stream = new FileStream(FullPath,FileMode.OpenOrCreate))
            {
                translations = (List<Translation>)Serializer.Deserialize(stream);
            }

            return translations;
        }
    }
}
