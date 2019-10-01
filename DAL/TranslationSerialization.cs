using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace DAL
{
    public static class TranslationSerialization
    {
        private static readonly string AppDataVf = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Vest-Fiber\Raporty\";
        private static readonly string FileName = "Tlumaczenia.txt";

        private static readonly string FullPath = AppDataVf + FileName;

        private static readonly JsonSerializer Serializer = new JsonSerializer();

        public static void SerializeTranslations(List<Translation> translations)
        {
            Directory.CreateDirectory(AppDataVf);

            using (var streamWriter = new StreamWriter(FullPath))
            {
                using (var writer = new JsonTextWriter(streamWriter))
                {
                    Serializer.Serialize(writer, translations);
                }
            }
        }

        public static List<Translation> DeserializeTranslations()
        {
            List<Translation> translations;

            using (var streamReader = new JsonTextReader(new StringReader(FullPath)))
            {
                translations = Serializer.Deserialize<List<Translation>>(streamReader);
            }

            return translations;
        }
    }
}
