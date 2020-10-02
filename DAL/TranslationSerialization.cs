using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace DAL
{
    public class TranslationSerialization
    {
        private readonly string AppDataVf;
        private readonly string FileName = "Tlumaczenia.txt";

        private string FullPath { get => AppDataVf + FileName;}

        private readonly JsonSerializer Serializer = new JsonSerializer();

        public TranslationSerialization(bool customTranslationDirectory, string serializationDirectory = null)
        {
            if (customTranslationDirectory && !string.IsNullOrWhiteSpace(serializationDirectory))
            {
                AppDataVf = serializationDirectory;
            }
            else
            {
                AppDataVf = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Vest-Fiber\Raporty\";
            }
        }

        public void SerializeTranslations(List<Translation> translations)
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

        public List<Translation> DeserializeTranslations()
        {
            if (!File.Exists(FullPath))
            {
                return new List<Translation>();
            }
            var json = File.ReadAllText(FullPath);

            return JsonConvert.DeserializeObject<List<Translation>>(json);
        }
    }
}