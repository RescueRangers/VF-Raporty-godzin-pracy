using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace DAL
{
    public class TranslationSerialization
    {
        private readonly string _appDataVf = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Vest-Fiber\Raporty\";
        private const string FileName = "Tlumaczenia.xml";

        private readonly string _fullPath;

        private readonly XmlSerializer _serializer = new XmlSerializer(typeof(List<Translation>));

        public TranslationSerialization()
        {
            _fullPath = _appDataVf + FileName;
            Directory.CreateDirectory(_appDataVf);
        }

        public void SerializeTranslations(List<Translation> translations)
        {
            using (var stream = new FileStream(_fullPath, FileMode.Create))
            {
                _serializer.Serialize(stream, translations);
            }
        }

        public List<Translation> DeserializeTranslations()
        {
            List<Translation> translations;

            using (var stream = new FileStream(_fullPath,FileMode.OpenOrCreate))
            {
                translations = (List<Translation>)_serializer.Deserialize(stream);
            }

            return translations;
        }
    }
}
