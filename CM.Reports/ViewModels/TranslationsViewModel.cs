﻿using System.Collections.ObjectModel;
using System.Linq;
using Caliburn.Micro;
using CM.Reports.Properties;
using DAL;

namespace CM.Reports.ViewModels
{
    internal class TranslationsViewModel : PropertyChangedBase
    {
        private ObservableCollection<Translation> _translatedHeaders;
        private ObservableCollection<Translation> _headersToTranslate = new ObservableCollection<Translation>();
        private Translation _selectedTranslation;
        private TranslationSerialization _translationSerialization;

        public ObservableCollection<Translation> TranslatedHeaders
        {
            get => _translatedHeaders;
            set
            {
                if (Equals(value, _translatedHeaders)) return;
                _translatedHeaders = value;
                NotifyOfPropertyChange(() => TranslatedHeaders);
            }
        }

        public ObservableCollection<Translation> HeadersToTranslate
        {
            get => _headersToTranslate;
            set
            {
                if (Equals(value, _headersToTranslate)) return;
                _headersToTranslate = value;
                NotifyOfPropertyChange(() => HeadersToTranslate);
                //NotifyOfPropertyChange(() => CanTranslate);
            }
        }

        public Translation SelectedTranslation
        {
            get => _selectedTranslation;
            set
            {
                if (Equals(value, _selectedTranslation)) return;
                _selectedTranslation = value;
                NotifyOfPropertyChange(() => SelectedTranslation);
                NotifyOfPropertyChange(() => CanDeleteTranslation);
            }
        }

        public bool CanDeleteTranslation => SelectedTranslation != null;
        //public bool CanTranslate => HeadersToTranslate != null && HeadersToTranslate.Any(h => !string.IsNullOrWhiteSpace(h.Translated));

        public TranslationsViewModel(TranslationSerialization translationSerialization)
        {
            //var isCustomDirectory = Settings.Default.UseCustomTranslationsDirectory;
            //var customDirectory = Settings.Default.CustomTranslationsDirectory;
            _translationSerialization = translationSerialization;
            TranslatedHeaders = new ObservableCollection<Translation>(_translationSerialization.DeserializeTranslations());
        }

        public void Translate()
        {
            foreach (var translation in HeadersToTranslate.Where(h => !string.IsNullOrWhiteSpace(h.Translated)).ToList())
            {
                TranslatedHeaders.Add(translation);
                HeadersToTranslate.Remove(translation);
            }
            _translationSerialization.SerializeTranslations(TranslatedHeaders.ToList());
        }

        public void DeleteTranslation()
        {
            HeadersToTranslate.Add(SelectedTranslation);
            var delete = TranslatedHeaders.Remove(SelectedTranslation);

            if (delete)
            {
                _translationSerialization.SerializeTranslations(TranslatedHeaders.ToList());
            }
        }
    }
}