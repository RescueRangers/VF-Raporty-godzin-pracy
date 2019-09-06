﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Caliburn.Micro;
using DAL;

namespace CM.Reports.ViewModels
{
    class TranslationsViewModel : PropertyChangedBase
    {
        private ObservableCollection<Translation> _translatedHeaders;
        private ObservableCollection<Translation> _headersToTranslate;
        private Translation _selectedTranslation;

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
            }
        }

        public bool CanDeleteTranslation => SelectedTranslation != null;
        public bool CanTranslate => HeadersToTranslate != null && HeadersToTranslate.Any(h => !string.IsNullOrWhiteSpace(h.Translated));

        public TranslationsViewModel()
        {
            TranslatedHeaders = new ObservableCollection<Translation>(TranslationSerialization.DeserializeTranslations());
        }

        public void Translate()
        {
            foreach (var translation in HeadersToTranslate.Where(h => !string.IsNullOrWhiteSpace(h.Translated)))
            {
                TranslatedHeaders.Add(translation);
            }
            TranslationSerialization.SerializeTranslations(TranslatedHeaders.ToList());
        }

        public void DeleteTranslation()
        {
            HeadersToTranslate.Add(SelectedTranslation);
            var delete = TranslatedHeaders.Remove(SelectedTranslation);

            if (delete)
            {
                TranslationSerialization.SerializeTranslations(TranslatedHeaders.ToList());
            }
        }
    }
}
