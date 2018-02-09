using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Messaging;
using VF_Raporty_Godzin_Pracy;
using VF_Raporty_Godzin_Pracy.Annotations;
using WinGUI.Utility;

namespace WinGUI.ViewModel
{
    public class TlumaczeniaViewModel : INotifyPropertyChanged
    {
        private List<Naglowek> _naglowkiDoTlumaczenia;

        public List<Naglowek> NaglowkiDoTlumaczenia
        {
            get => _naglowkiDoTlumaczenia;
            set
            {
                _naglowkiDoTlumaczenia = value; 
                OnPropertyChanged(nameof(NaglowkiDoTlumaczenia));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public TlumaczeniaViewModel()
        {
            Messenger.Default.Register<WyslijDoTlumaczenia>
            (
                this,
                OtrzymanoNaglowki
            );
        }

        private void OtrzymanoNaglowki(WyslijDoTlumaczenia action)
        {
            NaglowkiDoTlumaczenia = action.NaglowkiDoTlumaczenia;
        }
    }
}
