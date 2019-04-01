using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using GalaSoft.MvvmLight.Threading;
using VF_Raporty_Godzin_Pracy.Annotations;
using WinGUI.Servicess;
using WinGUI.Utility;

namespace WinGUI
{
    public class ProgressDialogViewModel : INotifyPropertyChanged 
    {
        #region Privates

        string windowTitle;
        string label;
        string subLabel;
        private int _currentTaskNumber;
        private int _maxTaskNumber;
        private bool _isIndeterminate;
        bool close;

        #endregion

        #region Publics

        public string WindowTitle
        {
            get { return windowTitle; }
            private set
            {
                windowTitle = value;
                RaisePropertyChanged(nameof(WindowTitle));
            }
        }

        public string Label
        {
            get { return label; }
            private set
            {
                label = value;
                RaisePropertyChanged(nameof(Label));
            }
        }

        public string SubLabel
        {
            get { return subLabel; }
            private set
            {
                subLabel = value;
                RaisePropertyChanged(nameof(SubLabel));
            }
        }

        public bool Close
        {
            get { return close; }
            set
            {
                close = value;
                RaisePropertyChanged(nameof(Close));
            }
        }

        public bool IsCancellable { get { return CancelCommand != null; } }
        public CancelCommand CancelCommand { get; private set; }
        public IProgress<ProgressReport> Progress { get; private set; }

        public bool IsIndeterminate
        {
            get { return _isIndeterminate; }
            set
            {
                _isIndeterminate = value;
                RaisePropertyChanged(nameof(IsIndeterminate));
            }
        }

        public int MaxTaskNumber
        {
            get { return _maxTaskNumber; }
            set
            {
                _maxTaskNumber = value; 
                RaisePropertyChanged(nameof(MaxTaskNumber));
            }
        }

        public int CurrentTaskNumber
        {
            get { return _currentTaskNumber; }
            set
            {
                _currentTaskNumber = value; 
                RaisePropertyChanged(nameof(CurrentTaskNumber));
            }
        }

        #endregion

        public ProgressDialogViewModel
        (
            ProgressDialogOptions options,
            CancellationToken cancellationToken,
            CancelCommand cancelCommand
        )
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            WindowTitle = options.WindowTitle;
            Label = options.Label;
            CancelCommand = cancelCommand; // can be null (not cancellable)
            cancellationToken.Register(OnCancelled);
            Progress = new Progress<ProgressReport>(OnProgress);
        }

        void OnCancelled()
        {
            // Cancellation may come from a background thread.
            if (DispatcherHelper.UIDispatcher != null)
                DispatcherHelper.CheckBeginInvokeOnUI(() => Close = true);
            else
                Close = true;
        }

        void OnProgress(ProgressReport obj)
        {
            // Progress will probably come from a background thread.
            if (DispatcherHelper.UIDispatcher != null)
                DispatcherHelper.CheckBeginInvokeOnUI(() => OnProgressReceived(obj));
            else
                OnProgressReceived(obj);

        }

        private void OnProgressReceived(ProgressReport progressReport)
        {
            if (progressReport.IsIndeterminate)
            {
                IsIndeterminate = progressReport.IsIndeterminate;
                SubLabel = progressReport.CurrentTask;
                return;
            }
            SubLabel = progressReport.CurrentTask;
            CurrentTaskNumber = progressReport.CurrentTaskNumber;
            MaxTaskNumber = progressReport.MaxTaskNumber;
            IsIndeterminate = progressReport.IsIndeterminate;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void RaisePropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
