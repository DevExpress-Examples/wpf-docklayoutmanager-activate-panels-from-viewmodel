using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;
using System.Collections.Specialized;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace ActivateViaViewModel {
    public abstract class NotifyPropertyChanged : INotifyPropertyChanged {
        protected virtual void OnPropertyChanged(string property) {
            PropertyChangedEventHandler handler = PropertyChanged;
            if(handler != null) {
                handler(this, new PropertyChangedEventArgs(property));
            }
        }
        #region INotifyPropertyChanged Members
        public event PropertyChangedEventHandler PropertyChanged;
        #endregion
    }
    public class MainWindowViewModel : NotifyPropertyChanged {
        int count;
        public MainWindowViewModel() {
            DocumentViewModel document = CreateNewDocument();
            Documents.Add(document);
        }
        ObservableCollection<DocumentViewModel> _Documents;
        public ObservableCollection<DocumentViewModel> Documents {
            get {
                if(_Documents == null) {
                    _Documents = new ObservableCollection<DocumentViewModel>();
                    _Documents.CollectionChanged += this.OnItemsChanged;
                }
                return _Documents;
            }
        }
        ICommand _AddNewTabCommand;
        public ICommand AddNewTabCommand {
            get {
                if(_AddNewTabCommand == null)
                    _AddNewTabCommand = new RelayCommand(new Action<object>(OnDocumentRequestAddNewTab));
                return _AddNewTabCommand;
            }
        }
        void OnItemsChanged(object sender, NotifyCollectionChangedEventArgs e) {
            if(e.NewItems != null && e.NewItems.Count != 0)
                foreach(DocumentViewModel document in e.NewItems) {
                    document.RequestClose += this.OnDocumentRequestClose;
                }
            if(e.OldItems != null && e.OldItems.Count != 0)
                foreach(DocumentViewModel document in e.OldItems)
                    document.RequestClose -= this.OnDocumentRequestClose;
        }
        void OnDocumentRequestClose(object sender, EventArgs e) {
            DocumentViewModel document = sender as DocumentViewModel;
            if(Documents.Count == 1) AddNewTab();
            if(document != null) {
                Documents.Remove(document);
            }
        }
        void OnDocumentRequestAddNewTab(object param) {
            var document = AddNewTab();
            document.IsActive = true;
        }
        DocumentViewModel AddNewTab() {
            DocumentViewModel document = CreateNewDocument();
            Documents.Add(document);
            return document;
        }
        DocumentViewModel CreateNewDocument() {
            DocumentViewModel document = new DocumentViewModel() { DisplayName = "Document" + count, Content = "Content" + count };
            count++;
            return document;
        }
    }
    public class DocumentViewModel : NotifyPropertyChanged {
        public DocumentViewModel() {
        }
        private bool _IsActive;

        public bool IsActive {
            get { return _IsActive; }
            set {
                if(_IsActive == value) return;
                _IsActive = value;
                OnPropertyChanged("IsActive");
            }
        }
        private string _DisplayName;
        public string DisplayName {
            get { return _DisplayName; }
            set {
                if(_DisplayName == value) return;
                _DisplayName = value;
                OnPropertyChanged("DisplayName");
            }
        }
        private object _Content;
        public object Content {
            get { return _Content; }
            set {
                if(_Content == value) return;
                _Content = value;
                OnPropertyChanged("Content");
            }
        }
        ICommand _CloseCommand;
        public ICommand CloseCommand {
            get {
                if(_CloseCommand == null)
                    _CloseCommand = new RelayCommand(new Action<object>(OnRequestClose));
                return _CloseCommand;
            }
        }
        public event EventHandler RequestClose;
        void OnRequestClose(object param) {
            EventHandler handler = this.RequestClose;
            if(handler != null)
                handler(this, EventArgs.Empty);
        }
    }
    public class RelayCommand : ICommand {
        readonly Action<object> _execute;
        readonly Predicate<object> _canExecute;
        public RelayCommand(Action<object> execute)
            : this(execute, null) {
        }
        public RelayCommand(Action<object> execute, Predicate<object> canExecute) {
            if(execute == null)
                throw new ArgumentNullException("execute");
            _execute = execute;
            _canExecute = canExecute;
        }
        #region ICommand Members
        public bool CanExecute(object parameter) {
            return _canExecute == null ? true : _canExecute(parameter);
        }
#if SILVERLIGHT
        public event EventHandler CanExecuteChanged;
        protected virtual void OnCanExecuteChanged(EventArgs e) {
            var canExecuteChanged = CanExecuteChanged;
            if(canExecuteChanged != null)
                canExecuteChanged(this, e);
        }
        public void RaiseCanExecuteChanged() {
            OnCanExecuteChanged(EventArgs.Empty);
        }
#else
        public event EventHandler CanExecuteChanged {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
#endif
        public void Execute(object parameter) {
            _execute(parameter);
        }
        #endregion // ICommand Members
    }
}
