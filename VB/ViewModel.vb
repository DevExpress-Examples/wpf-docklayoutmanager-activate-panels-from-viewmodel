Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows.Input
Imports System.Collections.Specialized
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Namespace ActivateViaViewModel
    Public MustInherit Class NotifyPropertyChanged
        Implements INotifyPropertyChanged

        Protected Overridable Sub OnPropertyChanged(ByVal [property] As String)
            Dim handler As PropertyChangedEventHandler = PropertyChangedEvent
            If handler IsNot Nothing Then
                handler(Me, New PropertyChangedEventArgs([property]))
            End If
        End Sub
        #Region "INotifyPropertyChanged Members"
        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
        #End Region
    End Class
    Public Class MainWindowViewModel
        Inherits NotifyPropertyChanged

        Private count As Integer
        Public Sub New()
            Dim document As DocumentViewModel = CreateNewDocument()
            Documents.Add(document)
        End Sub
        Private _Documents As ObservableCollection(Of DocumentViewModel)
        Public ReadOnly Property Documents() As ObservableCollection(Of DocumentViewModel)
            Get
                If _Documents Is Nothing Then
                    _Documents = New ObservableCollection(Of DocumentViewModel)()
                    AddHandler _Documents.CollectionChanged, AddressOf Me.OnItemsChanged
                End If
                Return _Documents
            End Get
        End Property
        Private _AddNewTabCommand As ICommand
        Public ReadOnly Property AddNewTabCommand() As ICommand
            Get
                If _AddNewTabCommand Is Nothing Then
                    _AddNewTabCommand = New RelayCommand(New Action(Of Object)(AddressOf OnDocumentRequestAddNewTab))
                End If
                Return _AddNewTabCommand
            End Get
        End Property
        Private Sub OnItemsChanged(ByVal sender As Object, ByVal e As NotifyCollectionChangedEventArgs)
            If e.NewItems IsNot Nothing AndAlso e.NewItems.Count <> 0 Then
                For Each document As DocumentViewModel In e.NewItems
                    AddHandler document.RequestClose, AddressOf Me.OnDocumentRequestClose
                Next document
            End If
            If e.OldItems IsNot Nothing AndAlso e.OldItems.Count <> 0 Then
                For Each document As DocumentViewModel In e.OldItems
                    RemoveHandler document.RequestClose, AddressOf OnDocumentRequestClose
                Next document
            End If
        End Sub
        Private Sub OnDocumentRequestClose(ByVal sender As Object, ByVal e As EventArgs)
            Dim document As DocumentViewModel = TryCast(sender, DocumentViewModel)
            If Documents.Count = 1 Then
                AddNewTab()
            End If
            If document IsNot Nothing Then
                Documents.Remove(document)
            End If
        End Sub
        Private Sub OnDocumentRequestAddNewTab(ByVal param As Object)
            Dim document = AddNewTab()
            document.IsActive = True
        End Sub
        Private Function AddNewTab() As DocumentViewModel
            Dim document As DocumentViewModel = CreateNewDocument()
            Documents.Add(document)
            Return document
        End Function
        Private Function CreateNewDocument() As DocumentViewModel
            Dim document As New DocumentViewModel() With {.DisplayName = "Document" & count, .Content = "Content" & count}
            count += 1
            Return document
        End Function
    End Class
    Public Class DocumentViewModel
        Inherits NotifyPropertyChanged

        Public Sub New()
        End Sub
        Private _IsActive As Boolean

        Public Property IsActive() As Boolean
            Get
                Return _IsActive
            End Get
            Set(ByVal value As Boolean)
                If _IsActive = value Then
                    Return
                End If
                _IsActive = value
                OnPropertyChanged("IsActive")
            End Set
        End Property
        Private _DisplayName As String
        Public Property DisplayName() As String
            Get
                Return _DisplayName
            End Get
            Set(ByVal value As String)
                If _DisplayName = value Then
                    Return
                End If
                _DisplayName = value
                OnPropertyChanged("DisplayName")
            End Set
        End Property
        Private _Content As Object
        Public Property Content() As Object
            Get
                Return _Content
            End Get
            Set(ByVal value As Object)
                If _Content Is value Then
                    Return
                End If
                _Content = value
                OnPropertyChanged("Content")
            End Set
        End Property
        Private _CloseCommand As ICommand
        Public ReadOnly Property CloseCommand() As ICommand
            Get
                If _CloseCommand Is Nothing Then
                    _CloseCommand = New RelayCommand(New Action(Of Object)(AddressOf OnRequestClose))
                End If
                Return _CloseCommand
            End Get
        End Property
        Public Event RequestClose As EventHandler
        Private Sub OnRequestClose(ByVal param As Object)
            Dim handler As EventHandler = Me.RequestCloseEvent
            If handler IsNot Nothing Then
                handler(Me, EventArgs.Empty)
            End If
        End Sub
    End Class
    Public Class RelayCommand
        Implements ICommand

        Private ReadOnly _execute As Action(Of Object)
        Private ReadOnly _canExecute As Predicate(Of Object)
        Public Sub New(ByVal execute As Action(Of Object))
            Me.New(execute, Nothing)
        End Sub
        Public Sub New(ByVal execute As Action(Of Object), ByVal canExecute As Predicate(Of Object))
            If execute Is Nothing Then
                Throw New ArgumentNullException("execute")
            End If
            _execute = execute
            _canExecute = canExecute
        End Sub
        #Region "ICommand Members"
        Public Function CanExecute(ByVal parameter As Object) As Boolean Implements ICommand.CanExecute
            Return If(_canExecute Is Nothing, True, _canExecute(parameter))
        End Function
#If SILVERLIGHT Then
        Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged
        Protected Overridable Sub OnCanExecuteChanged(ByVal e As EventArgs)
            Dim canExecuteChanged = CanExecuteChangedEvent
            If canExecuteChanged IsNot Nothing Then
                canExecuteChanged(Me, e)
            End If
        End Sub
        Public Sub RaiseCanExecuteChanged()
            OnCanExecuteChanged(EventArgs.Empty)
        End Sub
#Else
        Public Custom Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged
            AddHandler(ByVal value As EventHandler)
                AddHandler CommandManager.RequerySuggested, value
            End AddHandler
            RemoveHandler(ByVal value As EventHandler)
                RemoveHandler CommandManager.RequerySuggested, value
            End RemoveHandler
            RaiseEvent(ByVal sender As System.Object, ByVal e As System.EventArgs)
            End RaiseEvent
        End Event
#End If
        Public Sub Execute(ByVal parameter As Object) Implements ICommand.Execute
            _execute(parameter)
        End Sub
        #End Region ' ICommand Members
    End Class
End Namespace
