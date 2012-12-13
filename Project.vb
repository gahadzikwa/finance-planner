Public Class Project
    'Constructors and Construction Helpers
    Public Sub New()
        dstProject = New ProjectData()
        frmAllForms = New Collection()
        actLedger = New Ledger()
        actBudget = New Budget()
        jrnJournal = New Journal()
    End Sub

    'Constants
    Protected Const strDefaultTitle As String = "Project 1", _
                    strJournalDefaultTitle As String = "Journal", _
                    strLedgerDefaultTitle As String = "Ledger", _
                    strBudgetDefaultTitle As String = "Budget"

    'ADO.NET Members
    Protected dstProject As FinancePlanner.ProjectData
    Protected drwProject As FinancePlanner.ProjectData.ProjectRow
    Protected drwJournal As FinancePlanner.ProjectData.JournalRow
    Protected drwLedger As FinancePlanner.ProjectData.LedgerRow
    Protected drwBudget As FinancePlanner.ProjectData.BudgetRow
    Protected drwLayout As FinancePlanner.ProjectData.LayoutRow
    'Data Objects
    Protected actLedger As FinancePlanner.Ledger
    Protected actBudget As FinancePlanner.Budget
    Protected jrnJournal As FinancePlanner.Journal
    'Strings
    Protected strFileName As String = String.Empty
    'Associated Forms
    Protected WithEvents frmMainWindow As FinancePlanner.FrmMainWnd
    Protected frmAllForms As Collection


    'Properties
    Public ReadOnly Property FileName()
        Get
            Return strFileName
        End Get
    End Property
    Public Property Title() As String
        Get
            If Not drwProject Is Nothing Then
                Return drwProject.Title
            Else
                Return String.Empty
            End If
        End Get
        Set(ByVal Value As String)
            If Not drwProject Is Nothing Then
                drwProject.Title = Value
            End If
            RaiseEvent Changed()
        End Set
    End Property
    Public ReadOnly Property LedgerAccountCount() As Integer
        Get
            Return dstProject.LedgerAccount.Rows.Count
        End Get
    End Property
    'Pseudo Properties (Reference Returning Functions)
    Public Function GetBudget() As FinancePlanner.AccountTree
        Return actBudget
    End Function
    Public Function GetLedger() As FinancePlanner.AccountTree
        Return actLedger
    End Function
    Public Function GetJournal() As FinancePlanner.Journal
        Return jrnJournal
    End Function


    'Events
    Event Changed()
    Event FileLoaded()
    Event FileSaved()
    Event Closed()
    Event Posted()
    Event Cleared()
    Event Load()


    'Form Assignment
    Public Sub SetMainWindow(ByRef frmMainWindow As FinancePlanner.FrmMainWnd)
        If Not frmMainWindow Is Nothing Then
            Me.frmMainWindow = frmMainWindow
        End If
    End Sub
    'Form Retrieval
    Public Function GetMainWindow() As FrmMainWnd
        Return Me.frmMainWindow
    End Function
    Public Function AllForms() As Collection
        Return frmAllForms
    End Function
    Public Function NewProjectWindow(ByVal WindowType As FinancePlanner.ProjectWindowEnum) As FinancePlanner.FrmProjectWindowBase
        Dim frmNewForm As FrmProjectWindowBase
        Dim frmCurrentForm As FrmProjectWindowBase
        Dim blnAlreadyExists As Boolean

        blnAlreadyExists = False
        For Each frmCurrentForm In frmAllForms
            If frmCurrentForm.Type = WindowType Then
                blnAlreadyExists = True
                frmNewForm = frmCurrentForm
            End If
        Next

        If Not blnAlreadyExists Then
            Select Case WindowType
                Case ProjectWindowEnum.GeneralJournal
                    frmNewForm = New FinancePlanner.FrmGeneralJournal(Me)
                Case ProjectWindowEnum.CashJournal
                    'TODO: Implement Cash Journal
                Case ProjectWindowEnum.PayablesJournal
                    'TODO: Implement Payables Journal
                Case ProjectWindowEnum.ReceivablesJournal
                    'TODO: Implement Receivables Journal

                Case ProjectWindowEnum.GeneralLedger
                    frmNewForm = New FinancePlanner.FrmGeneralLedger(Me)
                Case ProjectWindowEnum.DetailedLedger
                    frmNewForm = New FinancePlanner.FrmDetailedLedger(Me)
                Case ProjectWindowEnum.AccountDetail
                    frmNewForm = New FinancePlanner.FrmAccountDetail(Me)

                Case ProjectWindowEnum.GeneralBudget
                    frmNewForm = New FinancePlanner.FrmGeneralBudget(Me)

                Case ProjectWindowEnum.IncomeStatement
                    'TODO: Implement Income Statement
                Case ProjectWindowEnum.BalanceSheet
                    'TODO: Implement Balance Sheet
                Case ProjectWindowEnum.CashFlowsStatement
                    'TODO: Implement Cash Flows Statement

                Case ProjectWindowEnum.TestWindow
                    frmNewForm = New FinancePlanner.FrmTest(Me)
                Case ProjectWindowEnum.ProjectProperties
                    frmNewForm = New FinancePlanner.FrmProjectProperties(Me)
            End Select
            If Not frmNewForm Is Nothing Then
                AddForm(frmNewForm)
                frmNewForm.Owner = frmMainWindow
                frmNewForm.Show()
            End If
        End If

        If Not frmNewForm Is Nothing Then frmNewForm.Activate()
        Return frmNewForm
    End Function
    Public Sub RemoveProjectWindow(ByVal WindowType As FinancePlanner.ProjectWindowEnum)
        Dim frmCurrentForm As FinancePlanner.FrmProjectWindowBase

        For Each frmCurrentForm In frmAllForms
            If frmCurrentForm.Type = WindowType Then
                RemoveForm(frmCurrentForm)
            End If
        Next
    End Sub
    Private Sub AddForm(ByRef frmForm As FinancePlanner.FrmProjectWindowBase)
        frmAllForms.Add(frmForm, CStr(frmForm.Type))
        RaiseEvent Changed()
    End Sub
    Private Sub RemoveForm(ByRef frmForm As FinancePlanner.FrmProjectWindowBase)
        frmAllForms.Remove(CStr(frmForm.Type))
        RaiseEvent Changed()
    End Sub
    'Layout Handling
    Public Sub LoadAllLayouts()
        Dim drwLayout As FinancePlanner.ProjectData.LayoutRow
        Dim frmForm As FinancePlanner.FrmProjectWindowBase

        For Each drwLayout In drwProject.GetLayoutRows()
            frmForm = NewProjectWindow(drwLayout.WindowType)
            LoadLayout(frmForm)
        Next
    End Sub
    Public Sub SaveAllLayouts()
        Dim frmForm As FinancePlanner.FrmProjectWindowBase

        dstProject.Layout.Clear()

        For Each frmForm In frmAllForms
            SaveLayout(frmForm)
        Next
    End Sub
    Private Sub SaveLayout(ByRef frmForm As FinancePlanner.FrmProjectWindowBase)
        Dim drwLayout As FinancePlanner.ProjectData.LayoutRow
        Dim blnExists As Boolean

        If Not frmForm Is Nothing Then
            blnExists = False
            For Each drwLayout In drwProject.GetLayoutRows()
                If drwLayout.WindowType = frmForm.Type Then
                    drwLayout.Height = frmForm.Height
                    drwLayout.Width = frmForm.Width
                    drwLayout.Top = frmForm.Top
                    drwLayout.Left = frmForm.Left
                    blnExists = True
                End If
            Next

            If Not blnExists Then
                dstProject.Layout.AddLayoutRow(frmForm.Type, frmForm.Height, frmForm.Width, _
                                               frmForm.Top, frmForm.Left, drwProject)
            End If
        End If
    End Sub
    Private Sub LoadLayout(ByRef frmForm As FinancePlanner.FrmProjectWindowBase)
        Dim drwLayout As FinancePlanner.ProjectData.LayoutRow
        If Not frmForm Is Nothing Then
            For Each drwLayout In drwProject.GetLayoutRows()
                If drwLayout.WindowType = frmForm.Type Then
                    frmForm.Height = drwLayout.Height
                    frmForm.Width = drwLayout.Width
                    frmForm.Top = drwLayout.Top
                    frmForm.Left = drwLayout.Left
                End If
            Next
        End If
    End Sub
    'Loading
    Public Function LoadFile(ByVal FileName As String) As Boolean
        Dim blnError As Boolean = False
        Try
            If FileName <> "" Then
                Me.strFileName = FileName
                dstProject.Clear()
                dstProject.ReadXml(FileName)
            End If
        Catch
            blnError = True
            MessageBox.Show("The file is corrupt.", "Error Loading File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        If Not blnError Then
            Try
                drwProject = dstProject.Project(0)
                drwJournal = drwProject.GetJournalRows()(0)
                drwLedger = drwProject.GetLedgerRows()(0)
                drwBudget = drwProject.GetBudgetRows()(0)
            Catch
                blnError = True
                MessageBox.Show("The file format is incorrect.", "Error Loading File", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If

        If Not blnError Then
            LoadObjectsFromDataset()
            LoadAllLayouts()
            RaiseEvent FileLoaded()
            RaiseEvent Load()
        End If

        Return Not blnError
    End Function
    Public Function ReloadFile() As Boolean
        Return LoadFile(Me.strFileName)
    End Function
    Private Sub LoadObjectsFromDataset()
        LoadJournalFromDataset()
        LoadLedgerFromDataset()
    End Sub
    Private Sub LoadLedgerFromDataset()
        Dim drwLedgerAccount As FinancePlanner.ProjectData.LedgerAccountRow
        Dim accLedgerAccount As FinancePlanner.Account

        actLedger.Clear()
        For Each drwLedgerAccount In drwLedger.GetLedgerAccountRows()
            accLedgerAccount = New FinancePlanner.Account()
            accLedgerAccount.Name = AccountPath.AccountName(drwLedgerAccount.Path)
            accLedgerAccount.Unit = drwLedgerAccount.Unit
            accLedgerAccount.NormalBalanceAsString = drwLedgerAccount.NormalBalance

            actLedger.AddAccountWithFullPath(accLedgerAccount, drwLedgerAccount.Path)
        Next
    End Sub
    Private Sub LoadJournalFromDataset()
        Dim jeEntry As FinancePlanner.JournalEntry
        Dim jtTransaction As FinancePlanner.JournalTransaction
        Dim drwEntry As FinancePlanner.ProjectData.EntryRow
        Dim drwTransaction As FinancePlanner.ProjectData.TransactionRow

        jrnJournal.Clear()

        For Each drwEntry In drwJournal.GetEntryRows
            jeEntry = New FinancePlanner.JournalEntry()
            jeEntry.ID = drwEntry.ID
            jeEntry._Date = drwEntry._Date
            jeEntry.Comment = drwEntry.Comment
            If Not drwEntry.IsTypeNull Then
                jeEntry.TypeAsString = drwEntry.Type
            Else
                jeEntry.Type = EntryEnum.Transaction
            End If

            For Each drwTransaction In drwEntry.GetTransactionRows

                jtTransaction = New FinancePlanner.JournalTransaction()
                jtTransaction.Path = drwTransaction.Path
                jtTransaction.NormalBalanceAsString = drwTransaction.NormalBalance
                If Not drwTransaction.IsValueNull() Then
                    jtTransaction.Balance = drwTransaction.Value
                End If

                If Not drwTransaction.IsQuantityNull Then
                    jtTransaction.Quantity = drwTransaction.Quantity
                End If

                jeEntry.AddTransaction(jtTransaction)
            Next

            jrnJournal.AddEntry(jeEntry, Journal.IDGenerationEnum.DoNotGenerateID)
        Next

    End Sub
    'Saving
    Public Function SaveFileAs(ByVal FileName As String) As Boolean
        Dim blnError As Boolean = False

        SaveObjectsToDataset()

        Try
            SaveAllLayouts()
            dstProject.WriteXml(FileName)
            Me.strFileName = FileName
        Catch
            blnError = True
        End Try

        If Not blnError Then
            RaiseEvent FileSaved()
        End If

        Return Not blnError
    End Function
    Public Function SaveFile() As Boolean
        Return SaveFileAs(Me.strFileName)
    End Function
    Private Sub SaveObjectsToDataset()
        SaveJournalToDataset()
        SaveLedgerToDataset()
    End Sub
    Private Sub SaveLedgerToDataset()
        Dim drwLedgerAccount As FinancePlanner.ProjectData.LedgerAccountRow
        Dim accRootAccount As FinancePlanner.Account

        'Clear accounts in dataset
        dstProject.LedgerAccount.Clear()

        'Parse account tree and save only leaf nodes
        For Each accRootAccount In actLedger.Accounts
            If accRootAccount.GetSubAccounts.Count > 0 Then
                SaveSubAccountsToDataset(accRootAccount)
            Else
                dstProject.LedgerAccount.AddLedgerAccountRow(accRootAccount.Unit, accRootAccount.NormalBalanceAsString, accRootAccount.FullPath, drwLedger)
            End If
        Next
    End Sub
    Private Sub SaveSubAccountsToDataset(ByRef accAccount As FinancePlanner.Account)
        Dim accSubAccount

        For Each accSubAccount In accAccount.GetSubAccounts
            If accSubAccount.SubAccounts.Count > 0 Then
                SaveSubAccountsToDataset(accSubAccount)
            Else
                dstProject.LedgerAccount.AddLedgerAccountRow(accSubAccount.Unit, accSubAccount.GetNormalBalanceAsString(), accSubAccount.FullPath, drwLedger)
            End If
        Next
    End Sub
    Private Sub SaveJournalToDataset()
        Dim jeEntry As FinancePlanner.JournalEntry
        Dim jtTransaction As FinancePlanner.JournalTransaction
        Dim drwThisEntry As FinancePlanner.ProjectData.EntryRow
        Dim drwThisTransaction As FinancePlanner.ProjectData.TransactionRow
        Dim blnErrorInEntry As Boolean

        For Each jeEntry In jrnJournal.EntriesAsArray(Journal.CompareByEnum.CompareByID)
            If Not jeEntry.GetTransactions Is Nothing Then
                Try
                    drwThisEntry = dstProject.Entry.AddEntryRow(jeEntry.Comment, _
                                   jeEntry.TypeAsString, jeEntry._Date, jeEntry.ID, drwJournal)
                Catch
                    blnErrorInEntry = True
                End Try

                If Not blnErrorInEntry Then
                    For Each jtTransaction In jeEntry.GetTransactions
                        Try
                            drwThisTransaction = dstProject.Transaction.AddTransactionRow( _
                             jtTransaction.Path, _
                             jtTransaction.Balance, jtTransaction.Quantity, _
                             jtTransaction.NormalBalanceAsString, _
                             drwThisEntry)
                        Catch
                            blnErrorInEntry = True
                        End Try
                    Next
                End If
            End If
        Next

    End Sub
    'Close the Project
    Public Sub Close()
        Dim frmProjectWindow As FrmProjectWindowBase

        For Each frmProjectWindow In frmAllForms
            frmProjectWindow.Dispose()
        Next

        RaiseEvent Closed()
    End Sub
    'Misc
    Public Sub Clear()
        dstProject.Clear()
        drwProject = dstProject.Project.AddProjectRow(strDefaultTitle)
        drwJournal = dstProject.Journal.AddJournalRow(drwProject)
        drwLedger = dstProject.Ledger.AddLedgerRow(StartOfThisYear(), Date.Today, drwProject)
        drwLedger.SetLastDateNull()
        drwBudget = dstProject.Budget.AddBudgetRow(drwProject)
        'Me.drwLayout = Me.dstProject.Layout.AddLayoutRow(pnlJournalPane.Width, _
        '                                                 pnlLedgerPane.Width, _
        '                                                 0, _
        '                                                 Me.drwProject)
        RaiseEvent Cleared()
        RaiseEvent Load()
    End Sub
    Public Sub RaiseChanged()
        RaiseEvent Changed()
    End Sub
    Public Sub Post()
        Dim frmForm As FrmProjectWindowBase

        For Each frmForm In frmAllForms
            If TypeOf (frmForm) Is FinancePlanner.IParses Then
                CType(frmForm, FinancePlanner.IParses).Parse()
            End If
        Next

        SumAccounts()

        RaiseEvent Changed()

        For Each frmForm In frmAllForms
            frmForm.Render()
        Next

        RaiseEvent Posted()
    End Sub
    Private Sub SumAccounts()
        Dim jeEntry As FinancePlanner.JournalEntry
        Dim jtTransaction As FinancePlanner.JournalTransaction
        Dim accMatchList() As FinancePlanner.Account
        Dim accAccount As FinancePlanner.Account
        Dim blnSingleMatchFound As Boolean

        actLedger.ResetAccountBalances()

        For Each jeEntry In jrnJournal.Entries
            If jeEntry.IsBalanced Then
                For Each jtTransaction In jeEntry.GetTransactions

                    blnSingleMatchFound = False

                    Do While Not blnSingleMatchFound
                        accMatchList = actLedger.FindAccountsWithAlias(jtTransaction.Path)

                        Select Case accMatchList.Length
                            Case 1
                                accAccount = accMatchList(0)
                                blnSingleMatchFound = True
                            Case 0
                                blnSingleMatchFound = False
                            Case Else
                                blnSingleMatchFound = False
                        End Select
                    Loop

                    accAccount.PostBalance(jtTransaction.Balance, jtTransaction.NormalBalance)
                    accAccount.PostQuantity(jtTransaction.Quantity, jtTransaction.NormalBalance)

                Next
            End If
        Next
    End Sub

End Class
