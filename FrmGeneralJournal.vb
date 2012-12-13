Public Class FrmGeneralJournal
    Inherits FinancePlanner.FrmProjectWindowBase
    Implements FinancePlanner.IParses

#Region " Windows Form Designer generated code "

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents pnlJournalPane As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents shtJournal As AxOWC11.AxSpreadsheet
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents lblAccounts As System.Windows.Forms.Label
    Friend WithEvents lblDebits As System.Windows.Forms.Label
    Friend WithEvents lblCredits As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmGeneralJournal))
        Me.pnlJournalPane = New System.Windows.Forms.Panel()
        Me.shtJournal = New AxOWC11.AxSpreadsheet()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblCredits = New System.Windows.Forms.Label()
        Me.lblDebits = New System.Windows.Forms.Label()
        Me.lblAccounts = New System.Windows.Forms.Label()
        Me.lblType = New System.Windows.Forms.Label()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.pnlJournalPane.SuspendLayout()
        CType(Me.shtJournal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlJournalPane
        '
        Me.pnlJournalPane.Controls.AddRange(New System.Windows.Forms.Control() {Me.shtJournal, Me.Panel1})
        Me.pnlJournalPane.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlJournalPane.Name = "pnlJournalPane"
        Me.pnlJournalPane.Size = New System.Drawing.Size(592, 370)
        Me.pnlJournalPane.TabIndex = 0
        '
        'shtJournal
        '
        Me.shtJournal.ContainingControl = Me
        Me.shtJournal.DataSource = Nothing
        Me.shtJournal.Dock = System.Windows.Forms.DockStyle.Fill
        Me.shtJournal.Enabled = True
        Me.shtJournal.Location = New System.Drawing.Point(0, 32)
        Me.shtJournal.Name = "shtJournal"
        Me.shtJournal.OcxState = CType(resources.GetObject("shtJournal.OcxState"), System.Windows.Forms.AxHost.State)
        Me.shtJournal.Size = New System.Drawing.Size(592, 338)
        Me.shtJournal.TabIndex = 1
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblCredits, Me.lblDebits, Me.lblAccounts, Me.lblType, Me.lblDate})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(592, 32)
        Me.Panel1.TabIndex = 0
        '
        'lblCredits
        '
        Me.lblCredits.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblCredits.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblCredits.Location = New System.Drawing.Point(320, 0)
        Me.lblCredits.Name = "lblCredits"
        Me.lblCredits.Size = New System.Drawing.Size(64, 32)
        Me.lblCredits.TabIndex = 4
        Me.lblCredits.Text = "Credits"
        Me.lblCredits.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblDebits
        '
        Me.lblDebits.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblDebits.ForeColor = System.Drawing.Color.DarkRed
        Me.lblDebits.Location = New System.Drawing.Point(256, 0)
        Me.lblDebits.Name = "lblDebits"
        Me.lblDebits.Size = New System.Drawing.Size(64, 32)
        Me.lblDebits.TabIndex = 3
        Me.lblDebits.Text = "Debits"
        Me.lblDebits.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblAccounts
        '
        Me.lblAccounts.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblAccounts.Location = New System.Drawing.Point(128, 0)
        Me.lblAccounts.Name = "lblAccounts"
        Me.lblAccounts.Size = New System.Drawing.Size(128, 32)
        Me.lblAccounts.TabIndex = 2
        Me.lblAccounts.Text = "Accounts and Comments"
        Me.lblAccounts.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblType
        '
        Me.lblType.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblType.Location = New System.Drawing.Point(64, 0)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(64, 32)
        Me.lblType.TabIndex = 1
        Me.lblType.Text = "Type"
        Me.lblType.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblDate
        '
        Me.lblDate.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(64, 32)
        Me.lblDate.TabIndex = 0
        Me.lblDate.Text = "Date"
        Me.lblDate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'FrmGeneralJournal
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 370)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.pnlJournalPane})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "FrmGeneralJournal"
        Me.Text = "Journal"
        Me.pnlJournalPane.ResumeLayout(False)
        CType(Me.shtJournal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private cbeSortSetting As FinancePlanner.Journal.CompareByEnum = Journal.CompareByEnum.CompareByDate

    Shadows Class Constants
        Public Const StartingRow = 1, _
                     StartingColumn = 1, _
                     DebitsColumn = StartingColumn + 4, _
                     CreditsColumn = DebitsColumn + 1, _
                     DateColumn = StartingColumn, _
                     TypeColumn = StartingColumn + 1, _
                     AccountsColumn = StartingColumn + 2, _
                     EndingColumn = CreditsColumn, _
                     EndingBlanks = 10, _
                     NumberFormat As String = "$0.00;$_(0.00_);""----""", _
                     TotalColumns = 6
        Public Const DateColumnWidth = 0.125, _
                     TypeColumnWidth = 0.125, _
                     AccountsDebitWidth = 0.1, _
                     AccountsCreditWidth = 0.4, _
                     DebitsColumnWidth = 0.125, _
                     CreditsColumnWidth = 0.125
        Class Colors
            Public Const DataForeground = &H0, _
                         DataBackground = &HCCCCCC, _
                         DebitForeground = &H66, _
                         CreditForeground = &H660000, _
                         CommentBackground = &HFFFFFF, _
                         CommentForeground = &H0, _
                         ErrorBackground = &HAA
        End Class
    End Class

    Public Sub New(ByRef prjProject As FinancePlanner.Project)
        MyBase.New(prjProject)
        InitializeComponent()
    End Sub

    'Window Type
    Public Overrides ReadOnly Property Type() As FinancePlanner.ProjectWindowEnum
        Get
            Return FinancePlanner.ProjectWindowEnum.GeneralJournal
        End Get
    End Property

    Public Overrides ReadOnly Property Title() As String
        Get
            Return "General Journal"
        End Get
    End Property

    'Events
    Event Changed() Implements IParses.Changed

    'Parsing Subs
    Public Sub Parse() Implements IParses.Parse
        'TODO: Debug General Journal Parse
        Dim blnEndOfJournal As Boolean, intBlankLines As Integer, _
            intRow As Integer, intEntryRows As Integer, _
            strEntryAccount As String, _
            dblEntryValue As Double, _
            strDate As String, intCounter As Integer, _
            strDebitValue As String, strCreditValue As String, _
            strType As String, _
            dblNetBalance As Double, _
            strComment As String = String.Empty, _
            blnErrorInEntry As Boolean, _
            blnErrorInTransaction As Boolean, _
            jeEntry As FinancePlanner.JournalEntry, _
            jtTransaction As FinancePlanner.JournalTransaction, _
            beTransactionBalance As FinancePlanner.BalanceEnum, _
            jrnJournal As FinancePlanner.Journal

        'Search entire journal for entries.

        jrnJournal = prjProject.GetJournal()
        jrnJournal.Clear()
        intRow = Constants.StartingRow

        Do

            strDate = shtJournal.ActiveSheet.Cells(intRow, Constants.DateColumn).Value
            strType = shtJournal.ActiveSheet.Cells(intRow, Constants.TypeColumn).Value
            If strType Is Nothing Then strType = String.Empty

            If strDate <> String.Empty Then
                intBlankLines = 0
                intEntryRows = 0
                blnErrorInEntry = False
                dblNetBalance = 0

                Try
                    jeEntry = New JournalEntry()
                    jeEntry._Date = CDate(strDate)
                    jeEntry.TypeAsString = strType
                Catch
                    blnErrorInEntry = True
                End Try

                strDebitValue = shtJournal.ActiveSheet.Cells(intRow, Constants.DebitsColumn).Value
                strCreditValue = shtJournal.ActiveSheet.Cells(intRow, Constants.CreditsColumn).Value

                Do Until strDebitValue Is Nothing And strCreditValue Is Nothing
                    blnErrorInTransaction = False
                    intEntryRows = intEntryRows + 1

                    If Not strDebitValue Is Nothing Then
                        strEntryAccount = _
                            shtJournal.ActiveSheet.Cells(intRow + intEntryRows - 1, _
                                                         Constants.AccountsColumn).Value
                        If strEntryAccount Is Nothing Then
                            strEntryAccount = _
                                shtJournal.ActiveSheet.Cells(intRow + intEntryRows - 1, _
                                                             Constants.AccountsColumn + 1).Value
                        End If
                        dblEntryValue = _
                            Math.Round(Val(strDebitValue), 2)

                        beTransactionBalance = BalanceEnum.Debit
                    End If
                    If Not strCreditValue Is Nothing Then
                        strEntryAccount = _
                            shtJournal.ActiveSheet.Cells(intRow + intEntryRows - 1, _
                                                         Constants.AccountsColumn + 1).Value
                        If strEntryAccount Is Nothing Then
                            strEntryAccount = _
                                shtJournal.ActiveSheet.Cells(intRow + intEntryRows - 1, _
                                                             Constants.AccountsColumn).Value
                        End If
                        dblEntryValue = _
                            Math.Round(Val(strCreditValue), 2)

                        beTransactionBalance = BalanceEnum.Credit
                    End If

                    Try
                        jtTransaction = New FinancePlanner.JournalTransaction()
                        jtTransaction.Path = strEntryAccount
                        jtTransaction.NormalBalance = beTransactionBalance
                        jtTransaction.Balance = dblEntryValue

                        jeEntry.AddTransaction(jtTransaction)
                    Catch
                        blnErrorInTransaction = True
                    End Try

                    strDebitValue = shtJournal.ActiveSheet.Cells(intRow + intEntryRows, Constants.DebitsColumn).Value
                    strCreditValue = shtJournal.ActiveSheet.Cells(intRow + intEntryRows, Constants.CreditsColumn).Value

                Loop

                blnErrorInEntry = Not jeEntry.IsBalanced()

                intRow = intRow + intEntryRows

                With shtJournal.ActiveSheet
                    If Not .Cells(intRow, Constants.AccountsColumn).Value Is Nothing Then
                        jeEntry.Comment = _
                            .Cells(intRow, Constants.AccountsColumn).Value
                    End If
                End With

                'Add the entry to the journal once it's been built
                jrnJournal.AddEntry(jeEntry, Journal.IDGenerationEnum.GenerateID)
            Else

                intBlankLines = intBlankLines + 1
                If intBlankLines >= Constants.EndingBlanks _
                    Then blnEndOfJournal = True

            End If

            intRow = intRow + 1

        Loop While Not blnEndOfJournal

    End Sub

    'Rendering Subs
    Public Overrides Sub Render()
        'MsgBox("Rendering General Journal")
        'TODO: Revise General Journal Render
        Dim jeEntry As FinancePlanner.JournalEntry
        Dim jtTransaction As FinancePlanner.JournalTransaction
        Dim intRow As Integer, intSubRow As Integer, intColumn As Integer

        shtJournal.ScreenUpdating = False

        With shtJournal.ActiveSheet
            .Cells.Clear()
        End With

        intRow = 0
        For Each jeEntry In prjProject.GetJournal.EntriesAsArray(cbeSortSetting)
            intRow += 1
            intSubRow = -1
            With shtJournal.ActiveSheet
                .Cells(intRow, Constants.DateColumn).Value = _
                    jeEntry._Date.ToShortDateString()
                .Cells(intRow, Constants.TypeColumn).Value = _
                    jeEntry.TypeAsString

                With .Range(.Cells(intRow, Constants.StartingColumn), _
                            .Cells(intRow, Constants.EndingColumn))
                    .Borders(shtJournal.Constants.xlEdgeTop).LineStyle = shtJournal.Constants.xlContinuous
                    .Borders(shtJournal.Constants.xlEdgeTop).Weight = shtJournal.Constants.xlThin
                End With
            End With

            For Each jtTransaction In jeEntry.GetTransactions
                intSubRow += 1
                With shtJournal.ActiveSheet
                    With .Range(.Cells(intRow + intSubRow, Constants.StartingColumn), _
                                .Cells(intRow + intSubRow, Constants.EndingColumn))
                        .Interior.Color = Constants.Colors.DataBackground
                        .Font.Color = Constants.Colors.DataForeground
                    End With

                    Select Case jtTransaction.NormalBalance
                        Case BalanceEnum.Debit
                            With .Cells(intRow + intSubRow, Constants.DebitsColumn)
                                .Value = jtTransaction.Balance
                                .Font.Color = Constants.Colors.DebitForeground
                                .NumberFormat = Constants.NumberFormat
                            End With
                            .Cells(intRow + intSubRow, Constants.AccountsColumn).Value = _
                                jtTransaction.Path
                        Case BalanceEnum.Credit
                            With .Cells(intRow + intSubRow, Constants.CreditsColumn)
                                .Value = jtTransaction.Balance
                                .Font.Color = Constants.Colors.CreditForeground
                                .NumberFormat = Constants.NumberFormat
                            End With
                            .Cells(intRow + intSubRow, Constants.AccountsColumn + 1).Value = _
                                jtTransaction.Path
                    End Select

                End With
            Next

            intRow += intSubRow + 1

            With shtJournal.ActiveSheet.Cells(intRow, Constants.AccountsColumn)
                .Value = jeEntry.Comment
            End With
            With shtJournal.ActiveSheet
                With .Range(.Cells(intRow, Constants.StartingColumn), _
                            .Cells(intRow, Constants.EndingColumn))
                    .Interior.Color = Constants.Colors.CommentBackground
                    .Font.Color = Constants.Colors.CommentForeground
                End With
            End With

            If Not jeEntry.IsBalanced() Then
                With shtJournal.ActiveSheet
                    With .Range(.Cells(intRow - (intSubRow + 1), Constants.StartingColumn), _
                                .Cells(intRow, Constants.EndingColumn))
                        .Interior.Color = Constants.Colors.ErrorBackground
                    End With
                End With
            End If

        Next

        If intRow > 0 Then
            For intColumn = Constants.StartingColumn To Constants.EndingColumn - 1
                With shtJournal.ActiveSheet
                    With .Range(.Cells(1, intColumn), _
                                .Cells(intRow, intColumn))
                        If intColumn <> Constants.AccountsColumn Then
                            .Borders(shtJournal.Constants.xlEdgeRight).LineStyle = shtJournal.Constants.xlContinuous
                            .Borders(shtJournal.Constants.xlEdgeRight).Weight = shtJournal.Constants.xlThin
                        End If
                    End With
                End With
            Next
        End If

        shtJournal.ScreenUpdating = True

    End Sub

    'Resize Journal Columns when the form is loaded or resized
    Protected Sub ResizeJournalSheet()
        Dim intCounter As Integer
        Dim blnCanResize As Boolean = True
        Dim dblJournalWidth As Double = CType(shtJournal.Width - 15, Double)
        Dim dblPixelRatio As Double = 7.4

        Try
            'test date column resizing
            shtJournal.ActiveSheet.Cells(1, Constants.DateColumn).ColumnWidth = _
                            dblJournalWidth * _
                            Constants.DateColumnWidth / dblPixelRatio
        Catch
            blnCanResize = False
        End Try

        If blnCanResize Then
            'if that worked, do the rest

            'first resize the date label
            lblDate.Width = dblJournalWidth * _
                            Constants.DateColumnWidth

            'Type column and label
            shtJournal.ActiveSheet.Cells(1, Constants.TypeColumn).ColumnWidth = _
                            dblJournalWidth * _
                            Constants.TypeColumnWidth / dblPixelRatio
            lblType.Width = dblJournalWidth * _
                            Constants.TypeColumnWidth

            'Accounts column # 1 and label
            shtJournal.ActiveSheet.Cells(1, Constants.AccountsColumn).ColumnWidth = _
                            dblJournalWidth * _
                            Constants.AccountsDebitWidth / dblPixelRatio
            lblAccounts.Width = dblJournalWidth * _
                            (Constants.AccountsDebitWidth + _
                            Constants.AccountsCreditWidth)
            'Accounts column # 2
            shtJournal.ActiveSheet.Cells(1, Constants.AccountsColumn + _
                            Constants.CreditsColumn - Constants.DebitsColumn).ColumnWidth = _
                            dblJournalWidth * _
                            Constants.AccountsCreditWidth / dblPixelRatio

            'Debits column and label
            shtJournal.ActiveSheet.Cells(1, Constants.DebitsColumn).ColumnWidth = _
                            dblJournalWidth * _
                            Constants.DebitsColumnWidth / dblPixelRatio
            lblDebits.Width = dblJournalWidth * _
                            Constants.DebitsColumnWidth

            'Credits column and label
            shtJournal.ActiveSheet.Cells(1, Constants.CreditsColumn).ColumnWidth = _
                            dblJournalWidth * _
                            Constants.CreditsColumnWidth / dblPixelRatio
            lblCredits.Width = dblJournalWidth * _
                            Constants.CreditsColumnWidth
        End If
    End Sub
    Private Sub shtJournal_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles shtJournal.Resize
        ResizeJournalSheet()
    End Sub
    Private Sub FrmJournal_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ResizeJournalSheet()
    End Sub

    Private Sub FrmGeneralJournal_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        RaiseEvent Changed()
        prjProject.RaiseChanged()
    End Sub

    Private Sub shtJournal_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles shtJournal.Enter

    End Sub
End Class