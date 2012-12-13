Public Class FrmProjectProperties
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
    Friend WithEvents tbcProperties As System.Windows.Forms.TabControl
    Friend WithEvents tbpGeneral As System.Windows.Forms.TabPage
    Friend WithEvents tbpJournal As System.Windows.Forms.TabPage
    Friend WithEvents tbpLedger As System.Windows.Forms.TabPage
    Friend WithEvents bnlButtons As System.Windows.Forms.Panel
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblProjectName As System.Windows.Forms.Label
    Friend WithEvents lblJournalName As System.Windows.Forms.Label
    Friend WithEvents lblLedgerName As System.Windows.Forms.Label
    Friend WithEvents tbpBudget As System.Windows.Forms.TabPage
    Friend WithEvents lblBudgetName As System.Windows.Forms.Label
    Friend WithEvents txtProjectTitle As System.Windows.Forms.TextBox
    Friend WithEvents txtJournalTitle As System.Windows.Forms.TextBox
    Friend WithEvents txtLedgerTitle As System.Windows.Forms.TextBox
    Friend WithEvents txtBudgetTitle As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmProjectProperties))
        Me.tbcProperties = New System.Windows.Forms.TabControl()
        Me.tbpGeneral = New System.Windows.Forms.TabPage()
        Me.txtProjectTitle = New System.Windows.Forms.TextBox()
        Me.lblProjectName = New System.Windows.Forms.Label()
        Me.tbpJournal = New System.Windows.Forms.TabPage()
        Me.txtJournalTitle = New System.Windows.Forms.TextBox()
        Me.lblJournalName = New System.Windows.Forms.Label()
        Me.tbpLedger = New System.Windows.Forms.TabPage()
        Me.txtLedgerTitle = New System.Windows.Forms.TextBox()
        Me.lblLedgerName = New System.Windows.Forms.Label()
        Me.tbpBudget = New System.Windows.Forms.TabPage()
        Me.txtBudgetTitle = New System.Windows.Forms.TextBox()
        Me.lblBudgetName = New System.Windows.Forms.Label()
        Me.bnlButtons = New System.Windows.Forms.Panel()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.tbcProperties.SuspendLayout()
        Me.tbpGeneral.SuspendLayout()
        Me.tbpJournal.SuspendLayout()
        Me.tbpLedger.SuspendLayout()
        Me.tbpBudget.SuspendLayout()
        Me.bnlButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbcProperties
        '
        Me.tbcProperties.Controls.AddRange(New System.Windows.Forms.Control() {Me.tbpGeneral, Me.tbpJournal, Me.tbpLedger, Me.tbpBudget})
        Me.tbcProperties.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbcProperties.Name = "tbcProperties"
        Me.tbcProperties.SelectedIndex = 0
        Me.tbcProperties.Size = New System.Drawing.Size(392, 232)
        Me.tbcProperties.TabIndex = 0
        '
        'tbpGeneral
        '
        Me.tbpGeneral.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtProjectTitle, Me.lblProjectName})
        Me.tbpGeneral.Location = New System.Drawing.Point(4, 22)
        Me.tbpGeneral.Name = "tbpGeneral"
        Me.tbpGeneral.Size = New System.Drawing.Size(384, 206)
        Me.tbpGeneral.TabIndex = 0
        Me.tbpGeneral.Text = "General"
        '
        'txtProjectTitle
        '
        Me.txtProjectTitle.Location = New System.Drawing.Point(112, 8)
        Me.txtProjectTitle.Name = "txtProjectTitle"
        Me.txtProjectTitle.Size = New System.Drawing.Size(264, 20)
        Me.txtProjectTitle.TabIndex = 1
        Me.txtProjectTitle.Text = ""
        '
        'lblProjectName
        '
        Me.lblProjectName.Location = New System.Drawing.Point(8, 8)
        Me.lblProjectName.Name = "lblProjectName"
        Me.lblProjectName.Size = New System.Drawing.Size(100, 24)
        Me.lblProjectName.TabIndex = 0
        Me.lblProjectName.Text = "Project Name:"
        Me.lblProjectName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tbpJournal
        '
        Me.tbpJournal.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtJournalTitle, Me.lblJournalName})
        Me.tbpJournal.Location = New System.Drawing.Point(4, 22)
        Me.tbpJournal.Name = "tbpJournal"
        Me.tbpJournal.Size = New System.Drawing.Size(384, 206)
        Me.tbpJournal.TabIndex = 1
        Me.tbpJournal.Text = "Journal"
        '
        'txtJournalTitle
        '
        Me.txtJournalTitle.Location = New System.Drawing.Point(112, 8)
        Me.txtJournalTitle.Name = "txtJournalTitle"
        Me.txtJournalTitle.Size = New System.Drawing.Size(264, 20)
        Me.txtJournalTitle.TabIndex = 3
        Me.txtJournalTitle.Text = ""
        '
        'lblJournalName
        '
        Me.lblJournalName.Location = New System.Drawing.Point(8, 8)
        Me.lblJournalName.Name = "lblJournalName"
        Me.lblJournalName.Size = New System.Drawing.Size(100, 24)
        Me.lblJournalName.TabIndex = 2
        Me.lblJournalName.Text = "Journal Name:"
        Me.lblJournalName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tbpLedger
        '
        Me.tbpLedger.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtLedgerTitle, Me.lblLedgerName})
        Me.tbpLedger.Location = New System.Drawing.Point(4, 22)
        Me.tbpLedger.Name = "tbpLedger"
        Me.tbpLedger.Size = New System.Drawing.Size(384, 206)
        Me.tbpLedger.TabIndex = 2
        Me.tbpLedger.Text = "Ledger"
        '
        'txtLedgerTitle
        '
        Me.txtLedgerTitle.Location = New System.Drawing.Point(112, 8)
        Me.txtLedgerTitle.Name = "txtLedgerTitle"
        Me.txtLedgerTitle.Size = New System.Drawing.Size(264, 20)
        Me.txtLedgerTitle.TabIndex = 3
        Me.txtLedgerTitle.Text = ""
        '
        'lblLedgerName
        '
        Me.lblLedgerName.Location = New System.Drawing.Point(8, 8)
        Me.lblLedgerName.Name = "lblLedgerName"
        Me.lblLedgerName.Size = New System.Drawing.Size(100, 24)
        Me.lblLedgerName.TabIndex = 2
        Me.lblLedgerName.Text = "Ledger Name:"
        Me.lblLedgerName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tbpBudget
        '
        Me.tbpBudget.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtBudgetTitle, Me.lblBudgetName})
        Me.tbpBudget.Location = New System.Drawing.Point(4, 22)
        Me.tbpBudget.Name = "tbpBudget"
        Me.tbpBudget.Size = New System.Drawing.Size(384, 206)
        Me.tbpBudget.TabIndex = 3
        Me.tbpBudget.Text = "Budget"
        '
        'txtBudgetTitle
        '
        Me.txtBudgetTitle.Location = New System.Drawing.Point(112, 8)
        Me.txtBudgetTitle.Name = "txtBudgetTitle"
        Me.txtBudgetTitle.Size = New System.Drawing.Size(264, 20)
        Me.txtBudgetTitle.TabIndex = 3
        Me.txtBudgetTitle.Text = ""
        '
        'lblBudgetName
        '
        Me.lblBudgetName.Location = New System.Drawing.Point(8, 8)
        Me.lblBudgetName.Name = "lblBudgetName"
        Me.lblBudgetName.Size = New System.Drawing.Size(100, 24)
        Me.lblBudgetName.TabIndex = 2
        Me.lblBudgetName.Text = "Budget Name:"
        Me.lblBudgetName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'bnlButtons
        '
        Me.bnlButtons.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnCancel, Me.btnSave})
        Me.bnlButtons.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.bnlButtons.Location = New System.Drawing.Point(0, 232)
        Me.bnlButtons.Name = "bnlButtons"
        Me.bnlButtons.Size = New System.Drawing.Size(392, 40)
        Me.bnlButtons.TabIndex = 1
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(216, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(304, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "Save"
        '
        'FrmProjectProperties
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(392, 272)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.tbcProperties, Me.bnlButtons})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmProjectProperties"
        Me.ShowInTaskbar = False
        Me.Text = "Project Properties"
        Me.TopMost = True
        Me.tbcProperties.ResumeLayout(False)
        Me.tbpGeneral.ResumeLayout(False)
        Me.tbpJournal.ResumeLayout(False)
        Me.tbpLedger.ResumeLayout(False)
        Me.tbpBudget.ResumeLayout(False)
        Me.bnlButtons.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub New(ByRef prjProject As FinancePlanner.Project)
        MyBase.New(prjProject)
        InitializeComponent()
    End Sub

    Public Overrides ReadOnly Property Type() As FinancePlanner.ProjectWindowEnum
        Get
            Return ProjectWindowEnum.ProjectProperties
        End Get
    End Property

    Public Overrides ReadOnly Property Title() As String
        Get
            Return "Properties"
        End Get
    End Property

    Event Changed() Implements IParses.Changed

    'Buttons
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Parse()
        Me.Close()
    End Sub

    'Subs
    Public Sub Parse() Implements IParses.Parse
        prjProject.Title = txtProjectTitle.Text
    End Sub

    Public Overrides Sub Render()
        txtProjectTitle.Text = prjProject.Title
    End Sub
End Class
