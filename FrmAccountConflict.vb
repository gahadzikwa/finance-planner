Public Class FrmAccountConflict
    Inherits FinancePlanner.FrmProjectWindowBase

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
    Friend WithEvents grpOptions As System.Windows.Forms.GroupBox
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents lblTransaction As System.Windows.Forms.Label
    Friend WithEvents txtAccountPath As System.Windows.Forms.TextBox
    Friend WithEvents tvwAccounts As FinancePlanner.CtlAccountTree
    Friend WithEvents CtlAccountTree1 As FinancePlanner.CtlAccountTree
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.grpOptions = New System.Windows.Forms.GroupBox()
        Me.CtlAccountTree1 = New FinancePlanner.CtlAccountTree()
        Me.txtAccountPath = New System.Windows.Forms.TextBox()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.lblTransaction = New System.Windows.Forms.Label()
        Me.grpOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpOptions
        '
        Me.grpOptions.Controls.AddRange(New System.Windows.Forms.Control() {Me.CtlAccountTree1, Me.txtAccountPath, Me.btnOk})
        Me.grpOptions.Location = New System.Drawing.Point(8, 80)
        Me.grpOptions.Name = "grpOptions"
        Me.grpOptions.Size = New System.Drawing.Size(392, 304)
        Me.grpOptions.TabIndex = 1
        Me.grpOptions.TabStop = False
        Me.grpOptions.Text = "Please verify the location of this account:"
        '
        'CtlAccountTree1
        '
        Me.CtlAccountTree1.Location = New System.Drawing.Point(8, 16)
        Me.CtlAccountTree1.Name = "CtlAccountTree1"
        Me.CtlAccountTree1.ShowControls = False
        Me.CtlAccountTree1.Size = New System.Drawing.Size(376, 248)
        Me.CtlAccountTree1.TabIndex = 6
        '
        'txtAccountPath
        '
        Me.txtAccountPath.Location = New System.Drawing.Point(8, 272)
        Me.txtAccountPath.Multiline = True
        Me.txtAccountPath.Name = "txtAccountPath"
        Me.txtAccountPath.Size = New System.Drawing.Size(296, 24)
        Me.txtAccountPath.TabIndex = 5
        Me.txtAccountPath.Text = "(Root Account)"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(312, 272)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 24)
        Me.btnOk.TabIndex = 3
        Me.btnOk.Text = "Ok"
        '
        'lblDescription
        '
        Me.lblDescription.Location = New System.Drawing.Point(8, 8)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(392, 24)
        Me.lblDescription.TabIndex = 2
        '
        'lblTransaction
        '
        Me.lblTransaction.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTransaction.Location = New System.Drawing.Point(8, 40)
        Me.lblTransaction.Name = "lblTransaction"
        Me.lblTransaction.Size = New System.Drawing.Size(392, 32)
        Me.lblTransaction.TabIndex = 4
        Me.lblTransaction.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'FrmAccountConflict
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(402, 388)
        Me.ControlBox = False
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblTransaction, Me.lblDescription, Me.grpOptions})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "FrmAccountConflict"
        Me.Text = "Account Conflict Found"
        Me.grpOptions.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub New(ByRef prjProject As FinancePlanner.Project)
        MyBase.New(prjProject)

        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub

    Public Overrides ReadOnly Property Type() As FinancePlanner.ProjectWindowEnum
        Get
            Return ProjectWindowEnum.AccountConflict
        End Get
    End Property

    Public Overrides ReadOnly Property Title() As String
        Get
            Return "Account Conflict Found"
        End Get
    End Property

    Private Const strRootAccountText As String = "(Root Account)"

End Class
