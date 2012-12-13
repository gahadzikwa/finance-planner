Public Class FrmTest
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
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents rtbResults As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlBottom = New System.Windows.Forms.Panel()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.rtbResults = New System.Windows.Forms.RichTextBox()
        Me.pnlBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlBottom
        '
        Me.pnlBottom.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnOk})
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 227)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(292, 40)
        Me.pnlBottom.TabIndex = 1
        '
        'btnOk
        '
        Me.btnOk.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.btnOk.Location = New System.Drawing.Point(112, 8)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.TabIndex = 0
        Me.btnOk.Text = "Ok"
        '
        'rtbResults
        '
        Me.rtbResults.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rtbResults.Name = "rtbResults"
        Me.rtbResults.Size = New System.Drawing.Size(292, 227)
        Me.rtbResults.TabIndex = 2
        Me.rtbResults.Text = ""
        '
        'FrmTest
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 267)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.rtbResults, Me.pnlBottom})
        Me.Name = "FrmTest"
        Me.Text = "FrmTest"
        Me.pnlBottom.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub New(ByRef prjProject As FinancePlanner.Project)
        MyBase.New(prjProject)
        InitializeComponent()
    End Sub

    Public Overrides ReadOnly Property Type() As FinancePlanner.ProjectWindowEnum
        Get
            Return ProjectWindowEnum.TestWindow
        End Get
    End Property

    Public Overrides ReadOnly Property Title() As String
        Get
            Return "Test Window"
        End Get
    End Property

    Private Sub FrmTest_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'DisplayResult("Expenses-Traveling-Fuel").IsSubaccountPathOfAlias(New AccountPath("Traveling-Fuel")))
        'DisplayResult(PathUtil.IsSubaccountPathOfAlias("Expenses-Traveling-Fuel", "Traveling-Fuel", False))
        'DisplayResult(PathUtil.IsSubaccountPathOfAlias("Equity-Expenses-Traveling-Fuel", "Expenses-Traveling", True))
        'DisplayResult(PathUtil.IsSubaccountPathOfAlias("Equity-Expenses-Traveling-Fuel", "Expenses-Traveling", False))

        'DisplayResult(ChangeAbbreviatedPath("Expenses-Promotion-Website", "Equity-Expenses-Promotion-Website", "Equity-Spending-Promotion-Website"))
        'DisplayResult(ChangeAbbreviatedPath("Net Income-Expenses-Advertising", "Equity-Income-GCC", "Equity-Revenue-GCC"))
        'DisplayResult(MakeSingular("Sidneys"))
        'DisplayResult(MakeSingular("Emergencies"))
        'DisplayResult(MakeSingular("Classes"))
        'DisplayResult(MakeSingular("Houses"))
        'DisplayResult(MakeSingular("Vaginas"))

        'DisplayResult(MakePlural("Ounce"))
        'DisplayResult(MakePlural("Widget"))
        'DisplayResult(MakePlural("Trees"))
        'DisplayResult(MakePlural("Glass"))
        'DisplayResult(MakePlural("Poopy"))
    End Sub

    Sub DisplayResult(ByVal Result As Object)
        rtbResults.Text += CType(Result, String) & ControlChars.CrLf
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Me.Dispose()
    End Sub
End Class
