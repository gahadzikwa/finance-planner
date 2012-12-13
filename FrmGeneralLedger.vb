Public Class FrmGeneralLedger
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
        Me.Text = "FrmGeneralLedger"
    End Sub

#End Region

    Shadows Class Constants
        Public Const StartingRow = 3, _
                     StartingColumn = 2, _
                     DebitsColumn = StartingColumn + 10, _
                     CreditsColumn = DebitsColumn + 1, _
                     BaseFontSize = 12, _
                     NumberFormat As String = "$##,#0.00;;""----"""
    End Class

    Public Sub New(ByRef prjProject As FinancePlanner.Project)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.prjProject = prjProject
    End Sub

    Public Overrides ReadOnly Property Type() As FinancePlanner.ProjectWindowEnum
        Get
            Return ProjectWindowEnum.GeneralLedger
        End Get
    End Property
    Public Overrides ReadOnly Property Title() As String
        Get
            Return "General Ledger"
        End Get
    End Property


    Public Overrides Sub Render()
        'MsgBox("Rendering General Ledger")
    End Sub

End Class
