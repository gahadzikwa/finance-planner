Public Class FrmGeneralBudget
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
        Me.Text = "FrmBudget"
    End Sub

#End Region

    Public Sub New(ByRef prjProject As FinancePlanner.Project)
        MyBase.New(prjProject)

        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub

    Public Overrides ReadOnly Property Type() As FinancePlanner.ProjectWindowEnum
        Get
            Return FinancePlanner.ProjectWindowEnum.GeneralBudget
        End Get
    End Property

    Public Overrides ReadOnly Property Title() As String
        Get
            Return "General Budget"
        End Get
    End Property

    Public Sub Parse() Implements IParses.Parse
        MsgBox("Parsing General Budget")
        'TODO: Implement Budget Parse
    End Sub

    Public Overrides Sub Render()
        MsgBox("Rendering General Budget")
        'TODO: Implement Budget Render
    End Sub

    Event Changed() Implements IParses.Changed

End Class
