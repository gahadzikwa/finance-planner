Public Class CtlProjectControlBase
    Inherits System.Windows.Forms.UserControl
    Implements FinancePlanner.IProjectControl

#Region " Windows Form Designer generated code "

    'UserControl overrides dispose to clean up the component list.
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
    End Sub

#End Region

    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub

    Public Sub New(ByRef prjProject As FinancePlanner.Project)
        MyBase.New()
        InitializeComponent()
        Me.prjProject = prjProject
    End Sub

    Protected prjProject As FinancePlanner.Project

    Public Overridable ReadOnly Property Type() As FinancePlanner.ProjectControlEnum Implements IProjectControl.Type
        Get
            Return ProjectControlEnum.ProjectControlBase
        End Get
    End Property

    Public Overridable Sub Render() Implements IProjectControl.Render
        'Do nothing in base class
    End Sub

End Class
