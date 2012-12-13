Public Class FrmProjectWindowBase
    Inherits System.Windows.Forms.Form
    Implements FinancePlanner.IProjectWindow

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
        '
        'FrmProjectWindowBase
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 370)
        Me.MinimizeBox = False
        Me.Name = "FrmProjectWindowBase"
        Me.ShowInTaskbar = False
        Me.Text = "FrmProjectWindow"

    End Sub

#End Region

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub

    Public Sub New(ByRef prjProject As FinancePlanner.Project)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        Me.prjProject = prjProject
    End Sub

    Public Overridable ReadOnly Property Type() As FinancePlanner.ProjectWindowEnum Implements IProjectWindow.Type
        Get
            Return ProjectWindowEnum.ProjectWindowBase
        End Get
    End Property
    Public Overridable ReadOnly Property Title() As String Implements IProjectWindow.Title
        Get
            Return "Project Window"
        End Get
    End Property
    Public Overridable Sub Render() Implements IProjectWindow.Render
        'Do nothing in base class
    End Sub
    'Public Overridable ReadOnly Property ParseType() As FinancePlanner.ParseTypeEnum Implements IProjectWindow.ParseType
    '    Get
    '        Return ParseTypeEnum.NoParse
    '    End Get
    'End Property

    'Project
    Protected WithEvents prjProject As FinancePlanner.Project

    'Form must be pointed to the correct project
    Public Function SetProject(ByRef prjProject As FinancePlanner.Project)
        Me.prjProject = prjProject
    End Function

    Protected Sub prjProject_Changed() Handles prjProject.Changed
        RefreshTitle()
    End Sub

    Private Sub FrmProjectWindow_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Render()
        RefreshTitle()
    End Sub

    Private Sub RefreshTitle()
        If Not prjProject Is Nothing Then
            Me.Text = prjProject.Title & " - " & Me.Title
        Else
            Me.Text = Me.Title
        End If
    End Sub

    Private Sub FrmProjectWindowBase_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        If Not prjProject Is Nothing Then
            prjProject.RemoveProjectWindow(Me.Type)
        End If
    End Sub
End Class
