Public Class FrmSplash
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

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
    Friend WithEvents pbSplash As System.Windows.Forms.PictureBox
    Friend WithEvents timSplashTimer As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmSplash))
        Me.pbSplash = New System.Windows.Forms.PictureBox()
        Me.timSplashTimer = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'pbSplash
        '
        Me.pbSplash.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pbSplash.Image = CType(resources.GetObject("pbSplash.Image"), System.Drawing.Bitmap)
        Me.pbSplash.Name = "pbSplash"
        Me.pbSplash.Size = New System.Drawing.Size(500, 70)
        Me.pbSplash.TabIndex = 0
        Me.pbSplash.TabStop = False
        '
        'timSplashTimer
        '
        Me.timSplashTimer.Interval = 3000
        '
        'FrmSplash
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(500, 70)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.pbSplash})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FrmSplash"
        Me.ShowInTaskbar = False
        Me.Text = "Finance Planner"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FrmSplash_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        timSplashTimer.Start()
    End Sub

    Private Sub timSplashTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles timSplashTimer.Tick
        timSplashTimer.Stop()
        Me.Dispose()
    End Sub
End Class
