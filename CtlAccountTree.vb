Public Class CtlAccountTree
    Inherits FinancePlanner.CtlProjectControlBase

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        InitializeComponent()
    End Sub

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
    Friend WithEvents pnlControls As System.Windows.Forms.Panel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnRename As System.Windows.Forms.Button
    Friend WithEvents btnMove As System.Windows.Forms.Button
    Friend WithEvents tvwAccounts As System.Windows.Forms.TreeView
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlControls = New System.Windows.Forms.Panel()
        Me.btnMove = New System.Windows.Forms.Button()
        Me.btnRename = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.tvwAccounts = New System.Windows.Forms.TreeView()
        Me.pnlControls.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlControls
        '
        Me.pnlControls.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnMove, Me.btnRename, Me.btnDelete})
        Me.pnlControls.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlControls.Location = New System.Drawing.Point(0, 260)
        Me.pnlControls.Name = "pnlControls"
        Me.pnlControls.Size = New System.Drawing.Size(250, 40)
        Me.pnlControls.TabIndex = 0
        '
        'btnMove
        '
        Me.btnMove.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.btnMove.Location = New System.Drawing.Point(173, 8)
        Me.btnMove.Name = "btnMove"
        Me.btnMove.Size = New System.Drawing.Size(64, 24)
        Me.btnMove.TabIndex = 2
        Me.btnMove.Text = "Move"
        '
        'btnRename
        '
        Me.btnRename.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.btnRename.Location = New System.Drawing.Point(93, 8)
        Me.btnRename.Name = "btnRename"
        Me.btnRename.Size = New System.Drawing.Size(64, 24)
        Me.btnRename.TabIndex = 1
        Me.btnRename.Text = "Rename"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.btnDelete.Location = New System.Drawing.Point(13, 8)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(64, 24)
        Me.btnDelete.TabIndex = 0
        Me.btnDelete.Text = "Delete"
        '
        'tvwAccounts
        '
        Me.tvwAccounts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvwAccounts.FullRowSelect = True
        Me.tvwAccounts.HideSelection = False
        Me.tvwAccounts.ImageIndex = -1
        Me.tvwAccounts.Name = "tvwAccounts"
        Me.tvwAccounts.SelectedImageIndex = -1
        Me.tvwAccounts.ShowLines = False
        Me.tvwAccounts.ShowPlusMinus = False
        Me.tvwAccounts.ShowRootLines = False
        Me.tvwAccounts.Size = New System.Drawing.Size(250, 260)
        Me.tvwAccounts.TabIndex = 1
        '
        'CtlAccountTree
        '
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.tvwAccounts, Me.pnlControls})
        Me.Name = "CtlAccountTree"
        Me.Size = New System.Drawing.Size(250, 300)
        Me.pnlControls.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub New(ByRef prjProject As FinancePlanner.Project)
        MyBase.New(prjProject)

        InitializeComponent()
    End Sub

    Public Overrides ReadOnly Property Type() As FinancePlanner.ProjectControlEnum
        Get
            Return ProjectControlEnum.AccountTree
        End Get
    End Property

    Private blnShowControls As Boolean = True

    Public Property ShowControls() As Boolean
        Get
            Return blnShowControls
        End Get
        Set(ByVal Value As Boolean)
            blnShowControls = Value
            Me.pnlControls.Visible = blnShowControls
        End Set
    End Property

    Public Function AddNode(ByVal laAccount As FinancePlanner.Account) As FinancePlanner.AccountNode
        Dim laParentAccount As FinancePlanner.Account
        Dim nodNewNode As AccountNode
        Dim nodParentNode As AccountNode

        nodNewNode = New AccountNode()
        nodNewNode.Account = laAccount

        nodParentNode = FindNode(laAccount.GetParentAccount().FullPath)
        If nodParentNode Is Nothing Then
            tvwAccounts.Nodes.Add(nodNewNode)
        Else
            nodParentNode.Nodes.Add(nodNewNode)
        End If
    End Function

    Private Function FindNode(ByVal strAccountPath As String) As AccountNode
        Dim nodParentNode As AccountNode
        Dim nodTargetNode As AccountNode

        For Each nodParentNode In tvwAccounts.Nodes
            If nodParentNode.Account.FullPath = strAccountPath Then
                Return nodParentNode
            Else
                nodTargetNode = FindNode(nodParentNode, strAccountPath)
                If Not nodTargetNode Is Nothing Then Exit For
            End If
        Next

        Return nodTargetNode
    End Function

    Private Function FindNode(ByRef nodStartingNode As AccountNode, ByVal strAccountPath As String) As AccountNode
        Dim nodParentNode As AccountNode
        Dim nodTargetNode As AccountNode

        For Each nodParentNode In nodStartingNode.Nodes
            If nodParentNode.Account.FullPath = strAccountPath Then
                Return nodParentNode
            Else
                nodTargetNode = FindNode(nodParentNode, strAccountPath)
                If Not nodTargetNode Is Nothing Then Exit For
            End If
        Next
        Return nodTargetNode
    End Function

    Public Overrides Sub Render()
        Dim accAccount As FinancePlanner.Account

        For Each accAccount In prjProject.GetLedger().Accounts

        Next
    End Sub
End Class
