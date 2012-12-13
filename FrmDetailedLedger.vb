Public Class FrmDetailedLedger
    Inherits FinancePlanner.FrmProjectWindowBase

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
    Friend WithEvents CtlAccountTree1 As FinancePlanner.CtlAccountTree
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.CtlAccountTree1 = New FinancePlanner.CtlAccountTree()
        Me.SuspendLayout()
        '
        'CtlAccountTree1
        '
        Me.CtlAccountTree1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CtlAccountTree1.Name = "CtlAccountTree1"
        Me.CtlAccountTree1.ShowControls = True
        Me.CtlAccountTree1.Size = New System.Drawing.Size(592, 370)
        Me.CtlAccountTree1.TabIndex = 0
        '
        'FrmDetailedLedger
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 370)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.CtlAccountTree1})
        Me.Name = "FrmDetailedLedger"
        Me.Text = "FrmDetailedLedger"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub New(ByRef prjProject As FinancePlanner.Project)
        MyBase.New(prjProject)
        InitializeComponent()
    End Sub

    'Overrides
    Public Overrides ReadOnly Property Type() As FinancePlanner.ProjectWindowEnum
        Get
            Return FinancePlanner.ProjectWindowEnum.DetailedLedger
        End Get
    End Property

    Public Overrides ReadOnly Property Title() As String
        Get
            Return "Detailed Ledger"
        End Get
    End Property

    'Form Rendering
    Public Overrides Sub Render()
        MsgBox("Rendering Detailed Ledger")
        'TODO: Implement Detailed Ledger Render
    End Sub

    ''Context Menu Items
    'Private Sub itmDeleteAccount_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles itmDeleteAccount.Click
    '    DeleteSelectedAccount()
    'End Sub
    'Private Sub itmRenameAccount_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles itmRenameAccount.Click
    '    RenameSelectedAccount()
    'End Sub
    'Private Sub itmMoveAccount_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles itmMoveAccount.Click
    '    MoveSelectedAccount()
    'End Sub

    ''Account Editing
    'Private Sub DeleteSelectedAccount()
    '    If Not tvwDetailedAccounts.SelectedNode Is Nothing Then
    '        Select Case MessageBox.Show("Are you sure you want to delete the account """ & _
    '                                    tvwDetailedAccounts.SelectedNode.Tag.Path & _
    '                                    """ and all of its subaccounts?", _
    '                                    "Delete Account", MessageBoxButtons.YesNo, _
    '                                    MessageBoxIcon.Question)
    '            Case DialogResult.Yes
    '                prjProject.DeleteLedgerAccount(tvwDetailedAccounts.SelectedNode.Tag.Path)
    '                Render()
    '            Case Else
    '                'do nothing
    '        End Select
    '    End If
    'End Sub
    'Private Sub RenameSelectedAccount()
    '    Dim strNewName As String
    '    Dim drwSelectedAccount As FinancePlanner.ProjectData.LedgerAccountRow

    '    If Not tvwDetailedAccounts.SelectedNode Is Nothing Then
    '        drwSelectedAccount = tvwDetailedAccounts.SelectedNode.Tag

    '        strNewName = InputBox("Please enter a new name for the account """ & _
    '                                          drwSelectedAccount.Path & """:", _
    '                                          "Rename Account", _
    '                                          AccountNameFromPath(drwSelectedAccount.Path))

    '        If strNewName <> String.Empty Then

    '            strNewName = AccountPathFromPath(drwSelectedAccount.Path) & _
    '                                     SubaccountDelimiter & strNewName

    '            'MsgBox(drwSelectedAccount.Path & "; " & strNewName)

    '            prjProject.ReplaceTransactionPaths(drwSelectedAccount.Path, strNewName)
    '            prjProject.ReplaceLedgerAccountPaths(drwSelectedAccount.Path, strNewName)

    '            Render()
    '            prjProject.JournalForm().Render()
    '        End If
    '    End If
    'End Sub
    'Private Sub MoveSelectedAccount()
    '    Dim strNewName As String
    '    Dim drwSelectedAccount As FinancePlanner.ProjectData.LedgerAccountRow

    '    If Not tvwDetailedAccounts.SelectedNode Is Nothing Then
    '        drwSelectedAccount = tvwDetailedAccounts.SelectedNode.Tag

    '        strNewName = InputBox("Please enter a new path for the account """ & _
    '                                          drwSelectedAccount.Path & """:", _
    '                                          "Move Account", _
    '                                          AccountPathFromPath(drwSelectedAccount.Path))

    '        If strNewName <> String.Empty Then

    '            strNewName = strNewName & SubaccountDelimiter & _
    '                         AccountNameFromPath(drwSelectedAccount.Path)

    '            'MsgBox(drwSelectedAccount.Path & "; " & strNewName)

    '            prjProject.ReplaceTransactionPaths(drwSelectedAccount.Path, strNewName)
    '            prjProject.ReplaceLedgerAccountPaths(drwSelectedAccount.Path, strNewName)

    '            Render()
    '            prjProject.JournalForm().Render()
    '        End If
    '    End If

    'End Sub
End Class
