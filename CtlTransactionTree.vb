Public Class CtlTransactionTree
    Inherits FinancePlanner.CtlProjectControlBase

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

    Public Sub New(ByRef prjProject As FinancePlanner.Project)
        MyBase.New(prjProject)

        InitializeComponent()
    End Sub

    Public Overrides ReadOnly Property Type() As FinancePlanner.ProjectControlEnum
        Get
            Return ProjectControlEnum.TransactionTree
        End Get
    End Property

    Public Overrides Sub Render()
        MsgBox("Rendering Transaction Tree")
        'TODO: Revise Transaction Tree Render
        'Dim jeEntry As FinancePlanner.JournalEntry
        'Dim jtTransaction As FinancePlanner.JournalTransaction
        'Dim jtRelatedTransaction As FinancePlanner.JournalTransaction
        'Dim laAccount As FinancePlanner.LedgerAccount
        'Dim nodRootNode As TreeNode
        'Dim nodDebitsNode As TreeNode
        'Dim nodCreditsNode As TreeNode
        'Dim nodTempNode As TreeNode
        'Dim nodTempNode2 As TreeNode
        'Dim dblDebits As Double, dblCredits As Double
        'Dim intDelimiterIndex As Integer

        'Me.Nodes.Clear()
        'nodRootNode = Me.Nodes.Add(drwSelectedAccount.Path)
        'nodRootNode.Text += ": " & drwSelectedAccount.Balance.ToString("c")
        'nodRootNode.Tag = drwSelectedAccount
        'If drwSelectedAccount.Balance < 0 Then
        '    nodRootNode.ForeColor = Color.DarkBlue
        'ElseIf drwSelectedAccount.Balance > 0 Then
        '    nodRootNode.ForeColor = Color.DarkRed
        'Else
        '    nodRootNode.ForeColor = Color.Black
        'End If
        'nodDebitsNode = nodRootNode.Nodes.Add("Debits")
        'nodCreditsNode = nodRootNode.Nodes.Add("Credits")
        'nodDebitsNode.ForeColor = Color.DarkRed
        'nodCreditsNode.ForeColor = Color.DarkBlue
        'dblDebits = 0
        'dblCredits = 0

        'For Each drwAccount In drwLedger.GetLedgerAccountRows()
        '    If IsSubAccountPath(drwAccount.Path, drwSelectedAccount.Path, True) Then
        '        For Each drwEntry In drwJournal.GetEntryRows()
        '            For Each drwTransaction In drwEntry.GetTransactionRows()
        '                If IsAliasFor(drwTransaction.Path, drwAccount.Path) Then
        '                    nodTempNode = Nothing
        '                    If Not drwTransaction.IsDebitValueNull() Then
        '                        nodTempNode = nodDebitsNode.Nodes.Add(drwEntry.Comment)
        '                        nodTempNode.ForeColor = Color.DarkRed
        '                        dblDebits += drwTransaction.DebitValue
        '                        If nodTempNode.Text = String.Empty Then
        '                            nodTempNode.Text = "(No Comment)"
        '                        End If
        '                        nodTempNode.Text += ": " & drwTransaction.DebitValue.ToString("c")
        '                        nodTempNode.Nodes.Add("Date: " & drwEntry._Date.ToLongDateString)
        '                        If Not drwAccount Is drwSelectedAccount Then
        '                            nodTempNode2 = nodTempNode.Nodes.Add("Account Debited: " & _
        '                                                                 drwTransaction.Path)
        '                            nodTempNode2.ForeColor = Color.DarkRed
        '                        End If
        '                        For Each drwRelatedTransaction In drwEntry.GetTransactionRows()
        '                            If Not drwRelatedTransaction.IsCreditValueNull() Then
        '                                nodTempNode2 = nodTempNode.Nodes.Add("Credited " & _
        '                                                        drwRelatedTransaction.Path & " for " & _
        '                                                        drwRelatedTransaction.CreditValue.ToString("c"))
        '                                nodTempNode2.ForeColor = Color.DarkBlue
        '                            End If
        '                        Next
        '                    End If
        '                    If Not drwTransaction.IsCreditValueNull() Then
        '                        nodTempNode = nodCreditsNode.Nodes.Add(drwEntry.Comment)
        '                        nodTempNode.ForeColor = Color.DarkBlue
        '                        dblCredits += drwTransaction.CreditValue
        '                        If nodTempNode.Text = String.Empty Then
        '                            nodTempNode.Text = "(No Comment)"
        '                        End If
        '                        nodTempNode.Text += ": " & drwTransaction.CreditValue.ToString("c")
        '                        nodTempNode.Nodes.Add("Date: " & drwEntry._Date.ToLongDateString)
        '                        If Not drwAccount Is drwSelectedAccount Then
        '                            nodTempNode2 = nodTempNode.Nodes.Add("Account Credited: " & _
        '                                                                 drwTransaction.Path)
        '                            nodTempNode2.ForeColor = Color.DarkBlue
        '                        End If
        '                        For Each drwRelatedTransaction In drwEntry.GetTransactionRows()
        '                            If Not drwRelatedTransaction.IsDebitValueNull() Then
        '                                nodTempNode2 = nodTempNode.Nodes.Add("Debited " & _
        '                                                        drwRelatedTransaction.Path & " for " & _
        '                                                        drwRelatedTransaction.DebitValue.ToString("c"))
        '                                nodTempNode2.ForeColor = Color.DarkRed
        '                            End If
        '                        Next
        '                    End If
        '                End If
        '            Next
        '        Next
        '    End If
        'Next

        'nodDebitsNode.Text += ": " & dblDebits.ToString("c")
        'nodCreditsNode.Text += ": " & dblCredits.ToString("c")
        'nodRootNode.Expand()
        'If drwSelectedAccount.Balance > 0 Then
        '    nodDebitsNode.Expand()
        'Else
        '    nodCreditsNode.Expand()
        'End If

    End Sub
End Class
