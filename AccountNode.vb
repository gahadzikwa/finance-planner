Public Class AccountNode
    Inherits System.Windows.Forms.TreeNode

    Public Sub New()
        MyBase.New()

    End Sub

    Public Property Account() As FinancePlanner.Account
        Get
            Return actAccount
        End Get
        Set(ByVal Value As FinancePlanner.Account)
            actAccount = Value
            Me.Text = actAccount.Name & ": " & actAccount.NetBalance.ToString.Format("c")
        End Set
    End Property

    Private actAccount As FinancePlanner.Account

End Class
