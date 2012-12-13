Public Class Ledger
    Inherits AccountTree
End Class

Public Class Budget
    Inherits AccountTree
End Class

Public Class AccountTree
    Public Sub New()
        Accounts = New Collection()
    End Sub

    Public Accounts As Collection

    Public Sub Clear()
        Accounts = New Collection()
    End Sub
    Public Sub ResetAccountBalances()
        Dim actAccount As FinancePlanner.Account

        For Each actAccount In Accounts
            actAccount.ResetBalance()
            actAccount.ResetSubAccountBalances()
        Next
    End Sub
    Public Sub AddRootAccount(ByRef accAccount As FinancePlanner.Account)
        Try
            Accounts.Add(accAccount, accAccount.Name)
        Catch
            'If already exists, ignore
        End Try
    End Sub
    Public Sub AddAccountWithFullPath(ByRef accAccount As FinancePlanner.Account, ByVal strFullPath As String)
        Dim accParentAccount As FinancePlanner.Account

        If AccountPath.Depth(strFullPath) <= 1 Then
            AddRootAccount(accAccount)
        Else
            accParentAccount = FindAccountWithFullPath(AccountPath.ParentPath(strFullPath))
            If Not accParentAccount Is Nothing Then
                accParentAccount.AddSubAccount(accAccount)
            Else
                accParentAccount = New FinancePlanner.Account()
                accParentAccount.Name = AccountPath.AccountName(AccountPath.ParentPath(strFullPath))
                accParentAccount.NormalBalance = accAccount.NormalBalance
                accParentAccount.Unit = accAccount.Unit
                accParentAccount.AddSubAccount(accAccount)
                AddAccountWithFullPath(accParentAccount, AccountPath.ParentPath(strFullPath))
            End If
        End If
    End Sub
    Public Sub RemoveRootAccount(ByVal strName As String)
        Try
            Accounts.Remove(strName)
        Catch
            'If doesn't exist, ignore
        End Try
    End Sub
    Public Sub RemoveAccountWithFullPath(ByVal strFullPath As String)
        Dim accParentAccount As FinancePlanner.Account

        If AccountPath.Depth(strFullPath) <= 1 Then
            RemoveRootAccount(strFullPath)
        Else
            accParentAccount = FindAccountWithFullPath(AccountPath.ParentPath(strFullPath))
            If Not accParentAccount Is Nothing Then
                accParentAccount.RemoveSubAccount(strFullPath)
            End If
        End If
    End Sub

    Public Function ShortestAliasFor(ByVal strFullPath As String) As String
        Dim strPathFragments() As String
        Dim intCounter As Integer
        Dim strShortestAlias As String
        Dim blnAliasFound As Boolean

        strPathFragments = strFullPath.Split(AccountPath.SubaccountDelimiter)
        strShortestAlias = String.Empty

        blnAliasFound = False
        intCounter = strPathFragments.Length - 1
        Do
            If strShortestAlias = String.Empty Then
                strShortestAlias = strPathFragments(intCounter)
            Else
                strShortestAlias = strPathFragments(intCounter) & AccountPath.SubaccountDelimiter & _
                                    strShortestAlias
            End If

            If FindAccountsWithAlias(strShortestAlias).Length <= 1 Then blnAliasFound = True

            intCounter -= 1
        Loop While intCounter >= 0 And Not blnAliasFound
    End Function
    Public Function FindAccountsWithAlias(ByVal strAlias As String) As FinancePlanner.Account()
        Dim accMatchList() As Account
        Dim accParentAccount As Account

        ReDim accMatchList(-1)

        For Each accParentAccount In Me.Accounts
            If IsAliasFor(strAlias, accParentAccount.FullPath) Then
                Utilities.AddItemToArray(accParentAccount, accMatchList)
            End If

            Utilities.AddItemToArray(accParentAccount.FindSubAccountsWithAlias(strAlias), accMatchList)
        Next

        Return accMatchList
    End Function
    Public Function FindAccountWithFullPath(ByVal strFullPath As String) As FinancePlanner.Account
        Dim accAccount As FinancePlanner.Account
        Dim strRootPath As String
        Dim strPathEnding As String
        Dim intDelimiterIndex As Integer

        intDelimiterIndex = strFullPath.IndexOf(SubaccountDelimiter)
        If intDelimiterIndex > 0 Then
            strRootPath = strFullPath.Substring(1, intDelimiterIndex - 1)
            strPathEnding = strFullPath.Substring(intDelimiterIndex + 1, strFullPath.Length - intDelimiterIndex)
        Else
            strRootPath = strFullPath
            strPathEnding = String.Empty
        End If

        For Each accAccount In Accounts
            If strRootPath.ToLower = accAccount.Name.ToLower Then
                If strPathEnding <> String.Empty Then
                    Return accAccount.FollowPath(strPathEnding)
                Else
                    Return accAccount
                End If
            End If
        Next

        Return Nothing
    End Function
End Class


Public Class Account
    Implements IComparable

    Public Const DefaultBalance As BalanceEnum = BalanceEnum.Debit
    Public Const SubAccountDelimiter As Char = AccountPath.SubaccountDelimiter

    Public Sub New()
        dblBalance = 0.0
        intQuantity = 0
        strUnit = String.Empty
        beNormalBalance = DefaultBalance
        strName = String.Empty
        accSubAccounts = New Collection()
        accParentAccount = Nothing
    End Sub

    Private dblBalance As Double
    Private intQuantity As Integer
    Private strUnit As String
    Private beNormalBalance As BalanceEnum
    Private strName As String
    Private accSubAccounts As Collection
    Private accParentAccount As Account

    Public Property Name() As String
        Get
            Return strName
        End Get
        Set(ByVal Value As String)
            strName = Value
        End Set
    End Property
    Public ReadOnly Property FullPath() As String
        Get
            If accParentAccount Is Nothing Then
                Return strName
            Else
                Return accParentAccount.FullPath & SubAccountDelimiter & strName
            End If
        End Get
    End Property
    Public Property NormalBalance() As BalanceEnum
        Get
            Return beNormalBalance
        End Get
        Set(ByVal Value As BalanceEnum)
            beNormalBalance = Value
        End Set
    End Property
    Public Property Unit() As String
        Get
            Return strUnit
        End Get
        Set(ByVal Value As String)
            strUnit = Value
        End Set
    End Property
    Public ReadOnly Property NetBalance() As Double
        Get
            Dim dblNetBalance As Double
            Dim accSubAccount As FinancePlanner.Account

            dblNetBalance = dblBalance

            For Each accSubAccount In accSubAccounts
                If accSubAccount.NormalBalance = Me.NormalBalance Then
                    dblNetBalance += accSubAccount.NetBalance
                Else
                    dblNetBalance -= accSubAccount.NetBalance
                End If
            Next
            Return dblNetBalance
        End Get
    End Property
    Public ReadOnly Property NetQuantity() As Integer
        Get
            Dim intNetQuantity As Integer
            Dim accSubAccount As FinancePlanner.Account

            intNetQuantity = intQuantity

            For Each accSubAccount In accSubAccounts
                If accSubAccount.Unit.ToLower = Me.Unit.ToLower Then
                    intNetQuantity += accSubAccount.NetQuantity
                End If
            Next
        End Get
    End Property
    Public ReadOnly Property FormattedQuantity() As String
        Get
            Dim intNetQuantity As Integer = NetQuantity
            If intNetQuantity = 1 Then
                Return CStr(intNetQuantity) & " " & MakeSingular(Unit)
            Else
                Return CStr(intNetQuantity) & " " & MakePlural(Unit)
            End If
        End Get
    End Property
    Public ReadOnly Property UnitCost() As Double
        Get
            Dim intNetQuantity = NetQuantity

            If intNetQuantity <> 0 Then
                Return NetBalance / CDbl(intNetQuantity)
            Else
                Return 0
            End If
        End Get
    End Property
    Public ReadOnly Property FormattedUnitCost() As String
        Get
            Return CStr(UnitCost).Format("c") & " Per " & MakeSingular(strUnit)
        End Get
    End Property
    Public Property NormalBalanceAsString()
        Get
            Select Case beNormalBalance
                Case BalanceEnum.Debit
                    Return "Debit"
                Case BalanceEnum.Credit
                    Return "Credit"
            End Select
        End Get
        Set(ByVal Value)
            Select Case Value.ToLower
                Case "debit", "deb", "de", "d"
                    beNormalBalance = BalanceEnum.Debit
                Case "credit", "cred", "cre", "cr", "c"
                    beNormalBalance = BalanceEnum.Credit
                Case Else
                    Throw New ArgumentException("Unknown Balance Type.  Must Be Debit or Credit.")
            End Select
        End Set
    End Property
    Public Function GetParentAccount() As Account
        Return accParentAccount
    End Function

    Public Function GetSubAccounts() As Collection
        Return accSubAccounts
    End Function
    Public Sub MoveTo(ByRef accParent As Account)
        Me.GetParentAccount().RemoveSubAccount(Me.Name)
        accParent.AddSubAccount(Me)
    End Sub
    Public Sub AddSubAccount(ByVal accAccount As Account)
        accAccount.accParentAccount = Me
        Try
            accSubAccounts.Add(accAccount, accAccount.Name)
        Catch
            'If already exists, ignore
        End Try
    End Sub
    Public Sub RemoveSubAccount(ByVal strName As String)
        Try
            accSubAccounts.Remove(strName)
        Catch
            'if doesn't exist, ignore
        End Try
    End Sub
    Public Function FindSubAccountsWithAlias(ByVal strAlias As String) As FinancePlanner.Account()
        Dim accMatchList() As Account
        Dim accParentAccount As Account

        ReDim accMatchList(-1)

        For Each accParentAccount In Me.GetSubAccounts
            If IsAliasFor(strAlias, accParentAccount.FullPath) Then
                Utilities.AddItemToArray(accParentAccount, accMatchList)
            End If

            Utilities.AddItemToArray(accParentAccount.FindSubAccountsWithAlias(strAlias), accMatchList)
        Next

        Return accMatchList
    End Function
    Public Function FollowPath(ByVal strPath As String) As FinancePlanner.Account
        Dim accAccount As FinancePlanner.Account
        Dim strRootPath As String
        Dim strPathEnding As String
        Dim intDelimiterIndex As Integer

        intDelimiterIndex = strPath.IndexOf(SubAccountDelimiter)
        If intDelimiterIndex > 0 Then
            strRootPath = strPath.Substring(1, intDelimiterIndex - 1)
            strPathEnding = strPath.Substring(intDelimiterIndex + 1, strPath.Length - intDelimiterIndex)
        Else
            strRootPath = strPath
            strPathEnding = String.Empty
        End If

        For Each accAccount In accSubAccounts
            If strRootPath.ToLower = accAccount.Name.ToLower Then
                If strPathEnding <> String.Empty Then
                    Return accAccount.FollowPath(strPathEnding)
                Else
                    Return accAccount
                End If
            End If
        Next

        Return Nothing
    End Function
    Public Function ResetBalance()
        dblBalance = 0.0
        intQuantity = 0
    End Function
    Public Function ResetSubAccountBalances()
        Dim accSubAccount As FinancePlanner.Account

        For Each accSubAccount In accSubAccounts
            accSubAccount.ResetBalance()
            accSubAccount.ResetSubAccountBalances()
        Next
    End Function
    Public Sub PostBalance(ByVal dblValue As Double, ByVal beBalanceType As BalanceEnum)
        If beBalanceType = beNormalBalance Then
            dblBalance += dblValue
        Else
            dblBalance -= dblValue
        End If
    End Sub
    Public Sub PostQuantity(ByVal intValue As Integer, ByVal beBalanceType As BalanceEnum)
        If beBalanceType = beNormalBalance Then
            intQuantity += intValue
        Else
            intQuantity -= intValue
        End If
    End Sub

    Public Function ComparePaths(ByVal obj As Object) As Integer Implements IComparable.CompareTo
        If TypeOf (obj) Is FinancePlanner.Account Then
            If FullPath < CType(obj, FinancePlanner.Account).FullPath Then
                Return -1
            ElseIf FullPath > CType(obj, FinancePlanner.Account).FullPath Then
                Return 1
            Else
                Return 0
            End If
        Else
            Throw New ArgumentException("Incorrect Argument Type Supplied")
        End If
    End Function
End Class


Public Class BudgetAccount
    Inherits FinancePlanner.Account

    Private baType As BudgetAccountEnum
    Private strReferencePath As String

    Public Property Type() As BudgetAccountEnum
        Get
            Return baType
        End Get
        Set(ByVal Value As BudgetAccountEnum)
            baType = Value
        End Set
    End Property
    Public Property ReferencePath() As String
        Get
            Return strReferencePath
        End Get
        Set(ByVal Value As String)
            strReferencePath = Value
        End Set
    End Property
End Class