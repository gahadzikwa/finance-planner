Public Class Journal
    Public Enum CompareByEnum
        CompareByType
        CompareByID
        CompareByDate
        CompareByComment
    End Enum
    Public Enum IDGenerationEnum
        GenerateID
        DoNotGenerateID
    End Enum

    Public Class EntryComparer
        Implements IComparer
        Public Sub New(ByVal cbeCompareBy As CompareByEnum)
            Me.cbeCompareBy = cbeCompareBy
        End Sub

        Private cbeCompareBy As CompareByEnum

        Public Property CompareBy() As CompareByEnum
            Get
                Return cbeCompareBy
            End Get
            Set(ByVal Value As CompareByEnum)
                cbeCompareBy = Value
            End Set
        End Property
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim jeX As FinancePlanner.JournalEntry, jeY As FinancePlanner.JournalEntry

            If Not (TypeOf (x) Is FinancePlanner.JournalEntry And TypeOf (y) Is FinancePlanner.JournalEntry) Then
                Throw New ArgumentException("Both objects being compared must be FinancePlanner.JournalEntry")
            Else

                jeX = CType(x, FinancePlanner.JournalEntry)
                jeY = CType(y, FinancePlanner.JournalEntry)

                Select Case cbeCompareBy
                    Case CompareByEnum.CompareByComment
                        If jeX.Comment > jeY.Comment Then
                            Return 1
                        ElseIf jeX.Comment < jeY.Comment Then
                            Return -1
                        Else
                            If jeX.ID > jeY.ID Then
                                Return 1
                            ElseIf jeX.ID < jeY.ID Then
                                Return -1
                            Else
                                Return 0
                            End If
                        End If
                    Case CompareByEnum.CompareByDate
                        If jeX._Date > jeY._Date Then
                            Return 1
                        ElseIf jeX._Date < jeY._Date Then
                            Return -1
                        Else
                            If jeX.ID > jeY.ID Then
                                Return 1
                            ElseIf jeX.ID < jeY.ID Then
                                Return -1
                            Else
                                Return 0
                            End If
                        End If
                    Case CompareByEnum.CompareByID
                        If jeX.ID > jeY.ID Then
                            Return 1
                        ElseIf jeX.ID < jeY.ID Then
                            Return -1
                        Else
                            Return 0
                        End If
                    Case CompareByEnum.CompareByType
                        If jeX.TypeAsString > jeY.TypeAsString Then
                            Return 1
                        ElseIf jeX.TypeAsString < jeY.TypeAsString Then
                            Return -1
                        Else
                            If jeX.ID > jeY.ID Then
                                Return 1
                            ElseIf jeX.ID < jeY.ID Then
                                Return -1
                            Else
                                Return 0
                            End If
                        End If
                End Select
            End If
        End Function
    End Class

    Public Sub New()
        jeEntries = New Collection()
    End Sub

    Private jeEntries As Collection

    Public ReadOnly Property Entries() As Collection
        Get
            Return jeEntries
        End Get
    End Property

    Public Sub Clear()
        jeEntries = New Collection()
    End Sub
    Public Sub AddEntry(ByVal jeEntry As FinancePlanner.JournalEntry, ByVal igeGenerateID As IDGenerationEnum)
        If igeGenerateID = IDGenerationEnum.GenerateID Then jeEntry.ID = jeEntries.Count

        jeEntries.Add(jeEntry, CStr(jeEntry.ID))
    End Sub
    Public Sub RemoveEntry(ByVal ID As Long)
        jeEntries.Remove(CStr(ID))
    End Sub

    Public Function EntriesAsArray(ByVal cbeSortBy As CompareByEnum) As FinancePlanner.JournalEntry()
        Dim jeEntry As FinancePlanner.JournalEntry
        Dim jeEntryArray() As FinancePlanner.JournalEntry

        ReDim jeEntryArray(-1)

        For Each jeEntry In jeEntries
            Utilities.AddItemToArray(jeEntry, jeEntryArray)
        Next

        Array.Sort(jeEntryArray, New EntryComparer(cbeSortBy))

        Return jeEntryArray
    End Function
End Class

Public Class JournalEntry
    Implements IComparable

    Public Sub New()

    End Sub


    Public Const BalanceTolerance = 0.015
    Private lngID As Integer
    Private dtmDate As Date
    Private eeType As EntryEnum
    Private strComment As String
    Private jtTransactions As Collection 'of Transactions
    Private dblNetBalance As Double

    Public Property TypeAsString()
        Get
            Select Case eeType
                Case EntryEnum.Opening
                    Return "Opening"
                Case EntryEnum.Closing
                    Return "Closing"
                Case EntryEnum.Adjusting
                    Return "Adjusting"
                Case EntryEnum.Transaction
                    Return "Transaction"
            End Select
        End Get
        Set(ByVal Value)
            Select Case Value.ToLower
                Case "opening", "open", "opn", "o"
                    eeType = EntryEnum.Opening
                Case "closing", "close", "clo", "c"
                    eeType = EntryEnum.Closing
                Case "adjusting", "adjust", "adj", "a"
                    eeType = EntryEnum.Adjusting
                Case Else
                    eeType = EntryEnum.Transaction
            End Select
        End Set
    End Property
    Public Property ID() As Long
        Get
            Return lngID
        End Get
        Set(ByVal Value As Long)
            lngID = Value
        End Set
    End Property
    Public Property _Date() As Date
        Get
            Return dtmDate
        End Get
        Set(ByVal Value As Date)
            dtmDate = Value
        End Set
    End Property
    Public Property Comment() As String
        Get
            Return strComment
        End Get
        Set(ByVal Value As String)
            strComment = Value
        End Set
    End Property
    Public Property Type() As EntryEnum
        Get
            Return eeType
        End Get
        Set(ByVal Value As EntryEnum)
            eeType = Value
        End Set
    End Property

    Public Function GetTransactions() As Collection
        Return jtTransactions
    End Function

    Public Function IsBalanced() As Boolean
        Return Math.Abs(dblNetBalance) < BalanceTolerance
    End Function
    Public Sub AddTransaction(ByVal jtTransaction As FinancePlanner.JournalTransaction)
        jtTransactions.Add(jtTransaction)
        If jtTransaction.NormalBalance = BalanceEnum.Debit Then
            dblNetBalance += jtTransaction.Balance
        Else
            dblNetBalance -= jtTransaction.Balance
        End If
    End Sub
    Public Sub RemoveTransaction(ByVal intIndex As Integer)
        If jtTransactions(intIndex).NormalBalance = BalanceEnum.Debit Then
            dblNetBalance -= jtTransactions(intIndex).Balance
        Else
            dblNetBalance += jtTransactions(intIndex).Balance
        End If
        jtTransactions.Remove(intIndex)
    End Sub

    Public Function CompareDates(ByVal obj As Object) As Integer Implements IComparable.CompareTo
        If TypeOf (obj) Is FinancePlanner.JournalEntry Then
            If _Date < CType(obj, FinancePlanner.JournalEntry)._Date Then
                Return -1
            ElseIf _Date > CType(obj, FinancePlanner.JournalEntry)._Date Then
                Return 1
            Else
                Return 0
            End If
        Else
            Throw New ArgumentException("Incorrect Argument Type Supplied")
        End If
    End Function
End Class


Public Class JournalTransaction
    Implements IComparable
    Public Path As String
    Public Balance As Double
    Public Quantity As Integer
    Public NormalBalance As BalanceEnum
    Public Function ComparePaths(ByVal obj As Object) As Integer Implements IComparable.CompareTo
        If TypeOf (obj) Is FinancePlanner.JournalTransaction Then
            If Path < CType(obj, FinancePlanner.JournalTransaction).Path Then
                Return -1
            ElseIf Path > CType(obj, FinancePlanner.JournalTransaction).Path Then
                Return 1
            Else
                Return 0
            End If
        Else
            Throw New ArgumentException("Incorrect Argument Type Supplied")
        End If
    End Function
    Public Property NormalBalanceAsString()
        Get
            Select Case NormalBalance
                Case BalanceEnum.Debit
                    Return "Debit"
                Case BalanceEnum.Credit
                    Return "Credit"
            End Select
        End Get
        Set(ByVal Value)
            Select Case Value.ToLower
                Case "debit", "deb", "de", "d"
                    NormalBalance = BalanceEnum.Debit
                Case "credit", "cred", "cre", "cr", "c"
                    NormalBalance = BalanceEnum.Credit
                Case Else
                    Throw New ArgumentException("Unknown Balance Type.  Must Be Debit or Credit.")
            End Select
        End Set
    End Property
End Class