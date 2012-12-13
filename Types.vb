'Types

Public Enum AccountMatchEnum
    MatchAlias
    MatchFullPath
End Enum
Public Enum ProjectWindowEnum
    'ID for a base class instance (for the visual designer)
    ProjectWindowBase = 0

    'Journal
    GeneralJournal = 1
    CashJournal = 2
    ReceivablesJournal = 3
    PayablesJournal = 4

    'Ledger
    GeneralLedger = 21
    DetailedLedger = 22
    AccountDetail = 23

    'Budget
    GeneralBudget = 41

    'Reports
    IncomeStatement = 61
    BalanceSheet = 62
    CashFlowsStatement = 63

    'Misc
    ProjectProperties = 81
    AccountConflict = 82
    TestWindow = 83
End Enum
Public Enum ProjectControlEnum
    'ID for a base class instance (for visual designer purposes)
    ProjectControlBase = 0

    'Controls associated w/ the Journal


    'Controls Associated w/ the Ledger
    AccountTree = 21
    TransactionTree = 22
    AccountDetailChart = 23

    'Controls Associated w/ the Budget


    'Controls Associated w/ Reports


    'Misc Controls

End Enum
Public Enum BalanceEnum
    Debit
    Credit
End Enum
Public Enum EntryEnum
    Opening
    Closing
    Adjusting
    Transaction
End Enum
Public Enum BudgetAccountEnum
    Fixed
    Variable
End Enum

Public Structure DocumentNameAndPath
    Public Name As String
    Public Path As String
End Structure
'Public Structure WindowLayout
'    Public Height As Integer
'    Public Width As Integer
'    Public Top As Integer
'    Public Left As Integer
'End Structure
