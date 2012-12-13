Interface IParses
    Inherits IProjectWindow

    Sub Parse()
    Event Changed()
End Interface

Interface IProjectWindow
    Sub Render()
    'ReadOnly Property ParseType() As FinancePlanner.ParseTypeEnum
    ReadOnly Property Type() As FinancePlanner.ProjectWindowEnum
    ReadOnly Property Title() As String
End Interface

Interface IProjectControl
    Sub Render()
    ReadOnly Property Type() As FinancePlanner.ProjectControlEnum
End Interface




