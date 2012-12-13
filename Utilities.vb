Module Utilities

    Public Function StartOfThisYear() As Date
        Return CType("1/1/" & Date.Now.Year.ToString, DateTime)
    End Function

    Public Function StartOfLastYear()
        Return CType("1/1/" & (Date.Now.Year - 1).ToString, DateTime)
    End Function

    Public Sub AddItemToArray(ByRef objItem As Object, ByRef objArray() As Object, Optional ByVal blnAddIfNothing As Boolean = False)
        Dim intArraySize As Integer
        Dim intItemSize As Integer
        Dim intCounter As Integer

        If objArray Is Nothing Then ReDim objArray(-1)

        If TypeOf objItem Is Array Then
            intItemSize = CType(objItem, Array).Length
        Else
            If objItem Is Nothing Then
                If blnAddIfNothing Then
                    intItemSize = 1
                Else
                    intItemSize = 0
                End If
            Else
                intItemSize = 1
            End If
        End If

        intArraySize = objArray.GetUpperBound(0)
        ReDim Preserve objArray(intArraySize + intItemSize)
        If intItemSize > 1 Then
            intCounter = 1
            Do While intCounter <= intItemSize
                objArray(intArraySize + intCounter) = CType(objItem, Array)(intCounter)
                intCounter += 1
            Loop
        Else
            objArray(intArraySize + intItemSize) = objItem
        End If
    End Sub

    Public Sub DoTest(ByRef prjProject As FinancePlanner.Project)
        Dim frmTest As FrmTest

        If Not prjProject Is Nothing Then
            frmTest = New FrmTest(prjProject)
            frmTest.Show()
        End If
    End Sub

    Public Function MakePlural(ByVal strNoun As String) As String
        If Not strNoun.EndsWith("s") Then
            If strNoun.EndsWith("y") Then
                Select Case strNoun.Chars(strNoun.Length - 1)
                    Case "a", "e", "i", "o", "u"
                        Return strNoun & "s"
                    Case Else
                        Return strNoun.Substring(0, strNoun.Length - 1) & "ies"
                End Select
            Else
                Return strNoun & "s"
            End If
        ElseIf strNoun.EndsWith("ss") Then
            Return strNoun & "es"
        Else
            Return strNoun
        End If
    End Function

    Public Function MakeSingular(ByVal strNoun As String) As String
        If strNoun.EndsWith("ies") Then
            Return strNoun.Substring(0, strNoun.Length - 3) & "y"
        ElseIf strNoun.EndsWith("s") And Not strNoun.EndsWith("ss") Then
            If strNoun.Chars(strNoun.Length - 2) = "s" Then
                Return strNoun.Substring(0, strNoun.Length - 2)
            Else
                Return strNoun.Substring(0, strNoun.Length - 1)
            End If
        Else
            Return strNoun
        End If
    End Function

End Module

