Public Module AccountPath
    Public Const SubaccountDelimiter As Char = "-"c

    'Functions
    Public Function Depth(ByVal strPath As String) As Integer
        If strPath = String.Empty Then
            Return 0
        Else
            Return strPath.Split(SubaccountDelimiter).Length
        End If
    End Function

    Public Function IsAliasFor(ByVal AliasPath As String, ByVal FullPath As String) As Boolean
        Dim intCounter As Integer, intAliasDelimiters As Integer = 0
        Dim intAccountPathDelimiters As Integer = 0
        Dim blnResult As Boolean = False
        Dim intDelimiterIndex As Integer

        For intCounter = 0 To AliasPath.Length - 1
            If AliasPath.Chars(intCounter) = SubaccountDelimiter Then
                intAliasDelimiters += 1
            End If
        Next

        For intCounter = 0 To FullPath.Length - 1
            If FullPath.Chars(intCounter) = SubaccountDelimiter Then
                intAccountPathDelimiters += 1
            End If
        Next

        For intCounter = 1 To intAccountPathDelimiters - intAliasDelimiters
            With FullPath
                intDelimiterIndex = .IndexOf(SubaccountDelimiter)
                FullPath = .Substring(intDelimiterIndex + 1, _
                                 .Length - (intDelimiterIndex + 1))
            End With
        Next

        If LCase(AliasPath) = LCase(FullPath) Then
            blnResult = True
        Else
            blnResult = False
        End If

        Return blnResult
    End Function

    Public Function IsSubAccountPathOf(ByVal SubAccountPath As String, ByVal ParentPath As String, _
                                Optional ByVal blnIncludeSamePath As Boolean = True) As Boolean
        Dim intDelimiterIndex As Integer

        If SubAccountPath.Length < ParentPath.Length Then
            Return False
        ElseIf SubAccountPath.Length > ParentPath.Length Then
            intDelimiterIndex = _
                            SubAccountPath.IndexOf(SubaccountDelimiter, _
                                                    ParentPath.Length)
            If intDelimiterIndex >= 0 Then
                Return IsSamePath(SubAccountPath.Substring(0, intDelimiterIndex), _
                                                ParentPath)
            Else
                Return False
            End If
        Else
            If blnIncludeSamePath Then
                Return IsSamePath(SubAccountPath, ParentPath)
            Else
                Return False
            End If
        End If
    End Function

    Public Function IsSamePath(ByVal Path1 As String, ByVal Path2 As String) As Boolean
        Return Path1.ToLower = Path2.ToLower
    End Function

    Public Function AccountName(ByVal FullPath As String) As String
        Dim intDelimiterLocation As Integer

        intDelimiterLocation = FullPath.LastIndexOf(SubaccountDelimiter)

        If intDelimiterLocation < FullPath.Length - 1 And intDelimiterLocation >= 0 Then
            Return FullPath.Substring(intDelimiterLocation + 1)
        Else
            Return String.Empty
        End If
    End Function

    Public Function ParentPath(ByVal FullPath As String) As String
        Dim intDelimiterLocation As Integer

        intDelimiterLocation = FullPath.LastIndexOf(SubaccountDelimiter)

        If intDelimiterLocation < FullPath.Length - 1 And intDelimiterLocation >= 0 Then
            Return FullPath.Substring(0, intDelimiterLocation)
        Else
            Return String.Empty
        End If
    End Function

    Public Function ChangeAbbreviatedPath(ByVal AbbreviatedPath As String, ByVal OriginalFullPath As String, ByVal NewFullPath As String) As String
        Dim strOriginalFragment() As String, strNewFragment() As String
        Dim intCounter As Integer

        strOriginalFragment = OriginalFullPath.Split(SubaccountDelimiter)
        strNewFragment = NewFullPath.Split(SubaccountDelimiter)

        If strNewFragment.Length <> strOriginalFragment.Length Then
            Throw New System.Exception("Original and New paths must have the same depth.")
        End If

        For intCounter = 0 To strOriginalFragment.Length - 1
            If strNewFragment(intCounter).ToLower() <> strOriginalFragment(intCounter).ToLower() Then
                AbbreviatedPath = Replace(AbbreviatedPath, strOriginalFragment(intCounter), strNewFragment(intCounter))
            End If
        Next

        Return AbbreviatedPath
    End Function

    Public Function IsSubaccountPathOfAlias(ByVal SubAccountPath As String, ByVal AliasPath As String, _
                                     Optional ByVal blnIncludeSamePath As Boolean = True) As Boolean
        Dim strSubaccountFragments() As String, strAliasFragments() As String
        Dim intCounter1 As Integer, intCounter2 As Integer
        Dim blnMatchFound As Boolean

        strSubaccountFragments = SubAccountPath.Split(SubaccountDelimiter)
        strAliasFragments = AliasPath.Split(SubaccountDelimiter)

        blnMatchFound = False
        For intCounter1 = strSubaccountFragments.Length - 1 To 0 Step -1
            If strSubaccountFragments(intCounter1).ToLower() = strAliasFragments(0).ToLower() Then
                blnMatchFound = True
                For intCounter2 = 1 To strAliasFragments.Length - 1
                    If intCounter1 + intCounter2 <= strSubaccountFragments.Length - 1 Then
                        If strAliasFragments(intCounter2).ToLower() <> _
                            strSubaccountFragments(intCounter1 + intCounter2).ToLower() Then
                            blnMatchFound = False
                        Else
                            If intCounter1 + intCounter2 = strSubaccountFragments.Length - 1 And _
                               blnIncludeSamePath = False Then
                                blnMatchFound = False
                            End If
                        End If
                    End If
                Next
            End If
        Next


        Return blnMatchFound
    End Function

    Public Function Replace(ByVal FullPath As String, ByVal OldValue As String, ByVal NewValue As String) As String
        Dim strPathFragments() As String
        Dim intCounter As Integer

        strPathFragments = FullPath.Split(SubaccountDelimiter)

        For intCounter = 0 To strPathFragments.Length - 1
            If strPathFragments(intCounter).ToLower() = OldValue.ToLower() Then
                strPathFragments(intCounter) = NewValue
            End If
        Next

        Return String.Join(SubaccountDelimiter, strPathFragments)
    End Function

    Public Function PathsMatch(ByVal strPrimaryPath As String, ByVal strSecondaryPath As String, ByVal amMatchType As FinancePlanner.AccountMatchEnum) As Boolean
        Select Case amMatchType
            Case AccountMatchEnum.MatchAlias
                Return IsAliasFor(strSecondaryPath, strPrimaryPath)
            Case AccountMatchEnum.MatchFullPath
                Return IsSamePath(strSecondaryPath, strPrimaryPath)
        End Select
    End Function

End Module