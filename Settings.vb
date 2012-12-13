Public Class Settings

    'Constructors and Construction Helpers
    Public Sub New()
        dstSettings = New SettingsData()
        ReLoad()
    End Sub

    Public Sub New(ByVal strFileName As String)
        dstSettings = New SettingsData()
        Load(strFileName)
    End Sub

    Protected Sub Clear()
        dstSettings.Clear()
        drwSettings = dstSettings.Settings.AddSettingsRow()
    End Sub


    'Data Members
    Protected dstSettings As FinancePlanner.SettingsData
    Protected drwSettings As FinancePlanner.SettingsData.SettingsRow


    'Strings
    Protected strFileName As String = System.IO.Directory.GetCurrentDirectory & "\settings.xml"


    'Properties
    Public Property FileName() As String
        Get
            Return strFileName
        End Get
        Set(ByVal Value As String)
            strFileName = Value
        End Set
    End Property


    'Events
    Event Changed()



    'Loading and Saving
    Public Sub Load(ByVal strFileName As String)
        Try
            Me.strFileName = strFileName
            dstSettings.ReadXml(strFileName)
            drwSettings = dstSettings.Settings.Rows(0)
        Catch
            MessageBox.Show("Error Loading Settings File.  A default file has been created.", _
                            "Settings File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Clear()
            Save()
        End Try
    End Sub

    Public Sub ReLoad()
        Load(strFileName)
    End Sub

    Public Sub SaveAs(ByVal strFileName As String)
        Try
            Me.strFileName = strFileName
            dstSettings.WriteXml(strFileName)
        Catch
            MessageBox.Show("Error Saving Settings File.  Application settings may be lost.", _
                            "Settings File Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub Save()
        SaveAs(Me.strFileName)
    End Sub


    'Public Functions
    Public Function AddRecentDocument(ByVal strTitle As String, ByVal strFileName As String)
        Dim blnOpenedBefore As Boolean
        Dim drwDocument As FinancePlanner.SettingsData.RecentDocumentsRow

        blnOpenedBefore = False
        For Each drwDocument In drwSettings.GetRecentDocumentsRows()
            If drwDocument.path = strFileName Then
                blnOpenedBefore = True
                drwDocument.name = strTitle
            End If
        Next
        If Not blnOpenedBefore Then
            dstSettings.RecentDocuments.AddRecentDocumentsRow(strFileName, strTitle, drwSettings)
        End If
    End Function

    Function GetRecentDocuments() As FinancePlanner.DocumentNameAndPath()
        Dim docDocuments() As FinancePlanner.DocumentNameAndPath
        Dim drwDocument As FinancePlanner.SettingsData.RecentDocumentsRow
        Dim intUpperBound As Integer

        ReDim docDocuments(-1)

        For Each drwDocument In drwSettings.GetRecentDocumentsRows
            intUpperBound = docDocuments.GetUpperBound(0) + 1
            ReDim Preserve docDocuments(intUpperBound)
            docDocuments(intUpperBound) = New FinancePlanner.DocumentNameAndPath()
            docDocuments(intUpperBound).Name = drwDocument.name
            docDocuments(intUpperBound).Path = drwDocument.path
        Next

        Return docDocuments
    End Function

    Public Function GetRecentDocumentPath(ByVal intIndex As Integer) As String
        Dim strPath As String

        Try
            strPath = drwSettings.GetRecentDocumentsRows()(intIndex).path
        Catch
            strPath = String.Empty
        End Try

        Return strPath
    End Function
End Class
