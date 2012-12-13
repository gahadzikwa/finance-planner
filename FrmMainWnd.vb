Public Class FrmMainWnd
    Inherits System.Windows.Forms.Form

    Private strCloseBaseText As String = "&Close", _
            strSaveBaseText As String = "&Save", _
            strTitleBaseText As String = "Finance Planner", _
            strSaveAsBaseText As String = "Save", _
            strSaveAsEndText As String = "&As..."

    'Data Members
    Protected WithEvents prjProject As FinancePlanner.Project
    Protected WithEvents stgSettings As FinancePlanner.Settings

    'Flags
    Protected blnProjectChanged As Boolean = False


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
    Friend WithEvents MainMenu As System.Windows.Forms.MainMenu
    Friend WithEvents itmFile As System.Windows.Forms.MenuItem
    Friend WithEvents itmNew As System.Windows.Forms.MenuItem
    Friend WithEvents itmOpen As System.Windows.Forms.MenuItem
    Friend WithEvents itmClose As System.Windows.Forms.MenuItem
    Friend WithEvents itmSave As System.Windows.Forms.MenuItem
    Friend WithEvents itmSaveAs As System.Windows.Forms.MenuItem
    Friend WithEvents itmExit As System.Windows.Forms.MenuItem
    Friend WithEvents itmSeparator2 As System.Windows.Forms.MenuItem
    Friend WithEvents dstSettings As FinancePlanner.SettingsData
    Friend WithEvents itmSeparator3 As System.Windows.Forms.MenuItem
    Friend WithEvents itmRecent As System.Windows.Forms.MenuItem
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog As System.Windows.Forms.SaveFileDialog
    Friend WithEvents tbMainButtons As System.Windows.Forms.ToolBar
    Friend WithEvents btnSep1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents ilButtonImages As System.Windows.Forms.ImageList
    Friend WithEvents btnPost As System.Windows.Forms.ToolBarButton
    Friend WithEvents itmProject As System.Windows.Forms.MenuItem
    Friend WithEvents itmViewJournal As System.Windows.Forms.MenuItem
    Friend WithEvents itmViewLedger As System.Windows.Forms.MenuItem
    Friend WithEvents itmViewBudget As System.Windows.Forms.MenuItem
    Friend WithEvents itmPost As System.Windows.Forms.MenuItem
    Friend WithEvents itmProperties As System.Windows.Forms.MenuItem
    Friend WithEvents itmSeparator1 As System.Windows.Forms.MenuItem
    Friend WithEvents itmSeparator4 As System.Windows.Forms.MenuItem
    Friend WithEvents itmSeparator5 As System.Windows.Forms.MenuItem
    Friend WithEvents itmTest As System.Windows.Forms.MenuItem
    Friend WithEvents btnGeneralJournal As System.Windows.Forms.ToolBarButton
    Friend WithEvents btnGeneralLedger As System.Windows.Forms.ToolBarButton
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmMainWnd))
        Me.MainMenu = New System.Windows.Forms.MainMenu()
        Me.itmFile = New System.Windows.Forms.MenuItem()
        Me.itmNew = New System.Windows.Forms.MenuItem()
        Me.itmSeparator1 = New System.Windows.Forms.MenuItem()
        Me.itmOpen = New System.Windows.Forms.MenuItem()
        Me.itmSave = New System.Windows.Forms.MenuItem()
        Me.itmSaveAs = New System.Windows.Forms.MenuItem()
        Me.itmClose = New System.Windows.Forms.MenuItem()
        Me.itmSeparator2 = New System.Windows.Forms.MenuItem()
        Me.itmRecent = New System.Windows.Forms.MenuItem()
        Me.itmSeparator3 = New System.Windows.Forms.MenuItem()
        Me.itmExit = New System.Windows.Forms.MenuItem()
        Me.itmProject = New System.Windows.Forms.MenuItem()
        Me.itmViewJournal = New System.Windows.Forms.MenuItem()
        Me.itmViewLedger = New System.Windows.Forms.MenuItem()
        Me.itmViewBudget = New System.Windows.Forms.MenuItem()
        Me.itmSeparator4 = New System.Windows.Forms.MenuItem()
        Me.itmPost = New System.Windows.Forms.MenuItem()
        Me.itmSeparator5 = New System.Windows.Forms.MenuItem()
        Me.itmProperties = New System.Windows.Forms.MenuItem()
        Me.itmTest = New System.Windows.Forms.MenuItem()
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.dstSettings = New FinancePlanner.SettingsData()
        Me.SaveFileDialog = New System.Windows.Forms.SaveFileDialog()
        Me.tbMainButtons = New System.Windows.Forms.ToolBar()
        Me.btnGeneralJournal = New System.Windows.Forms.ToolBarButton()
        Me.btnGeneralLedger = New System.Windows.Forms.ToolBarButton()
        Me.btnSep1 = New System.Windows.Forms.ToolBarButton()
        Me.btnPost = New System.Windows.Forms.ToolBarButton()
        Me.ilButtonImages = New System.Windows.Forms.ImageList(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        CType(Me.dstSettings, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu
        '
        Me.MainMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.itmFile, Me.itmProject, Me.MenuItem1})
        '
        'itmFile
        '
        Me.itmFile.Index = 0
        Me.itmFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.itmNew, Me.itmSeparator1, Me.itmOpen, Me.itmSave, Me.itmSaveAs, Me.itmClose, Me.itmSeparator2, Me.itmRecent, Me.itmSeparator3, Me.itmExit})
        Me.itmFile.Text = "&File"
        '
        'itmNew
        '
        Me.itmNew.Index = 0
        Me.itmNew.Text = "&New Project..."
        '
        'itmSeparator1
        '
        Me.itmSeparator1.Index = 1
        Me.itmSeparator1.Text = "-"
        '
        'itmOpen
        '
        Me.itmOpen.Index = 2
        Me.itmOpen.Text = "&Open Project..."
        '
        'itmSave
        '
        Me.itmSave.Enabled = False
        Me.itmSave.Index = 3
        Me.itmSave.Text = "&Save"
        '
        'itmSaveAs
        '
        Me.itmSaveAs.Enabled = False
        Me.itmSaveAs.Index = 4
        Me.itmSaveAs.Text = "Save &As..."
        '
        'itmClose
        '
        Me.itmClose.Enabled = False
        Me.itmClose.Index = 5
        Me.itmClose.Text = "&Close"
        '
        'itmSeparator2
        '
        Me.itmSeparator2.Index = 6
        Me.itmSeparator2.Text = "-"
        '
        'itmRecent
        '
        Me.itmRecent.Index = 7
        Me.itmRecent.Text = "&Recent Files"
        '
        'itmSeparator3
        '
        Me.itmSeparator3.Index = 8
        Me.itmSeparator3.Text = "-"
        '
        'itmExit
        '
        Me.itmExit.Index = 9
        Me.itmExit.Text = "&Exit"
        '
        'itmProject
        '
        Me.itmProject.Enabled = False
        Me.itmProject.Index = 1
        Me.itmProject.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.itmViewJournal, Me.itmViewLedger, Me.itmViewBudget, Me.itmSeparator4, Me.itmPost, Me.itmTest, Me.itmSeparator5, Me.itmProperties})
        Me.itmProject.Text = "&Project"
        '
        'itmViewJournal
        '
        Me.itmViewJournal.Index = 0
        Me.itmViewJournal.Text = "View &Journal"
        '
        'itmViewLedger
        '
        Me.itmViewLedger.Index = 1
        Me.itmViewLedger.Text = "View &Ledger"
        '
        'itmViewBudget
        '
        Me.itmViewBudget.Enabled = False
        Me.itmViewBudget.Index = 2
        Me.itmViewBudget.Text = "View &Budget"
        '
        'itmSeparator4
        '
        Me.itmSeparator4.Index = 3
        Me.itmSeparator4.Text = "-"
        '
        'itmPost
        '
        Me.itmPost.Index = 4
        Me.itmPost.Text = "&Post"
        '
        'itmSeparator5
        '
        Me.itmSeparator5.Index = 6
        Me.itmSeparator5.Text = "-"
        '
        'itmProperties
        '
        Me.itmProperties.Index = 7
        Me.itmProperties.Text = "P&roperties"
        '
        'itmTest
        '
        Me.itmTest.Index = 5
        Me.itmTest.Text = "Test"
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.Filter = "Finance Planner Project|*.fpp|Raw XML|*.xml|All Files|*.*"
        Me.OpenFileDialog.Title = "Open File"
        '
        'dstSettings
        '
        Me.dstSettings.DataSetName = "Settings"
        Me.dstSettings.Locale = New System.Globalization.CultureInfo("en-US")
        Me.dstSettings.Namespace = "http://tempuri.org/Settings.xsd"
        '
        'SaveFileDialog
        '
        Me.SaveFileDialog.DefaultExt = "xml"
        Me.SaveFileDialog.FileName = "document1"
        Me.SaveFileDialog.Filter = "Finance Planner Project|*.fpp|Raw XML|*.xml|All Files|*.*"
        Me.SaveFileDialog.Title = "Save Journal"
        '
        'tbMainButtons
        '
        Me.tbMainButtons.Appearance = System.Windows.Forms.ToolBarAppearance.Flat
        Me.tbMainButtons.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.btnGeneralJournal, Me.btnGeneralLedger, Me.btnSep1, Me.btnPost})
        Me.tbMainButtons.Divider = False
        Me.tbMainButtons.DropDownArrows = True
        Me.tbMainButtons.ImageList = Me.ilButtonImages
        Me.tbMainButtons.Name = "tbMainButtons"
        Me.tbMainButtons.ShowToolTips = True
        Me.tbMainButtons.Size = New System.Drawing.Size(594, 37)
        Me.tbMainButtons.TabIndex = 1
        '
        'btnGeneralJournal
        '
        Me.btnGeneralJournal.Enabled = False
        Me.btnGeneralJournal.ImageIndex = 0
        Me.btnGeneralJournal.Tag = "1"
        Me.btnGeneralJournal.Text = "Journal"
        Me.btnGeneralJournal.ToolTipText = "Show the General Journal View"
        '
        'btnGeneralLedger
        '
        Me.btnGeneralLedger.Enabled = False
        Me.btnGeneralLedger.ImageIndex = 1
        Me.btnGeneralLedger.Tag = "21"
        Me.btnGeneralLedger.Text = "Ledger"
        Me.btnGeneralLedger.ToolTipText = "Show the General Ledger View"
        '
        'btnSep1
        '
        Me.btnSep1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'btnPost
        '
        Me.btnPost.Enabled = False
        Me.btnPost.ImageIndex = 2
        Me.btnPost.Tag = "-1"
        Me.btnPost.Text = "Post"
        Me.btnPost.ToolTipText = "Post To Ledger"
        '
        'ilButtonImages
        '
        Me.ilButtonImages.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ilButtonImages.ImageSize = New System.Drawing.Size(16, 16)
        Me.ilButtonImages.ImageStream = CType(resources.GetObject("ilButtonImages.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ilButtonImages.TransparentColor = System.Drawing.Color.Transparent
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 2
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem3})
        Me.MenuItem1.Text = "Window"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "Tile"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "-"
        '
        'FrmMainWnd
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(594, 40)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.tbMainButtons})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu
        Me.Name = "FrmMainWnd"
        Me.Text = "Finance Planner"
        CType(Me.dstSettings, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


    'Event Handlers

    'Opening
    Private Sub FrmMainWnd_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetToDefaultLocation()

        Dim frmSplash As FrmSplash
        frmSplash = New FrmSplash()
        frmSplash.StartPosition = FormStartPosition.CenterScreen
        'frmSplash.MdiParent = Me
        frmSplash.Show()

        stgSettings = New Settings()
        AddRecentDocumentsToMenuItem(itmRecent)
    End Sub

    Sub SetToDefaultLocation()
        Me.Left = Screen.GetWorkingArea(Me).Left
        Me.Top = Screen.GetWorkingArea(Me).Top
        Me.Width = Screen.GetWorkingArea(Me).Width
    End Sub


    'Closing and Exiting
    Private Sub FrmMainWnd_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If ConfirmCloseAndSaveProject() Then
            stgSettings.Save()
        Else
            e.Cancel = True
        End If
    End Sub
    Private Sub itmExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles itmExit.Click
        Me.Close()
    End Sub
    Private Sub itmClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles itmClose.Click
        ConfirmCloseAndSaveProject()
    End Sub


    'Project Events
    Protected Sub prjProject_Changed() Handles prjProject.Changed
        itmSave.Enabled = True
        blnProjectChanged = True
        UpdateMainWindowText()
    End Sub
    Protected Sub prjProject_Unchanged() Handles prjProject.FileSaved, prjProject.FileLoaded, prjProject.Cleared
        blnProjectChanged = False
        itmSave.Enabled = False
        UpdateMainWindowText()
    End Sub
    Protected Sub prjProject_Closed() Handles prjProject.Closed
        itmClose.Enabled = False
        itmSaveAs.Enabled = False
        itmSave.Enabled = False
        itmProject.Enabled = False
        DisableButtons()
        UpdateMainWindowText()
    End Sub
    Protected Sub prjProject_Load() Handles prjProject.Load
        itmClose.Enabled = True
        itmSave.Enabled = True
        itmSaveAs.Enabled = True
        itmProject.Enabled = True
        EnableButtons()
        UpdateMainWindowText()
    End Sub

    Private Sub EnableButtons()
        Dim btnButton As ToolBarButton

        For Each btnButton In tbMainButtons.Buttons
            btnButton.Enabled = True
        Next
    End Sub
    Private Sub DisableButtons()
        Dim btnButton As ToolBarButton

        For Each btnButton In tbMainButtons.Buttons
            btnButton.Enabled = False
        Next
    End Sub

    'Saving and Loading Dialog Boxes
    Private Sub MainOpenFileDialog_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog.FileOk
        LoadFile(OpenFileDialog.FileName)
    End Sub
    Protected Sub JournalSaveFileDialog_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog.FileOk
        If Not (prjProject Is Nothing) Then
            prjProject.SaveFileAs(SaveFileDialog.FileName)
            stgSettings.AddRecentDocument(prjProject.Title, prjProject.FileName)
        End If
    End Sub

    Public Sub LoadFile(ByVal strFileName As String)
        Dim strExtension As String, intLastDotIndex As Integer

        intLastDotIndex = strFileName.LastIndexOf(".")
        If intLastDotIndex <> -1 Then
            strExtension = strFileName.Substring(intLastDotIndex)
        End If
        Select Case strExtension
            Case ".fpp", ".xml"
                OpenProject(strFileName)
            Case Else
                MessageBox.Show("Invalid file type selected.  Please select a valid file type.", _
                    "Error Opening File", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Select
    End Sub
    Public Sub SaveFile()
        If Not (prjProject Is Nothing) Then
            If prjProject.FileName <> String.Empty Then
                prjProject.SaveFile()
                stgSettings.AddRecentDocument(prjProject.Title, prjProject.FileName)
            Else
                ShowSaveDialog()
            End If
        End If
    End Sub


    'Menu Items
    Private Sub itmOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmOpen.Click
        OpenFileDialog.ShowDialog()
    End Sub
    Private Sub itmNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmNew.Click
        OpenProject(String.Empty)
    End Sub
    Private Sub itmSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles itmSave.Click
        SaveFile()
    End Sub
    Private Sub itmSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmSaveAs.Click
        ShowSaveDialog()
    End Sub
    Private Sub itmProperties_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles itmProperties.Click
        prjProject.NewProjectWindow(ProjectWindowEnum.ProjectProperties)
    End Sub
    Private Sub itmPost_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles itmPost.Click
        prjProject.Post()
    End Sub
    Public Sub LoadRecentDocument(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim intItemIndex As Integer
        Dim strPath As String

        intItemIndex = CType(sender, MenuItem).Index
        strPath = stgSettings.GetRecentDocumentPath(intItemIndex)

        LoadFile(strPath)
    End Sub
    Public Function AddRecentDocumentsToMenuItem(ByRef itmRecent As MenuItem)
        Dim docDocument As FinancePlanner.DocumentNameAndPath
        Dim intCounter As Integer = 0
        Dim itmRecentDocument As MenuItem

        For Each docDocument In stgSettings.GetRecentDocuments()
            itmRecentDocument = New MenuItem(docDocument.Name)
            itmRecentDocument.Index = intCounter
            itmRecent.MenuItems.Add(itmRecentDocument)
            AddHandler itmRecentDocument.Click, New System.EventHandler(AddressOf Me.LoadRecentDocument)
            intCounter += 1
        Next
    End Function

    'Toolbar Buttons
    Protected Sub tbMainButtons_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles tbMainButtons.ButtonClick
        Dim enmButtonType As ProjectWindowEnum
        enmButtonType = CType(e.Button.Tag, Integer)

        '-1 means this is an action button, not a form button
        If enmButtonType > 0 Then
            prjProject.NewProjectWindow(enmButtonType)
        Else
            If e.Button Is Me.btnPost Then
                prjProject.Post()
            End If
        End If
    End Sub

    'Utility Functions
    Protected Sub OpenProject(ByVal strFileName As String)
        Dim blnProjectClosed As Boolean

        If Not prjProject Is Nothing Then
            blnProjectClosed = ConfirmCloseAndSaveProject()
        Else
            blnProjectClosed = True
        End If

        If blnProjectClosed Then
            If strFileName <> String.Empty Then
                prjProject = New Project()
                prjProject.SetMainWindow(Me)
                If prjProject.LoadFile(strFileName) Then
                    stgSettings.AddRecentDocument(prjProject.Title, prjProject.FileName)
                Else
                    prjProject.Close()
                    prjProject = Nothing
                End If
            Else
                prjProject = New Project()
                prjProject.Clear()
                prjProject.SetMainWindow(Me)
                prjProject.NewProjectWindow(ProjectWindowEnum.GeneralJournal)
                prjProject.NewProjectWindow(ProjectWindowEnum.ProjectProperties)
            End If
        End If
    End Sub
    Protected Sub CloseProject()
        If Not prjProject Is Nothing Then
            prjProject.Close()
            prjProject = Nothing
        End If
    End Sub
    Protected Function ConfirmCloseAndSaveProject() As Boolean
        If blnProjectChanged Then
            Select Case MessageBox.Show("Would you like to save " & prjProject.Title & " before closing it?", _
                                        "File Not Saved", _
                                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                Case DialogResult.Yes
                    prjProject.SaveFile()
                    CloseProject()
                    Return True
                Case DialogResult.No
                    CloseProject()
                    Return True
                Case DialogResult.Cancel
                    Return False
            End Select
        Else
            CloseProject()
            Return True
        End If
    End Function
    Protected Sub ShowSaveDialog()
        If Not (prjProject Is Nothing) Then
            If prjProject.FileName <> String.Empty Then
                Me.SaveFileDialog.FileName = prjProject.FileName
            End If
        End If
        Me.SaveFileDialog.ShowDialog()
    End Sub
    Protected Sub UpdateMainWindowText()
        If Not prjProject Is Nothing Then
            If blnProjectChanged Then
                Me.Text = prjProject.Title & "* - " & strTitleBaseText
            Else
                Me.Text = prjProject.Title & " - " & strTitleBaseText
            End If
            itmSaveAs.Text = strSaveAsBaseText & " " & prjProject.Title & _
                                       " " & strSaveAsEndText
            itmSave.Text = strSaveBaseText & " " & prjProject.Title
            itmClose.Text = strCloseBaseText & " " & prjProject.Title
        Else
            Me.Text = strTitleBaseText
            itmSaveAs.Text = strSaveAsBaseText & " " & strSaveAsEndText
            itmSave.Text = strSaveBaseText
            itmClose.Text = strCloseBaseText
        End If
    End Sub

    Private Sub FrmMainWnd_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

    End Sub

    Private Sub itmTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles itmTest.Click
        Utilities.DoTest(prjProject)
    End Sub
End Class
