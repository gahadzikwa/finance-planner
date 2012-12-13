Public Class FrmAccountDetail
    Inherits FinancePlanner.FrmProjectWindowBase

#Region " Windows Form Designer generated code "

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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        '
        'AccountDetail
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 274)
        Me.Name = "AccountDetail"
        Me.Text = "AccountDetail"

    End Sub

#End Region

    Public Sub New(ByRef prjProject As FinancePlanner.Project)
        MyBase.New(prjProject)

        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub

    Public Overrides ReadOnly Property Type() As FinancePlanner.ProjectWindowEnum
        Get
            Return ProjectWindowEnum.AccountDetail
        End Get
    End Property

    Public Overrides ReadOnly Property Title() As String
        Get
            Return "Account Details"
        End Get
    End Property

    Public Overrides Sub Render()
        RenderChart()
    End Sub

    Public Sub RenderChart()
        MsgBox("Rendering Chart")
        'TODO: Revise Chart Render
        '    Dim drwAccount As FinancePlanner.ProjectData.LedgerAccountRow
        '    Dim drwSelectedAccount As FinancePlanner.ProjectData.LedgerAccountRow
        '    Dim strChartData(,) As String
        '    Dim dblBinValues() As Double
        '    Dim intValue As Integer
        '    Dim tndNode As TreeNode
        '    Dim intDelimiterIndex As Integer
        '    Dim strTitle As String
        '    Dim dblBalance As Double
        '    Dim intBins As Integer
        '    Dim dtmBinValue As DateTime
        '    Dim drwEntry As FinancePlanner.ProjectData.EntryRow
        '    Dim drwTransaction As FinancePlanner.ProjectData.TransactionRow
        '    Dim intDateValue As Integer
        '    Dim intBinNumber As Integer
        '    Dim intCounter As Integer

        '    If Not Me.tvwDetailedAccounts.SelectedNode Is Nothing Then
        '        drwSelectedAccount = Me.tvwDetailedAccounts.SelectedNode.Tag

        '        If Me.radBalanceHistory.Checked Or _
        '           Me.radChangeHistory.Checked Then
        '            chDetailChart.chartType = MSChart20Lib.VtChChartType.VtChChartType2dLine
        '            chDetailChart.Plot.DataSeriesInRow = True
        '            chDetailChart.ShowLegend = False
        '            strTitle = drwSelectedAccount.Path

        '            intBins = Math.Floor(chDetailChart.Width / 5)  'width in pixels of a bin
        '            ReDim dblBinValues(intBins - 1)

        '            For Each drwEntry In prjProject.Journal().GetEntryRows()
        '                For Each drwTransaction In drwEntry.GetTransactionRows()
        '                    If IsSubaccountPathOfAlias(drwSelectedAccount.Path, drwTransaction.Path, True) Then
        '                        Try
        '                            intBinNumber = Math.Round( _
        '                                (drwEntry._Date.Ticks - dtmStart.Value.Ticks) / _
        '                                (dtmEnd.Value.Ticks - dtmStart.Value.Ticks) * intBins)
        '                        Catch
        '                            MessageBox.Show("The Start and End Dates cannot be the same." & _
        '                                            "Resetting to default values.", _
        '                                            "Chart Error", MessageBoxButtons.OK, _
        '                                            MessageBoxIcon.Error)

        '                            If Date.Now.Month = 1 & _
        '                                Date.Now.Day = 1 Then
        '                                dtmStart.Value = StartOfLastYear()
        '                            Else
        '                                dtmStart.Value = StartOfThisYear()
        '                            End If
        '                            dtmEnd.Value = DateTime.Now

        '                            intBinNumber = Math.Floor( _
        '                                (drwEntry._Date.Ticks - dtmStart.Value.Ticks) / _
        '                                (dtmEnd.Value.Ticks - dtmStart.Value.Ticks) * intBins)
        '                        End Try

        '                        If Me.radBalanceHistory.Checked Then
        '                            If intBinNumber < 0 Then intBinNumber = 0

        '                            If intBinNumber <= intBins - 1 Then
        '                                For intCounter = intBinNumber To intBins - 1
        '                                    If Not drwTransaction.IsDebitValueNull() Then
        '                                        dblBinValues(intCounter) += _
        '                                            drwTransaction.DebitValue
        '                                    End If
        '                                    If Not drwTransaction.IsCreditValueNull() Then
        '                                        dblBinValues(intCounter) -= _
        '                                            drwTransaction.CreditValue
        '                                    End If
        '                                Next
        '                            End If
        '                        Else
        '                            If intBinNumber >= 0 And intBinNumber <= intBins - 1 Then
        '                                If Not drwTransaction.IsDebitValueNull() Then
        '                                    dblBinValues(intBinNumber) += _
        '                                        drwTransaction.DebitValue
        '                                End If
        '                                If Not drwTransaction.IsCreditValueNull() Then
        '                                    dblBinValues(intBinNumber) -= _
        '                                        drwTransaction.CreditValue
        '                                End If
        '                            End If
        '                        End If
        '                    End If
        '                Next
        '            Next



        '            For intBinNumber = _
        '                dblBinValues.GetLowerBound(0) To dblBinValues.GetUpperBound(0)
        '                If strChartData Is Nothing Then
        '                    ReDim strChartData(1, 0)
        '                Else
        '                    ReDim Preserve strChartData(1, strChartData.GetUpperBound(1) + 1)
        '                End If

        '                dtmBinValue = New DateTime((intBinNumber / intBins * (dtmEnd.Value.Ticks - dtmStart.Value.Ticks) + _
        '                     dtmStart.Value.Ticks))

        '                strChartData(0, strChartData.GetUpperBound(1)) = _
        '                    dtmBinValue.ToShortDateString

        '                If drwSelectedAccount.IsNormalBalanceNull() Then
        '                    If drwSelectedAccount.Balance >= 0 Then
        '                        strChartData(1, strChartData.GetUpperBound(1)) = _
        '                                 dblBinValues(intBinNumber).ToString
        '                    Else
        '                        strChartData(1, strChartData.GetUpperBound(1)) = _
        '                                 -dblBinValues(intBinNumber).ToString
        '                    End If
        '                Else
        '                    If drwSelectedAccount.NormalBalance = Debit Then
        '                        strChartData(1, strChartData.GetUpperBound(1)) = _
        '                                     dblBinValues(intBinNumber).ToString
        '                    ElseIf drwSelectedAccount.NormalBalance = Credit Then
        '                        strChartData(1, strChartData.GetUpperBound(1)) = _
        '                                     -dblBinValues(intBinNumber).ToString
        '                    Else
        '                        If drwSelectedAccount.Balance >= 0 Then
        '                            strChartData(1, strChartData.GetUpperBound(1)) = _
        '                                     dblBinValues(intBinNumber).ToString
        '                        Else
        '                            strChartData(1, strChartData.GetUpperBound(1)) = _
        '                                     -dblBinValues(intBinNumber).ToString
        '                        End If
        '                    End If
        '                End If

        '            Next

        '            chDetailChart.ChartData = strChartData
        '            chDetailChart.DataGrid.RowLabel(1, 1) = strTitle

        '        ElseIf Me.radSubaccountBalances.Checked Then
        '            chDetailChart.chartType = MSChart20Lib.VtChChartType.VtChChartType2dPie
        '            chDetailChart.Plot.DataSeriesInRow = False
        '            chDetailChart.ShowLegend = True
        '            drwSelectedAccount = Me.tvwDetailedAccounts.SelectedNode.Tag

        '            If Not tvwDetailedAccounts.SelectedNode.FirstNode Is Nothing Then
        '                tndNode = tvwDetailedAccounts.SelectedNode.FirstNode
        '            ElseIf Not tvwDetailedAccounts.SelectedNode.Parent.FirstNode Is Nothing Then
        '                tndNode = tvwDetailedAccounts.SelectedNode.Parent.FirstNode
        '            Else
        '                tndNode = Nothing
        '            End If

        '            If Not tndNode Is Nothing Then
        '                strTitle = "Subaccounts Of " & tndNode.Parent.Tag.Path
        '                dblBalance = tndNode.Parent.Tag.Balance

        '                Do While Not tndNode Is Nothing
        '                    drwAccount = tndNode.Tag
        '                    If strChartData Is Nothing Then
        '                        ReDim strChartData(1, 0)
        '                    Else
        '                        ReDim Preserve strChartData(1, strChartData.GetUpperBound(1) + 1)
        '                    End If

        '                    intDelimiterIndex = drwAccount.Path.LastIndexOf(SubaccountDelimiter)
        '                    strChartData(0, strChartData.GetUpperBound(1)) = _
        '                        drwAccount.Path.Substring(intDelimiterIndex + 1) & ": " & _
        '                        (drwAccount.Balance / dblBalance).ToString("p")
        '                    strChartData(1, strChartData.GetUpperBound(1)) = drwAccount.Balance.ToString

        '                    tndNode = tndNode.NextNode
        '                Loop

        '                chDetailChart.ChartData = strChartData
        '                chDetailChart.DataGrid.RowLabel(1, 1) = strTitle
        '            End If
        '        End If
        '    End If
    End Sub

End Class
