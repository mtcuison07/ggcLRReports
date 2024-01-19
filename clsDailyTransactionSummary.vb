Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports CrystalDecisions.CrystalReports.Engine

Public Class clsDailyTransactionSummary
    Private psReportID As String
    Private psReportNm As String
    Private psReportFl As String
    Private psReportHd As String

    Private p_oAppDrivr As GRider
    Private p_oStandard As frmDateCriteria
    Private p_oReport As ReportDocument
    Private p_oSTRept As DataSet
    Private p_oDTSrce As DataTable

    'Report Parameters
    Private p_dDateFrom As Date
    Private p_dDateThru As Date

    Public Property AppDriver() As GRider
        Get
            Return p_oAppDrivr
        End Get
        Set(ByVal value As GRider)
            p_oAppDrivr = value
        End Set
    End Property

    Public ReadOnly Property ReportSource() As ReportDocument
        Get
            Return p_oReport
        End Get
    End Property

    Sub CloseReport()
        p_oStandard = Nothing
        Debug.Print("Closing report....")
    End Sub

    Function InitReport(ByVal ReportID As String, ByVal ReportName As String) As Boolean
        Debug.Print("Initializing Report....")

        psReportID = ReportID
        psReportNm = ReportName

        p_oStandard = New frmDateCriteria

        InitReport = True
    End Function

    Function ProcessReport() As Boolean
        Debug.Print("Processing Report....")

        'Start try here
        Try
            With p_oStandard
                .Text = "Date Criteria"
                .AppDriver = p_oAppDrivr

                .ShowDialog()

                'Check if user pressed OK Button
                If Not .isOkey Then
                    MsgBox("Report Generation was Cancelled", vbInformation, "Notice")
                    GoTo endProc
                End If


                'Load additional report information
                Dim lsSQL As String
                lsSQL = "SELECT * FROM xxxReportDetail" & _
                       " WHERE sReportID = " & strParm(psReportID) & _
                         " AND nEntryNox = 1"

                Dim loData As DataTable
                loData = p_oAppDrivr.ExecuteQuery(lsSQL)

                'Exit if no additional report information was retrieved...
                If loData.Rows.Count = 0 Then
                    CloseReport()
                    MsgBox("Unable to Retrieve Report Info..." & lsSQL, vbCritical, "Warning")
                    GoTo endProc
                End If

                psReportFl = loData(0).Item("sFileName")
                psReportHd = loData(0).Item("sReportHd")

                p_dDateFrom = CDate(.txtDateFrom.Text)
                p_dDateThru = CDate(.txtDateThru.Text)

                p_oReport = New ReportDocument
                p_oReport.Load(p_oAppDrivr.AppPath & "\vb.net\Reports\" & loData(0).Item("sFileName") & ".rpt")

                If prcOfficialReceipt() = False Then GoTo endProc

            End With

            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Throw ex
            MsgBox("Lumampas ako sa Throw")
            GoTo endProc
        End Try
        Return True
endProc:
        Return False
    End Function

    Private Function prcOfficialReceipt() As Boolean
        Debug.Print("Collecting info from database....")

        'Start try here
        Try
            Dim lsSQL As String
            lsSQL = "SELECT" & _
                        " a.sTransNox" & _
                        ", b.sBranchCd" & _
                        ", b.sBranchNm" & _
                        ", a.dTransact" & _
                        ", a.sAcctNmbr" & _
                        ", Case a.cTranType" & _
                        " WHEN '0' THEN 'LR' " & _
                        " WHEN '2' THEN 'MP' " & _
                        " WHEN '3' THEN 'CB' " & _
                        " WHEN '4' THEN 'DB' " & _
                        " END cTranType" & _
                        ", a.sReferNox" & _
                        ", a.nIntAmtxx" & _
                        ", a.nRebatesx" & _
                        ", a.nPenaltyX" & _
                        ", a.nAmountxx" & _
                        ", c.sCompnyNm" & _
                        ", a.sSourceCD" & _
                    " FROM LR_Payment_Master a" & _
                        " LEFT JOIN Client_Master c" & _
                            " ON a.sClientID = c.sClientID" & _
                    ", Branch b" & _
                    " WHERE LEFT(a.sTransNox,4) = b.sBranchCd" & _
                    " AND b.sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & _
                    " AND a.dTransact BETWEEN " & dateParm(p_oStandard.txtDateFrom.Text) & " AND " & dateParm(p_oStandard.txtDateThru.Text) & _
                    " AND a.cPostedxx = " & strParm(xeTranStat.TRANS_POSTED) & _
                    " ORDER BY a.sReferNox"


            p_oDTSrce = p_oAppDrivr.ExecuteQuery(lsSQL)

            Dim loDtaTbl As DataTable = getRptTable()

            For lnCtr = 0 To p_oDTSrce.Rows.Count - 1
                loDtaTbl.Rows.Add(addRow(lnCtr, loDtaTbl))
            Next

            Dim loTxtObj As CrystalDecisions.CrystalReports.Engine.TextObject
            loTxtObj = p_oReport.ReportDefinition.Sections(0).ReportObjects("txtCompany")
            loTxtObj.Text = p_oAppDrivr.BranchName

            'Set Branch Address
            loTxtObj = p_oReport.ReportDefinition.Sections(0).ReportObjects("txtAddress")
            loTxtObj.Text = p_oAppDrivr.Address

            'Set First Header
            loTxtObj = p_oReport.ReportDefinition.Sections(1).ReportObjects("txtHeading1")
            loTxtObj.Text = psReportNm

            'Set Second Header
            loTxtObj = p_oReport.ReportDefinition.Sections(1).ReportObjects("txtHeading2")
            loTxtObj.Text = Format(p_dDateFrom, "yyyy-MM-dd") & " TO " & Format(p_dDateThru, "yyyy-MM-dd") & ")"

            loTxtObj = p_oReport.ReportDefinition.Sections(3).ReportObjects("txtRptUser")
            loTxtObj.Text = Decrypt(p_oAppDrivr.UserName, "08220326")

            p_oReport.SetDataSource(p_oSTRept)
            'p_oReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, "d:\what.pdf")

            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Throw ex
            MsgBox("Lumampas ako sa Throw")
            GoTo endProc
        End Try

        Return True
endProc:
        Return False
    End Function

    Private Function prcProvisionaryReceipt() As Boolean
        Debug.Print("Collecting info from database....")

        'Start try here
        Try
            Dim lsSQL As String
            lsSQL = "SELECT" & _
                        " a.sTransNox" & _
                        ", b.sBranchCd" & _
                        ", b.sBranchNm" & _
                        ", a.dTransact" & _
                        ", a.sAcctNmbr" & _
                        ", Case a.cTranType" & _
                        " WHEN '0' THEN 'LR' " & _
                        " WHEN '2' THEN 'MP' " & _
                        " WHEN '3' THEN 'CB' " & _
                        " WHEN '4' THEN 'DB' " & _
                        " END cTranType" & _
                        ", a.sReferNox" & _
                        ", a.nIntAmtxx" & _
                        ", a.nRebatesx" & _
                        ", a.nPenaltyX" & _
                        ", a.nAmountxx" & _
                        ", c.sCompnyNm" & _
                    " FROM LR_Payment_Master_PR a" & _
                        " LEFT JOIN Client_Master c" & _
                            " ON a.sClientID = c.sClientID" & _
                    ", Branch b" & _
                    " WHERE LEFT(a.sTransNox,4) = b.sBranchCd" & _
                    " AND b.sBranchCd = " & strParm(p_oAppDrivr.BranchCode) & _
                    " AND a.dTransact BETWEEN " & dateParm(p_oStandard.txtDateFrom.Text) & " AND " & dateParm(p_oStandard.txtDateThru.Text) & _
                    " AND a.cPostedxx = " & strParm(xeTranStat.TRANS_POSTED)


            p_oDTSrce = p_oAppDrivr.ExecuteQuery(lsSQL)

            Debug.Print(lsSQL)
            Dim loDtaTbl As DataTable = getRptTable()
            Dim lnCtr As Integer

            For lnCtr = 0 To p_oDTSrce.Rows.Count - 1
                loDtaTbl.Rows.Add(addRow(lnCtr, loDtaTbl))
            Next

            Dim loTxtObj As CrystalDecisions.CrystalReports.Engine.TextObject
            loTxtObj = p_oReport.ReportDefinition.Sections(0).ReportObjects("txtCompany")
            loTxtObj.Text = p_oAppDrivr.BranchName

            'Set Branch Address
            loTxtObj = p_oReport.ReportDefinition.Sections(0).ReportObjects("txtAddress")
            loTxtObj.Text = p_oAppDrivr.Address

            'Set First Header
            loTxtObj = p_oReport.ReportDefinition.Sections(1).ReportObjects("txtHeading1")
            loTxtObj.Text = psReportNm

            'Set Second Header
            loTxtObj = p_oReport.ReportDefinition.Sections(1).ReportObjects("txtHeading2")
            loTxtObj.Text = Format(p_dDateFrom, "yyyy-MM-dd") & " TO " & Format(p_dDateThru, "yyyy-MM-dd") & ")"

            loTxtObj = p_oReport.ReportDefinition.Sections(3).ReportObjects("txtRptUser")
            loTxtObj.Text = Decrypt(p_oAppDrivr.UserName, "08220326")

            p_oReport.SetDataSource(p_oSTRept)
            'p_oReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, "d:\what.pdf")

            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Throw ex
            MsgBox("Lumampas ako sa Throw")
            GoTo endProc
        End Try

        Return True
endProc:
        Return False
    End Function

    Private Function getRptTable() As DataTable
        Try
            'Initialize DataSet
            p_oSTRept = New DataSet

            'Load the data structure of the Dataset
            'Data structure was saved at DataSet1.xsd 
            p_oSTRept.ReadXmlSchema(p_oAppDrivr.AppPath & "\vb.net\Reports\DataSet1.xsd")

            'Return the schema of the datatable derive from the DataSet 
            Return p_oSTRept.Tables(0)
        Catch ex As Exception
            MsgBox(ex.Message)
            Throw ex
        End Try
    End Function

    Private Function addRow(ByVal lnRow As Integer, ByVal foSchemaTable As DataTable) As DataRow
        Try
            'ByVal foDTInclue As DataTable
            Dim loDtaRow As DataRow

            'Create row based on the schema of foSchemaTable
            loDtaRow = foSchemaTable.NewRow

            'loDtaRow.Item("sField01") = p_oDTSrce(lnRow).Item("sBranchNm")
            loDtaRow.Item("sField02") = p_oDTSrce(lnRow).Item("sAcctNmbr")
            loDtaRow.Item("sField03") = p_oDTSrce(lnRow).Item("sCompnyNm")
            loDtaRow.Item("sField04") = p_oDTSrce(lnRow).Item("cTranType")
            loDtaRow.Item("sField05") = p_oDTSrce(lnRow).Item("sReferNox")
            loDtaRow.Item("sField06") = Format(p_oDTSrce(lnRow).Item("dTransact"), "yyyy/MM/dd")
            loDtaRow.Item("lField01") = p_oDTSrce(lnRow).Item("nAmountxx")
            loDtaRow.Item("lField02") = p_oDTSrce(lnRow).Item("nIntAmtxx")
            loDtaRow.Item("lField03") = p_oDTSrce(lnRow).Item("nPenaltyx")

            Return loDtaRow
        Catch ex As Exception
            MsgBox(ex.Message)
            Throw ex
        End Try
    End Function

    Public Sub New()
        psReportID = ""
        psReportNm = ""
        psReportFl = ""
        psReportHd = ""
    End Sub
End Class

