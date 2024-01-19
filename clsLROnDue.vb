'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'     Client Master Object
'
' Copyright 2016 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Kalyptus [ 08/03/2017 10:57 am ]
'      Started creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports CrystalDecisions.CrystalReports.Engine

Public Class clsLROnDue
    Private psReportID As String
    Private psReportNm As String
    Private psReportFl As String
    Private psReportHd As String

    Private p_oAppDrivr As GRider
    Private p_oStandard As frmLRStandard
    Private p_oReport As ReportDocument
    Private p_oSTRept As DataSet
    Private p_oDTSrce As DataTable

    'Report Parameters
    Private p_nReptType As Integer      '1=Summary;2=Detail
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

        p_oStandard = New frmLRStandard

        InitReport = True
    End Function

    Function ProcessReport() As Boolean
        Debug.Print("Processing Report....")

        'Start try here
        Try
            With p_oStandard
                .Text = "LR Account On-Due Criteria"
                .AppDriver = p_oAppDrivr

                'Disable additional filtering
                .txtField82.Enabled = False
                .txtField81.Enabled = False
                .txtField83.Enabled = False
                .txtField80.Enabled = False

                'Set rbtDetail as Default
                .rbtSummary.Checked = False
                .rbtDetail.Checked = True

                'Disable radio buttons
                .rbtSummary.Enabled = False
                .rbtDetail.Enabled = False

                .ShowDialog()

                'Check if user pressed OK Button
                If Not .isOkey Then
                    MsgBox("Report Generation was Cancelled", vbInformation, "Notice")
                    GoTo endProc
                End If

                'Determient the type of report
                p_nReptType = .Presentation

                'Load additional report information
                Dim lsSQL As String
                lsSQL = "SELECT * FROM xxxReportDetail" & _
                       " WHERE sReportID = " & strParm(psReportID) & _
                         " AND nEntryNox = " & p_nReptType
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

                p_dDateFrom = CDate(.txtField00.Text)
                p_dDateThru = CDate(.txtField01.Text)

                p_oReport = New ReportDocument
                p_oReport.Load(p_oAppDrivr.AppPath & "\vb.net\Reports\" & loData(0).Item("sFileName") & ".rpt")

                If .Presentation = 2 Then
                    If prcLRReleaseDetail() = False Then GoTo endProc
                End If
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

    Private Function prcLRReleaseDetail() As Boolean
        Debug.Print("Collecting info from database....")

        'Start try here
        Try
            Dim lsSQL As String
            lsSQL = "SELECT c.sBranchNm, a.sAcctNmbr, b.sCompnyNm, a.dDueDatex, a.sCollNote, a.nABalance, a.dLastPaym, a.nLastPaym, a.nMonAmort, (a.nInterest - a.nIntTotal) nIntBlnce, (a.nPrincipl * a.nIntRatex / 100) nIntAmort" & _
                   " FROM LR_Master a" & _
                        " LEFT JOIN Client_Master b ON a.sClientID = b.sClientID" & _
                        " LEFT JOIN Branch c ON a.sBranchCd = c.sBranchCd" & _
                   " WHERE a.dDueDatex BETWEEN " & dateParm(p_dDateFrom) & " AND " & dateParm(p_dDateThru) & _
                      IIf(p_oAppDrivr.ProductID = "LRTrackr", "", " AND b.sTransNox LIKE " & strParm(p_oAppDrivr.BranchCode & "%")) & _
                   " ORDER BY c.sBranchNm, a.dDueDatex, b.sCompnyNm"
            p_oDTSrce = p_oAppDrivr.ExecuteQuery(lsSQL)

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
            loTxtObj.Text = "LR Account On-Due Report (" & Format(p_dDateFrom, "yyyy-MM-dd") & " TO " & Format(p_dDateThru, "yyyy-MM-dd") & ")"

            'Set Second Header
            loTxtObj = p_oReport.ReportDefinition.Sections(1).ReportObjects("txtHeading2")
            loTxtObj.Text = "As of " & Format(p_oAppDrivr.SysDate, "MMMM dd, yyyy")

            loTxtObj = p_oReport.ReportDefinition.Sections(6).ReportObjects("txtRptUser")
            loTxtObj.Text = Decrypt(p_oAppDrivr.UserName, "08220326")

            'MsgBox(p_oSTRept.Tables(0).Rows.Count)
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

            loDtaRow.Item("sField01") = p_oDTSrce(lnRow).Item("sBranchNm")
            loDtaRow.Item("sField02") = p_oDTSrce(lnRow).Item("sAcctNmbr")
            loDtaRow.Item("sField03") = p_oDTSrce(lnRow).Item("sCompnyNm")
            loDtaRow.Item("sField04") = Format(p_oDTSrce(lnRow).Item("dDueDatex"), "yyyy/MM/dd")
            loDtaRow.Item("sField05") = Format(p_oDTSrce(lnRow).Item("dLastPaym"), "yyyy/MM/dd")
            loDtaRow.Item("sField06") = p_oDTSrce(lnRow).Item("sCollNote")

            loDtaRow.Item("lField01") = p_oDTSrce(lnRow).Item("nABalance")
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
