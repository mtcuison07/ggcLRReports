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

Public Class clsActiveAcctSummary
    Private psReportID As String
    Private psReportNm As String
    Private psReportFl As String
    Private psReportHd As String

    Private p_oAppDrivr As GRider

    Public Property AppDriver() As GRider
        Get
            Return p_oAppDrivr
        End Get
        Set(ByVal value As GRider)
            p_oAppDrivr = value
        End Set
    End Property

    Sub CloseReport()
        ' p_oProgress.CloseProgress()
        ' p_oStandard = Nothing
        Debug.Print("Closing report....")
    End Sub

    Function InitReport(ByVal ReportID As String, ByVal ReportName As String) As Boolean
        Debug.Print("Initializing Report....")

        psReportID = ReportID
        psReportNm = ReportName

        'p_oStandard = New frmRegRepCriteria
        'p_oStandard.AppDriver = p_oAppDrivr

        InitReport = True
    End Function

    Function ProcessReport() As Boolean
        Dim lors As Recordset, loTemp As Recordset
        Dim lsProcName As String
        Dim lnEntryNo As Integer

        Debug.Print("Processing Report....")

        ''Start try here
        'Try
        '    With p_oStandard
        '        .Caption = "Collection Summary Criteria"
        '        .Show(1)

        '        If .Cancelled Then
        '            MsgBox("Report Generation was Cancelled", vbInformation, "Notice")
        '            GoTo endProc
        '        End If

        '        If .Presentation = 2 Then
        '            lnEntryNo = 5
        '        Else
        '            lnEntryNo = .Presentation + 1
        '        End If
        '        If .Presentation = 1 Then
        '            ' Detail Presentation, then show collection type
        '            p_oPaymLoc.Caption = "Collection Summary Criteria"
        '            p_oPaymLoc.Show(1)

        '            If p_oPaymLoc.Cancelled Then
        '                MsgBox("Report Generation was Cancelled", vbInformation, "Notice")
        '                GoTo endProc
        '            End If

        '            lnEntryNo = p_oPaymLoc.Presentation + 2
        '        End If

        '        With p_oProgress
        '            .InitProgress("Processing...", 5, 3)
        '            .PrimaryRemarks = "Processing Report"
        '            .MoveProgress("Setting Retrieval Info...")

        '            psSQL = "SELECT * FROM xxxReportDetail" & _
        '                    " WHERE sReportID = " & strParm(psReportID) & _
        '                       " and nEntryNox = " & lnEntryNo
        '            .MoveProgress("Retriving Report Specification...")

        '            p_oRawSource.Open(psSQL, p_oAppDrivr.Connection, , , adCmdText)

        '            .MoveProgress("Processing Report Specification...")
        '        End With

        '        If p_oRawSource.EOF Then
        '            CloseReport()
        '            MsgBox("Unable to Retrieve Report Info..." & psSQL, vbCritical, "Warning")
        '            GoTo endProc
        '        End If

        '        psReportFl = p_oRawSource("sFileName")
        '        psReportHd = p_oRawSource("sReportHd")
        '        p_oRawSource.Close()

        '        If .Presentation = 1 Then
        '            If prcBranchCollection = False Then GoTo endProc
        '        Else
        '            If prcSummaryPresentation = False Then GoTo endProc
        '        End If
        '    End With

        '    Return True
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        '    Throw ex
        '    MsgBox("Lumampas ako sa Throw")
        '    GoTo endProc
        'End Try
        Return True
endProc:
    End Function

    Public Sub New()
        psReportID = ""
        psReportNm = ""
        psReportFl = ""
        psReportHd = ""
    End Sub

End Class
