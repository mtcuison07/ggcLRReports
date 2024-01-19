Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver
Imports CrystalDecisions.CrystalReports.Engine

Public Class clsLRRep
    Private p_oAppDrivr As GRider
    Private p_oRepBrowse As frmReports
    Private p_oRepSource As Object

    Public ReadOnly Property ReportSource() As ReportDocument
        Get
            Return p_oRepSource.ReportSource
        End Get
    End Property

    Public Function ShowReport() As Boolean
        Dim lsProcName As String
        Dim lsRepName As String
        Dim lsSQL As String
        Dim lnRow As Long

        lsProcName = "ShowReports"

        p_oRepBrowse = New frmReports
        p_oRepBrowse.AppDriver = p_oAppDrivr

        ShowReport = False

        With p_oRepBrowse
            .ShowDialog()

            'MsgBox("You Selected she" + .ReportName)

            Try
                p_oRepSource = Activator.CreateInstance(Type.GetType(Trim(.ReportLibrary) & "." & Trim(.ReportClass)))
                p_oRepSource.AppDriver = p_oAppDrivr
                p_oRepSource.InitReport(.ReportID, .ReportName)
                If p_oRepSource.ProcessReport = False Then GoTo endProc

            Catch ex As Exception
                MsgBox(ex.Message)
                Throw ex
                MsgBox("Lumampas ako sa Throw")
                GoTo endProc
            End Try

            lsRepName = getReportCode()

            If .LogReport Then
                With p_oAppDrivr
                    lsSQL = "INSERT INTO xxxReportsLog (" & _
                                "  sReportID" & _
                                ", dGenerate" & _
                                ", sUserIDxx" & _
                                ", sRepFName" & _
                             " ) VALUES (" & _
                                strParm(p_oRepBrowse.ReportID) & _
                                ", " & dateParm(.getSysDate) & _
                                ", " & strParm(.UserID) & _
                                ", " & strParm(lsRepName) & " )"
                    lnRow = .ExecuteActionQuery(lsSQL)
                    If lnRow = 0 Then
                        MsgBox("Unable to Register Report Generation!!!", vbCritical, "Warning")
                        GoTo endProc
                    End If
                End With
            End If
        End With

        ShowReport = True

endProc:
        p_oRepBrowse.Close()
        p_oRepBrowse = Nothing

        Exit Function
    End Function

    Sub CloseReport()
        p_oRepSource.CloseReport()
    End Sub

    Private Function getReportCode() As String
        Dim loDta As DataTable
        Dim lsProcName As String
        Dim lsSQL As String

        lsProcName = "getReportCode"

        Try
            ' first get the computer id
            getReportCode = Format(p_oAppDrivr.getSysDate, "yy") & _
                     p_oAppDrivr.BranchCode

            ' then get the latest Report name based on the year and computer id
            lsSQL = "SELECT sRepFName " & _
                     " FROM xxxReportsLog" & _
                     " WHERE sRepFName LIKE " & strParm(getReportCode & "%") & _
                     " ORDER BY sRepFName DESC" & _
                     " LIMIT 1"

            loDta = p_oAppDrivr.ExecuteQuery(lsSQL)

            If loDta.Rows.Count = 0 Then
                getReportCode = getReportCode & Format(1, "000000")
            Else
                getReportCode = getReportCode & Format(Val(Right(loDta(0).Item("sRepFName"), 4)) + 1, "000000")
            End If
        Catch ex As Exception
            getReportCode = ""
            MsgBox(ex.Message)
            Throw ex
        End Try
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oAppDrivr = foRider
    End Sub
End Class
