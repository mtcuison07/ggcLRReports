Imports ggcAppDriver
Imports System.Windows.Forms

Public Class frmReports
    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_nRowSel As Integer
    Private pnLoadx As Integer

    Public Property AppDriver() As GRider
        Get
            Return p_oApp
        End Get
        Set(ByVal value As GRider)
            p_oApp = value
        End Set
    End Property

    Public ReadOnly Property ReportID() As String
        Get
            If p_nRowSel >= 0 Then
                Return p_oDTMstr(p_nRowSel).Item("sReportID")
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property ReportName() As String
        Get
            If p_nRowSel >= 0 Then
                Return p_oDTMstr(p_nRowSel).Item("sReportNm")
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property ReportLibrary() As String
        Get
            If p_nRowSel >= 0 Then
                Return p_oDTMstr(p_nRowSel).Item("sRepLibxx")
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property ReportClass() As String
        Get
            If p_nRowSel >= 0 Then
                Return p_oDTMstr(p_nRowSel).Item("sRepClass")
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property LogReport() As Boolean
        Get
            If p_nRowSel >= 0 Then
                Return p_oDTMstr(p_nRowSel).Item("cLogRepxx") = "1"
            Else
                Return False
            End If
        End Get
    End Property

    Private Sub LoadList()
        Dim lsProcName As String
        Dim lsSQL As String
        Dim lnCtr As Integer

        lsProcName = "LoadList"

        lsSQL = "SELECT" & _
                    "  sReportID" & _
                    ", sReportNm" & _
                    ", sRepLibxx" & _
                    ", sRepClass" & _
                    ", cLogRepxx" & _
                    ", cSaveRepx" & _
                 " FROM xxxReportMaster" & _
                 " WHERE sProdctID LIKE " & strParm("%" & p_oApp.ProductID & "%") & _
                    " AND nUserRght & " & p_oApp.UserLevel & " > 0" & _
                    " AND sRepLibxx IN('ggcLRReports')" & _
                 " ORDER BY sReportNm"

        p_oDTMstr = p_oApp.ExecuteQuery(lsSQL)

        With dgView
            lnCtr = 0
            Do While lnCtr < p_oDTMstr.Rows.Count
                .Rows.Add()
                .Rows(lnCtr).Cells(0).Value = p_oDTMstr(lnCtr).Item("sReportNm")
                lnCtr = lnCtr + 1
            Loop
        End With
    End Sub

    Private Sub frmReports_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        Debug.Print("frmReports_Activated")
        If pnLoadx = 1 Then
            Call LoadList()
            pnLoadx = 2
        End If
    End Sub

    Private Sub frmReports_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Debug.Print("frmReports_Load")
        If pnLoadx = 0 Then
            Call grpEventHandler(Me, GetType(Button), "cmdButtn", "Click", AddressOf cmdButton_Click)
            pnLoadx = 1
        End If
    End Sub

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim loChk As Button
        loChk = CType(sender, System.Windows.Forms.Button)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        Select Case lnIndex
            Case 0 ' Ok
                p_nRowSel = dgView.CurrentRow.Index
                Me.Hide()
            Case 1 ' Cancel Update
                p_nRowSel = -1
                Me.Hide()
        End Select
    End Sub
End Class