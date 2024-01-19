Imports ggcAppDriver
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmLRStandard
    Private p_oApp As GRider
    Private poControl As Control

    Private p_sCompnyID As String
    Private p_sBranchCD As String
    Private p_sAreaCode As String
    Private p_sCollctID As String
    Private p_nButton As Integer = -1
    Private pnLoadx As Integer

    Public WriteOnly Property AppDriver() As GRider
        Set(ByVal value As GRider)
            p_oApp = value
        End Set
    End Property

    Public ReadOnly Property isOkey() As Boolean
        Get
            Return p_nButton = 0
        End Get
    End Property

    Public ReadOnly Property CompanyID() As String
        Get
            If p_nButton = 0 Then
                Return p_sCompnyID
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property BranchCode() As String
        Get
            If p_nButton = 0 Then
                Return p_sBranchCD
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property AreaCode() As String
        Get
            If p_nButton = 0 Then
                Return p_sAreaCode
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property CollectorID() As String
        Get
            If p_nButton = 0 Then
                Return p_sCollctID
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property Presentation() As Integer
        Get
            If p_nButton = 0 Then
                If rbtSummary.Checked Then
                    Return 1
                ElseIf rbtDetail.Checked Then
                    Return 2
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        End Get
    End Property

    Private Sub frmLRStandard_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Debug.Print("frmLRStandard_Activated")
        If pnLoadx = 1 Then
            txtField00.Focus()
            pnLoadx = 2
        End If
    End Sub

    Private Sub frmLRStandard_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F3 Or e.KeyCode = Keys.Enter Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)
            Dim loIndex As Integer
            loIndex = Val(Mid(loTxt.Name, 9))

            If Mid(loTxt.Name, 1, 8) = "txtField" Then
                Select Case loIndex
                    Case 80
                    Case 81
                    Case 82
                    Case 83
                End Select
            End If

            If TypeOf poControl Is TextBox Then
                SelectNextControl(loTxt, True, True, True, True)
            End If
        End If
    End Sub

    Private Sub frmLRStandard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Debug.Print("frmLRStandard_Load")
        If pnLoadx = 0 Then

            'Set event Handler for txtField
            Call grpEventHandler(Me, GetType(TextBox), "txtField", "GotFocus", AddressOf txtField_GotFocus)
            Call grpEventHandler(Me, GetType(TextBox), "txtField", "LostFocus", AddressOf txtField_LostFocus)
            Call grpEventHandler(Me, GetType(Button), "cmdButtn", "Click", AddressOf cmdButton_Click)
            pnLoadx = 1
        End If
    End Sub

    'Handles GotFocus Events for txtField & txtItems
    Private Sub txtField_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim loIndex As Integer
        loIndex = Val(Mid(sender.Name, 9))

        Console.WriteLine("»Got Focus: " & sender.Name)

        If Mid(sender.Name, 1, 8) = "txtField" Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)

            If loTxt.Enabled Then
                Select Case loIndex
                    Case 0, 1
                        If IsDate(loTxt.Text) Then
                            loTxt.Text = Format(CDate(loTxt.Text), "yyyy/MM/dd")
                        Else
                            loTxt.Text = ""
                        End If
                End Select

                loTxt.BackColor = Color.Azure
                loTxt.SelectAll()
            End If

            poControl = loTxt
        End If
    End Sub

    'Handles LostFocus Events for txtField & txtItems
    Private Sub txtField_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)

        Console.WriteLine("Lost Focus: " & sender.Name)

        Dim loIndex As Integer
        loIndex = Val(Mid(sender.Name, 9))

        If Mid(sender.Name, 1, 8) = "txtField" Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)

            If loTxt.Enabled Then
                Select Case loIndex
                    Case 0, 1
                        If IsDate(loTxt.Text) Then
                            loTxt.Text = Format(CDate(loTxt.Text), "MMMM dd, yyyy")
                        Else
                            loTxt.Text = ""
                        End If
                End Select

                loTxt.BackColor = SystemColors.Window
            End If
        End If
    End Sub


    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim loChk As Button
        loChk = CType(sender, System.Windows.Forms.Button)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        Select Case lnIndex
            Case 0 ' Ok
                If isEntryOk() Then
                    p_nButton = 0
                    Me.Hide()
                End If
            Case 1 ' Cancel Update
                p_nButton = 1
                Me.Hide()
        End Select
    End Sub

    Private Function isEntryOk() As Boolean
        If Not IsDate(txtField00.Text) Then
            MsgBox("Invalid date from detected!", vbOKOnly, "LR Standard Report Criteria")
            txtField00.Focus()
            Return False
        ElseIf Not IsDate(txtField01.Text) Then
            MsgBox("Invalid date thru detected...", vbOKOnly, "LR Standard Report Criteria")
            txtField01.Focus()
            Return False
        End If

        If CDate(txtField00.Text) > CDate(txtField01.Text) Then
            MsgBox("Date from is higher than date thru...", vbOKOnly, "LR Standard Report Criteria")
            txtField01.Focus()
            Return False
        End If

        Return True
    End Function

End Class