Imports System.Drawing
Imports System.Drawing.Printing

Public Class clsDirectPrintSF
    Private p_oPrintFont As Font
    Private p_oPrintDocx As PrintDocument
    Private p_oPrintSetx As PrinterSettings
    Private p_nChrHeight As Single
    Private p_nChrWidthx As Single
    Private p_oDTMstr As DataTable
    Private p_oOthers As Collection

    Public Property PrintFont() As Font
        Get
            Return p_oPrintFont
        End Get
        Set(ByVal value As Font)
            p_oPrintFont = value
            p_nChrHeight = p_oPrintFont.GetHeight
            p_nChrWidthx = p_oPrintFont.Size

            Dim margins As New Margins(0, 0, 0, 0)
            p_oPrintDocx.DefaultPageSettings.Margins = margins
        End Set
    End Property

    Public Property Settings() As PrinterSettings
        Get
            Return p_oPrintSetx
        End Get
        Set(ByVal value As PrinterSettings)
            p_oPrintSetx = value
        End Set
    End Property

    Public Sub print_properties()
        MsgBox("Height: " & p_nChrHeight)
        MsgBox("Width : " & p_nChrWidthx)
    End Sub

    Public Sub Print(Row As Single, Column As Single, value As String)
        Print(Row, Column, value, StringAlignment.Near)
    End Sub

    Public Sub Print(Row As Single, Column As Single, value As String, alignment As Integer)
        Print(Row, Column, value, alignment, p_oPrintFont)
    End Sub

    Public Sub Print(Row As Single, Column As Single, value As String, alignment As Integer, oFont As Font)
        Print(Row, Column, value, alignment, p_oPrintFont, False)
    End Sub

    Public Sub Print(Row As Single, Column As Single, value As String, alignment As Integer, oFont As Font, isNextPage As Boolean)
        Dim oOthers As New Others

        oOthers.nPrntrRow = Row
        oOthers.nPrntrCol = Column
        oOthers.sMessagex = value
        oOthers.nAlignmnt = alignment
        oOthers.oPrintFnt = oFont
        oOthers.bNextPage = isNextPage
        p_oOthers.Add(oOthers)

    End Sub

    Public Function PrintEnd() As Boolean
        Using (p_oPrintDocx)
            If Not p_oPrintSetx Is Nothing Then p_oPrintDocx.PrinterSettings = p_oPrintSetx

            '// Adds a handler for PrintDocument.PrintPage 
            '(the sub PrintPageHandler)
            AddHandler p_oPrintDocx.PrintPage, _
               AddressOf Me.PrintPageHandler
            '\\

            p_oPrintDocx.Print() 'Prints.

            '// Removes the handler for PrintDocument.PrintPage 
            '(the sub PrintPageHandler)
            RemoveHandler p_oPrintDocx.PrintPage, _
               AddressOf Me.PrintPageHandler
            '\\
        End Using


        Dim loOthers As Others
        For Each loOthers In p_oOthers
            loOthers.oPrintFnt = Nothing
        Next
        p_oOthers.Clear()
        p_oOthers = Nothing
        Return True
    End Function

    Private Sub PrintPageHandler(ByVal sender As Object, _
        ByVal args As Printing.PrintPageEventArgs)

        Dim lftMargin As Single = args.MarginBounds.Left
        Dim topMargin As Single = args.MarginBounds.Top

        Dim lnCtr As Integer
        lnCtr = 1

        While lnCtr <= p_oOthers.Count
            Dim oOthers As Others = p_oOthers(lnCtr)
            Dim loSF As StringFormat = New StringFormat()
            loSF.Alignment = oOthers.nAlignmnt

            args.Graphics.DrawString(oOthers.sMessagex, _
               New Font(oOthers.oPrintFnt.Name, oOthers.oPrintFnt.Size), _
               Brushes.Black, _
               lftMargin + oOthers.nPrntrCol * 100, _
               topMargin + oOthers.nPrntrRow * p_nChrHeight, loSF)

            lnCtr = lnCtr + 1
        End While

        args.HasMorePages = False
    End Sub

    Public Sub PrintBegin()
        p_oOthers = New Collection
    End Sub

    Public Sub New()
        p_oPrintFont = New Font("Courier New", 10)
        p_oPrintDocx = New PrintDocument

        Dim margins As New Margins(0, 0, 0, 0)
        p_oPrintDocx.DefaultPageSettings.Margins = margins

        p_nChrHeight = p_oPrintFont.Height
        p_nChrWidthx = p_oPrintFont.Size

        Call PrintBegin()
    End Sub

    Private Class Others
        Public nPrntrRow As Single
        Public nPrntrCol As Single
        Public sMessagex As String
        Public nAlignmnt As Integer
        Public oPrintFnt As Font
        Public bNextPage As Boolean
    End Class
End Class
