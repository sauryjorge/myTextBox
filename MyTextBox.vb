Imports System.ComponentModel

Public Class myTextBox
    Inherits TextBox

    Dim TextError As New ErrorProvider

#Region "AutoSize"
    <Browsable(True)>
    <EditorBrowsable(EditorBrowsableState.Always)>
    <DefaultValue(True)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)>
    Public Overrides Property AutoSize() As Boolean
        Get
            Return MyBase.AutoSize
        End Get
        Set(ByVal value As Boolean)
            MyBase.AutoSize = value
        End Set
    End Property
#End Region

#Region "Mascara de Agua"
#Region "Attributes"
    Private _waterMarkColor As Color = Drawing.Color.Gray

    Public Property WaterMarkColor() As Color
        Get
            Return _waterMarkColor
        End Get
        Set(ByVal value As Color)
            _waterMarkColor = value
            Me.Invalidate()
        End Set
    End Property

    Private _waterMarkText As String = "Marca de Agua"

    Public Property WaterMarkText() As String
        Get
            Return _waterMarkText
        End Get
        Set(ByVal value As String)
            _waterMarkText = value
            Me.Invalidate()
        End Set
    End Property
#End Region

    Private oldFont As Font = Nothing
    Private waterMarkTextEnabled As Boolean = False

    Private Sub JoinEvents(ByVal join As Boolean)
        If join Then
            AddHandler(TextChanged), AddressOf WaterMark_Toggle
            AddHandler(LostFocus), AddressOf WaterMark_Toggle
            AddHandler(FontChanged), AddressOf WaterMark_FontChanged
            'No one of the above events will start immeddiatlly 
            'TextBox control still in constructing, so,
            'Font object (for example) couldn't be catched from within WaterMark_Toggle
            'So, call WaterMark_Toggel through OnCreateControl after TextBox is totally created
            'No doupt, it will be only one time call

            'Old solution uses Timer.Tick event to check Create property
        End If
    End Sub

    Private Sub WaterMark_Toggle(ByVal sender As Object, ByVal args As EventArgs)
        If Me.Text.Length <= 0 Then
            EnableWaterMark()
        Else
            DisableWaterMark()
        End If
    End Sub

    Private Sub WaterMark_FontChanged(ByVal sender As Object, ByVal args As EventArgs)
        If waterMarkTextEnabled Then
            oldFont = New Font(Font.FontFamily, Font.Size, Font.Style, Font.Unit)
            Refresh()
        End If
    End Sub

    Private Sub EnableWaterMark()
        'Save current font until returning the UserPaint style to false (NOTE: It is a try and error advice)
        oldFont = New Font(Font.FontFamily, Font.Size, Font.Style, Font.Unit)

        'Enable OnPaint Event Handler
        Me.SetStyle(ControlStyles.UserPaint, True)
        Me.waterMarkTextEnabled = True

        'Trigger OnPaint immediatly
        Refresh()

    End Sub

    Private Sub DisableWaterMark()
        'Disbale OnPaint event handler
        Me.waterMarkTextEnabled = False
        Me.SetStyle(ControlStyles.UserPaint, False)

        'Return oldFont if existed
        If Not oldFont Is Nothing Then
            Me.Font = New Font(Font.FontFamily, Font.Size, Font.Style, Font.Unit)
        End If
    End Sub

    ' Override OnCreateControl 
    Protected Overrides Sub OnCreateControl()
        MyBase.OnCreateControl()
        WaterMark_Toggle(Nothing, Nothing)
    End Sub

    ' Override OnPaint
    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        ' Use the same font that was defined in base class
        Dim fontItaly As New FontStyle
        fontItaly = FontStyle.Italic
        Dim drawFont As Font = New Font(Font.FontFamily, Font.Size, fontItaly, Font.Unit)
        ' Create new brush with gray color or 
        Dim drawBrush As SolidBrush = New SolidBrush(Me.WaterMarkColor) 'use WaterMarkColor
        ' Draw Test or WaterMark
        e.Graphics.DrawString(IIf(waterMarkTextEnabled, WaterMarkText, Text).ToString(), drawFont, drawBrush, New Point(0, 0))
        MyBase.OnPaint(e)
    End Sub
#End Region

#Region "Propiedades"
    Private numFormat As Boolean

    <System.ComponentModel.Category("NumericBehavior")>
    Property FormatoNumerico() As Boolean
        Get
            Return numFormat
        End Get
        Set(ByVal value As Boolean)
            numFormat = value
            If value Then TextAlign = HorizontalAlignment.Right
            If Not value Then TextAlign = HorizontalAlignment.Left

            If FormatoNumerico = True Then _
            If IsNumeric(Text) = False Then Text = "0.00"
        End Set
    End Property

    Private max As ULong = 999999999999
    <System.ComponentModel.Category("NumericData")>
    Property Maximum() As ULong
        Get
            Return max
        End Get
        Set(ByVal value As ULong)
            If numFormat Then max = value
        End Set
    End Property

    Private decPlc As Byte = 2
    <System.ComponentModel.Category("NumericData")>
    <System.ComponentModel.Description("Cantidad de decimales")>
    Property DecimalPlaces() As Byte
        Get
            Return decPlc
        End Get
        Set(ByVal value As Byte)
            If numFormat = True Then decPlc = value
        End Set
    End Property

    Private decSep As Char = "."
    <System.ComponentModel.Category("NumericData")>
    Property DecimalSeparator() As Char
        Get
            Return decSep
        End Get
        Set(ByVal value As Char)
            If value <> "," And value <> "." Then Return
            If numFormat Then decSep = value
        End Set
    End Property


#End Region

#Region "Eventos"
    Protected Overrides Sub OnGotFocus(e As EventArgs)
        '2015-11-16 03:43pm SSJF

        'Cuando el control no esta habilitado no cambia de color
        If Me.ReadOnly = False Then Me.BackColor = Color.Yellow

        'Cuando el control esta en estado numerico y con valor 0.00 entonces pone 'VACIO' cuando recibe el foco
        If Me.FormatoNumerico = True Then If Me.Text = "0.00" And Me.ReadOnly = False Then Me.Text = String.Empty
        If Me.FormatoNumerico = True Then If Me.Text = "0" And Me.ReadOnly = False Then Me.Text = String.Empty

        MyBase.OnGotFocus(e)
    End Sub
    Protected Overrides Sub OnLostFocus(ByVal e As System.EventArgs)
        '2015-11-16 03:43pm SSJF

        'Cuando el control no esta habilitado no cambia de color
        If Me.ReadOnly = False Then Me.BackColor = Color.White
        Try
            If FormatoNumerico = True Then
                If Text = String.Empty And Me.ReadOnly = False Then Text = "0.00"
                Text = FormatNumber(Decimal.Parse(Text), decPlc)
            End If
        Catch ex As Exception
            TextError.SetError(Me, ex.Message)
            Focus()
            Exit Sub
        End Try
        MyBase.OnLostFocus(e)
    End Sub

    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Me.FormatoNumerico Then
            Dim keyAsc As Short = Asc(e.KeyChar)
            Select Case keyAsc
                Case 48 To 57, 8, Asc(decSep)
                    If keyAsc = Asc(decSep) And DecimalPlaces > 0 Then e.Handled = True : Exit Sub
                    If keyAsc = Asc(decSep) AndAlso Text.Contains(decSep) Then e.Handled = True
                Case Else
                    e.Handled = True
            End Select
        End If
        MyBase.OnKeyPress(e)
    End Sub

    Private preText As String

    Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
        TextError.Clear()
        On Error Resume Next
        If FormatoNumerico AndAlso Text <> "" Then
            Dim intLength As Short = Text.IndexOf(decSep)
            Dim decLength As Short = Text.Length - intLength - 1
            If intLength = -1 Then
                intLength = Text.Length
                decLength = 0
            End If
            Dim intValue As Decimal = Microsoft.VisualBasic.Left(Double.Parse(Text), intLength)
            If intValue > Maximum Or decLength > DecimalPlaces Then
                Dim curPos As Integer = SelectionStart
                Text = preText
                SelectionStart = curPos - 1
            End If

            'Puesto para que maneje los millares mientras escriben
            'SSJF 2016-08-09
            Dim Position As Integer = SelectionStart
            Dim CantidadCaracteres = Len(FormatNumber(Replace(Text, ",", ""), decPlc))

            If Len(Text) - CantidadCaracteres = -1 Then
                Text = FormatNumber(Replace(Text, ",", ""), decPlc)
                SelectionStart = Position + 1
            ElseIf Len(Text) - CantidadCaracteres = 1 Then
                Text = FormatNumber(Replace(Text, ",", ""), decPlc)
                SelectionStart = Position - 1
            Else
                Text = FormatNumber(Replace(Text, ",", ""), decPlc)
                SelectionStart = Position
            End If
            preText = Text
            End If

            MyBase.OnTextChanged(e)
    End Sub

#End Region

    Public Sub New()
        JoinEvents(True)
    End Sub
End Class

