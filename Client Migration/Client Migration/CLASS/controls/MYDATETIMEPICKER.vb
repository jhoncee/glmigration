Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms.VisualStyles

Module colors
    Public _BaseColor As Color = Color.FromArgb(45, 47, 49)
    Public _LineColor As Color = Color.FromArgb(25, 27, 29)
    Public _backDisabledColor As Color = Color.Gray
    Public W, H As Integer
    Public _StartIndex As Integer = 0
    Public x, y As Integer


    Public _BGColor As Color = Color.FromArgb(45, 47, 49)
    Public _HoverColor As Color = Color.FromArgb(35, 168, 109)

End Module
Public Class Mydatetimepicker : Inherits DateTimePicker


    Public Sub New()
        SetStyle(ControlStyles.UserPaint, True)
        BackColor = _BaseColor
        ForeColor = _FlatColor
        Height = 22
        Format = DateTimePickerFormat.Short
        Font = New Font("Tahoma", 8, FontStyle.Regular)
    End Sub
    Friend G As Graphics, B As Bitmap

    Friend NearSF As New StringFormat() With {.Alignment = StringAlignment.Near, .LineAlignment = StringAlignment.Near}
    Friend CenterSF As New StringFormat() With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
    Private _BGColor As Color = Color.FromArgb(45, 47, 49)

    <Category("Colors")>
    Public Property HoverColor As Color
        Get
            Return _HoverColor
        End Get
        Set(value As Color)
            _HoverColor = value
        End Set
    End Property

    Protected Overrides Sub OnPaint(e As System.Windows.Forms.PaintEventArgs)
        B = New Bitmap(Width, Height) : G = Graphics.FromImage(B)
        W = Width : H = Height

        Dim Base As New Rectangle(0, 0, W, H)
        Dim Button As New Rectangle(CInt(W - 40), 0, W, H)
        Dim GP, GP2 As New GraphicsPath

        With G
            .Clear(Color.FromArgb(45, 45, 48))
            .SmoothingMode = 2
            .PixelOffsetMode = 2
            .TextRenderingHint = 5

            '-- Base
            .FillRectangle(New SolidBrush(_BGColor), Base)

            '-- Button
            GP.Reset()
            GP.AddRectangle(Button)
            .SetClip(GP)
            .FillRectangle(New SolidBrush(_BaseColor), Button)
            .ResetClip()

            '-- Lines
            .DrawLine(Pens.White, W - 10, 7, W - 30, 7)
            .DrawLine(Pens.White, W - 10, 14, W - 30, 14)
            .DrawLine(Pens.White, W - 10, 21, W - 30, 21)

            '-- Text
            .DrawString(Text, Font, Brushes.White, New Point(4, 6), NearSF)
        End With

        G.Dispose()
        e.Graphics.InterpolationMode = 7
        e.Graphics.DrawImageUnscaled(B, 0, 0)
        B.Dispose()
    End Sub
    <Browsable(True)>
    Public Overrides Property BackColor() As Color
        Get
            Return MyBase.BackColor
        End Get
        Set
            MyBase.BackColor = Value
        End Set
    End Property

End Class

Public Class RANZcombobox : Inherits SergeUtils.EasyCompletionComboBox
    <Browsable(True)>
    Public Overrides Property BackColor() As Color
        Get
            Return MyBase.BackColor
        End Get
        Set
            MyBase.BackColor = Value
        End Set
    End Property
#Region " Variables"

    Private W, H As Integer
    Private _StartIndex As Integer = 0
    Private x, y As Integer

#End Region

#Region " Properties"

#Region " Mouse States"

    Private State As MouseState = MouseState.None
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        State = MouseState.Down : Invalidate()
    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        State = MouseState.Over : Invalidate()
    End Sub
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        State = MouseState.Over : Invalidate()
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        State = MouseState.None : Invalidate()
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        x = e.Location.X
        y = e.Location.Y
        Invalidate()
        If e.X < Width - 41 Then Cursor = Cursors.IBeam Else Cursor = Cursors.Hand
    End Sub

    Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
        MyBase.OnDrawItem(e) : Invalidate()
        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnClick(e As EventArgs)
        MyBase.OnClick(e) : Invalidate()
    End Sub

#End Region

#Region " Colors"

    <Category("Colors")>
    Public Property HoverColor As Color
        Get
            Return _HoverColor
        End Get
        Set(value As Color)
            _HoverColor = value
        End Set
    End Property

#End Region

    Private Property StartIndex As Integer
        Get
            Return _StartIndex
        End Get
        Set(ByVal value As Integer)
            _StartIndex = value
            Try
                MyBase.SelectedIndex = value
            Catch
            End Try
            Invalidate()
        End Set
    End Property



    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        Height = 18
    End Sub

#End Region

#Region " Colors"

    Private _BaseColor As Color = Color.FromArgb(25, 27, 29)
    Private _BGColor As Color = Color.FromArgb(45, 47, 49)
    Private _HoverColor As Color = Color.FromArgb(35, 168, 109)

#End Region

    Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or
                 ControlStyles.ResizeRedraw Or ControlStyles.OptimizedDoubleBuffer, True)
        DoubleBuffered = True

        DrawMode = DrawMode.OwnerDrawFixed
        BackColor = Color.FromArgb(45, 45, 48)
        ForeColor = Color.White
        ' DropDownStyle = ComboBoxStyle.DropDownList

        Cursor = Cursors.Hand
        StartIndex = 0
        ItemHeight = 18
        Font = New Font("Segoe UI", 8, FontStyle.Regular)
    End Sub


End Class
