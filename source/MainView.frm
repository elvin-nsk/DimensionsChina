VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   Caption         =   "Dimensions"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4335
   OleObjectBlob   =   "MainView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const DimensionsText = "Dimensions_Text"
Private Const DimensionsLine = "Dimensions_Line"

Private Sub CheckBox1_Click()
    If CheckBox1 Then CheckBox4 = False
End Sub

Private Sub CheckBox2_Click()
    If CheckBox2 Then CheckBox4 = False
End Sub

Private Sub CheckBox3_Click()
    If CheckBox3 Then CheckBox1 = False
    CheckBox2 = False
    CheckBox4 = False
End Sub

Private Sub CheckBox4_Click()
    If CheckBox4 Then CheckBox1 = False
    CheckBox2 = False
    CheckBox3 = False
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ButtonMovedOut(ApplyButton)
    Call ButtonMovedOut(CancelButton)
End Sub

Private Sub SpinButton1_SpinDown()
    ChangeTextSize ActiveSelection.Shapes.All, -1
End Sub
Private Sub SpinButton1_SpinUp()
    ChangeTextSize ActiveSelection.Shapes.All, 1
End Sub

Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    SetFontSize
End Sub

Private Sub ElvinLogo_Click()
    VBA.CreateObject("WScript.Shell").Run "https://vk.com/elvin_macro"
End Sub

Private Sub VBASomewhat_Click()
    VBA.CreateObject("WScript.Shell").Run "www.cdrvba.com"
End Sub

Private Sub ApplyButton_Click()
    ActiveDocument.BeginCommandGroup "Set dimensions"
    If CheckBox1 Or CheckBox2 Then Call LabelWithWidthAndHeight
    If CheckBox3 Then Call DimLineLength1
    If CheckBox4 Then Call DimLineLength2
    ActiveDocument.EndCommandGroup
End Sub

Private Sub ApplyButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ButtonMovedIn(ApplyButton)
End Sub

Private Sub CancelButton_Click()
    ActiveDocument.BeginCommandGroup "Remove dimensions"
    RemoveDimensions
    ActiveDocument.EndCommandGroup
End Sub

Private Sub CancelButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ButtonMovedIn(CancelButton)
End Sub

Private Sub Kaibein_Click()
    VBA.CreateObject("WScript.Shell").Run "https://www.cdrvba.com/user-register?referralCode=cxx8inoftc4l"
End Sub

Private Sub UserForm_Initialize()
    Tis.BackColor = RGB(40, 170, 20)
    Tis.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub LabelWithWidthAndHeight()
    ActiveDocument.Unit = cdrMillimeter
    Dim s As Shape, st1 As Shape, st2 As Shape
    Set s = ActiveShape
    If s Is Nothing Then Exit Sub
    Optimization = True '优化启动
    For Each s In ActiveSelection.Shapes
        If CheckBox1 Then
            Set st1 = ActiveLayer.CreateArtisticText(s.LeftX, s.TopY + 4, Round(s.SizeWidth, 0) & "mm", , , , TextBox1.Value, , , , cdrCenterAlignment)
                st1.Text.Story.CharSpacing = 0 '字符间距
                st1.Fill.UniformColor.RGBAssign 40, 170, 20 ' 填充颜色
                st1.Move s.SizeWidth / 2, 0
                st1.Name = DimensionsText ' 设置名
            Set sox = ActiveLayer.CreateLineSegment(s.LeftX, s.TopY + 3, s.RightX, s.TopY + 3)
                sox.Outline.Color.RGBAssign 40, 170, 20 ' 填充颜色
                sox.Name = DimensionsLine
            Set sox1 = ActiveLayer.CreateLineSegment(s.LeftX, s.TopY + 1, s.LeftX, s.TopY + 3)
                sox1.Outline.Color.RGBAssign 40, 170, 20 ' 填充颜色
                sox1.Name = DimensionsLine
            Set sox2 = ActiveLayer.CreateLineSegment(s.RightX, s.TopY + 1, s.RightX, s.TopY + 3)
                sox2.Outline.Color.RGBAssign 40, 170, 20 ' 填充颜色
                sox2.Name = DimensionsLine
            s.CreateSelection
        End If
        If CheckBox2 Then
            Set st2 = ActiveLayer.CreateArtisticText(s.LeftX - 4, s.BottomY, Round(s.SizeHeight, 0) & "mm", , , , TextBox1.Value, , , , cdrCenterAlignment)
            st2.Text.Story.CharSpacing = 0 '字符间距
            st2.Fill.UniformColor.RGBAssign 40, 170, 20 ' 填充颜色
            st2.Rotate 90
            st2.Move -st2.SizeWidth / 2, s.SizeHeight / 2
            st2.Name = DimensionsText ' 设置名
            Set soy = ActiveLayer.CreateLineSegment(s.LeftX - 3, s.BottomY, s.LeftX - 3, s.TopY)
                soy.Outline.Color.RGBAssign 40, 170, 20 ' 填充颜色
                soy.Name = DimensionsLine
            Set soy1 = ActiveLayer.CreateLineSegment(s.LeftX - 1, s.BottomY, s.LeftX - 3, s.BottomY)
                soy1.Outline.Color.RGBAssign 40, 170, 20 ' 填充颜色
                soy1.Name = DimensionsLine
            Set soy2 = ActiveLayer.CreateLineSegment(s.LeftX - 1, s.TopY, s.LeftX - 3, s.TopY)
                soy2.Outline.Color.RGBAssign 40, 170, 20 ' 填充颜色
                soy2.Name = DimensionsLine
            s.CreateSelection
        End If
    Next
    Optimization = False '优化关闭
    ActiveWindow.Refresh '刷新文档窗口
End Sub

Private Sub DimLineLength2()
    ActiveDocument.Unit = cdrMillimeter
    Dim s As Shape, s1 As Shape, s2 As Shape, sc As Shape, st1 As Shape, st2 As Shape
    Set s = ActiveShape
    If s Is Nothing Then Exit Sub
    Optimization = True '优化启动
    For Each s In ActiveSelection.Shapes
        If s.Type <> cdrTextShape Then
            s.Copy
            Set sc = ActiveLayer.Paste
            sc.ConvertToCurves
            sc.Curve.Nodes.All.BreakApart
            sc.BreakApart
            For Each s1 In ActiveSelection.Shapes
                Set st1 = ActiveLayer.CreateArtisticText(0, 0, Round(s1.Curve.Length, 0), , , , TextBox1.Value)
                st1.Text.Story.CharSpacing = 0 '字符间距
                st1.Fill.UniformColor.RGBAssign 40, 170, 20 ' 填充颜色
                st1.Text.FitToPath s1
                ' 获取或设置文本与路径的偏移量
                st1.Effects(1).TextOnPath.Offset = s1.Curve.Length * 0.5 - st1.SizeWidth * 0.55
                ' 获取或设置文本与路径的距离
                st1.Effects(1).TextOnPath.DistanceFromPath = 1
                st1.Name = DimensionsText ' 设置名
                s1.Outline.SetNoOutline
                s1.OrderToBack
                s1.Name = DimensionsLine
            Next
            Set st2 = ActiveLayer.CreateArtisticText(s.RightX + 3, s.BottomY, "Units: mm", , , , TextBox1.Value)
            st2.Text.Story.CharSpacing = 0 '字符间距
            st2.Fill.UniformColor.RGBAssign 40, 170, 20 ' 填充颜色
            st2.Name = DimensionsText ' 设置名
         End If
    Next
    Optimization = False '优化关闭
    ActiveWindow.Refresh '刷新文档窗口
End Sub
Private Sub DimLineLength1()
    ActiveDocument.Unit = cdrMillimeter
    Dim s As Shape, st1 As Shape
    Set s = ActiveShape
    If s Is Nothing Then Exit Sub
    Optimization = True '优化启动
    For Each s In ActiveSelection.Shapes
        If s.Type <> cdrTextShape Then
            X = s.LeftX
            Y = s.BottomY
            Set st1 = ActiveLayer.CreateArtisticText(X, Y, "Perimeter: " & Round(s.DisplayCurve.Length, 0) & "mm", , , , TextBox1.Value, , , , cdrLeftAlignment)
            st1.Text.Story.CharSpacing = 0 '字符间距
            st1.Fill.UniformColor.RGBAssign 40, 170, 20 ' 填充颜色
            st1.Move 0, -st1.SizeHeight * 2
            st1.Name = DimensionsText ' 设置名
            s.CreateSelection
        End If
    Next
    Optimization = False '优化关闭
    ActiveWindow.Refresh '刷新文档窗口
End Sub

Private Sub ChangeTextSize( _
                ByVal ShapesToSearch As ShapeRange, _
                ByVal SizeChange As Long _
            )
    Dim s As Shape
    Optimization = True '优化启动
    If TextBox1.Value > 0 Then
        TextBox1.Value = TextBox1.Value + SizeChange
        For Each s In ShapesToSearch.Shapes.FindShapes(Name:=DimensionsText)
            s.Text.Story.Size = s.Text.Story.Size + SizeChange
        Next
    End If
    Optimization = False '优化关闭
    ActiveWindow.Refresh '刷新文档窗口
End Sub

Private Sub SetFontSize()
    Dim s As Shape
    Optimization = True '优化启动
    If TextBox1.Value > 0 Then
        For Each s In ActiveSelection.Shapes.FindShapes(Name:=DimensionsText)
            s.Text.Story.Size = TextBox1.Value
        Next
    End If
    Optimization = False '优化关闭
    ActiveWindow.Refresh '刷新文档窗口
End Sub
Private Sub RemoveDimensions()
    If ActiveSelection.Shapes.Count > 0 Then
        ActiveSelection.Shapes.FindShapes(Name:=DimensionsText).Delete
        ActiveSelection.Shapes.FindShapes(Name:=DimensionsLine).Delete
    Else
        ActivePage.Shapes.FindShapes(Name:=DimensionsText).Delete
        ActivePage.Shapes.FindShapes(Name:=DimensionsLine).Delete
    End If
End Sub
