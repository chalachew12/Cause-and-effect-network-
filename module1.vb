Public tt As Integer
Public nn As Integer

Private Sub CommandButton2_Click()
Dim i As Integer
Dim n As String
Dim m As String
Dim S As String
For i = Sheets("Diagram2").Shapes.count To 1 Step -1
If StrComp(Sheets("Diagram2").Shapes.Item(i).Name, "back1") And StrComp(Sheets("Diagram2").Shapes.Item(i).Name, "back") And StrComp(Sheets("Diagram2").Shapes.Item(i).Name, "pri") And StrComp(Sheets("Diagram2").Shapes.Item(i).Name, "back") And StrComp(Sheets("Diagram2").Shapes.Item(i).Name, "header") And StrComp(Sheets("Diagram2").Shapes.Item(i).Name, "group") And StrComp(Sheets("Diagram2").Shapes.Item(i).Name, "Im") Then
  Sheets("Diagram2").Shapes(i).Delete
  End If
Next
For i = Sheets("direction-diagram").Shapes.count To 1 Step -1
  Sheets("direction-diagram").Shapes(i).Delete
Next
For i = Sheets("strength-diagram").Shapes.count To 1 Step -1
  Sheets("strength-diagram").Shapes(i).Delete
Next

For i = Sheets("effect-strength-diagram").Shapes.count To 1 Step -1
  Sheets("effect-strength-diagram").Shapes(i).Delete
Next
End Sub

Private Sub CommandButton1_Click()
Application.ScreenUpdating = False
Dim inpute As Range
Dim Head As Range
Dim header As Range
Dim i As Integer, j As Integer, k As Integer, n As Integer, m As Integer
Dim t As Integer
Dim z As Integer
Dim count As Integer
Dim effect As String
Dim effect2 As String
Dim effect1 As String
Dim effect3 As String
Dim effect4 As String
Dim effect5 As String
Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim x As Integer
Dim y As Integer
Dim pTwo As Integer
Dim y2 As Double
Dim per As Integer
Dim per2 As Integer
Dim shp As Shape
Dim twostep1 As Double
Dim twostep2 As Double
Dim pos1
Dim pos2
Dim pos3
Dim p1 As Integer
Dim p2 As Integer
Dim p3 As Integer
Dim count2 As Double
Dim totper As Integer
Dim p As Integer
Dim count3 As Double
Dim perspectivestart As Integer
Dim perspectiveend As Integer
Dim cx As Integer
Dim S As Shape
Dim b As Integer
Dim xx As Integer
Dim cc As Integer
Dim dd As Integer
Dim hh As String
Dim mm As Integer
Dim ff As Integer
Dim mmm As Long

Set Matrix = Range("G6:DO118")
Set header = Range("A6:E118")

pos1 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
pos2 = Array(255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 128, 192, 128, 153, 51, 255, 255, 0, 128, 102, 204, 0, 0, 255, 255, 0, 0, 128, 0, 204, 255, 255, 255, 204, 153, 153, 204, 102, 204, 204, 204, 153, 102, 102, 150, 51, 153, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255)
pos3 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

n = header.Rows.count
m = Matrix.Columns.count

t = 1000
z = 600
count = 2
totper = 1
For i = 1 To m
  
  effect1 = header(i, 2).Value
  effect2 = header(i, 3).Value
  effect3 = header(i, 1).Value
  effect4 = header(i, 4).Value
  effect5 = header(i, 5).Value
  effect = effect1 + vbNewLine + effect2 + effect4 + vbNewLine + effect5
  If totper = 7 Then
         If count = 2 Then
          Call CreateAutoshapes4(t - 200, z, effect3)
         End If
        
         Call CreateAutoshapes(t, z, effect, i)
         t = t + 500
         count = count + 1
  ElseIf totper = 6 Then
     If count > 16 Then
         count = 2
         t = 1000
         z = z + 500
         Call CreateAutoshapes4(t - 200, z, effect3)
         
         t = t + 300
         Call CreateAutoshapes(t, z, effect, i)
         t = t + 500
         count = count + 1
         totper = totper + 1
      Else
         If count = 2 Then
          Call CreateAutoshapes4(t - 200, z, effect3)
         End If
         Call CreateAutoshapes(t, z, effect, i)
         t = t + 180
         count = count + 1
      End If
  ElseIf totper = 5 Then
     If count > 16 Then
         count = 2
         t = 1000
         z = z + 500
         Call CreateAutoshapes4(t - 200, z, effect3)
         Call CreateAutoshapes(t, z, effect, i)
         t = t + 180
         count = count + 1
         totper = totper + 1
      Else
         If count = 2 Then
          Call CreateAutoshapes4(t - 200, z, effect3)
         End If
         Call CreateAutoshapes(t, z, effect, i)
         t = t + 180
         count = count + 1
      End If
  ElseIf totper = 4 Then
     If count > 25 Then
         count = 2
         t = 1000
         z = z + 500
         Call CreateAutoshapes4(t - 200, z, effect3)
         Call CreateAutoshapes(t, z, effect, i)
         t = t + 180
         count = count + 1
         totper = totper + 1
     Else
        Call CreateAutoshapes(t, z, effect, i)
        t = t + 180
        count = count + 1
     End If
  ElseIf totper < 4 Then
     If count > 19 Then
        count = 2
        t = 1000
        z = z + 500
        Call CreateAutoshapes4(t - 200, z, effect3)
        Call CreateAutoshapes(t, z, effect, i)
        t = t + 180
        count = count + 1
        totper = totper + 1
     Else
        If count = 2 Then
          Call CreateAutoshapes4(t - 200, z, effect3)
        End If
           Call CreateAutoshapes(t, z, effect, i)
           t = t + 180
           count = count + 1
     End If
  End If
Next i
For cc = 1 To m
  For dd = 1 To m
    If Matrix(cc, dd).Value = 1 Or Matrix(dd, cc) = 1 Then
        cx = cx + 1
    End If
  Next dd
  If cx = 0 Then
     For b = Sheets("Diagram2").Shapes.count To 1 Step -1
         If Sheets("Diagram2").Shapes.Item(b).Name = cc Then
             Sheets("Diagram2").Shapes(b).Delete
         End If
     Next
  End If
  cx = 0
Next cc


y2 = 0.96
y = 100
pTwo = 100
x = 1
per = 1
per2 = 1
twostep1 = -0.5
twostep2 = 1
count2 = 1
count3 = 1
perspectivestart = 1
perspectiveend = 1
cx = 0
b = 0
mm = 0
nn = 0
tt = 0

Dim ttt
Dim ttt1
Dim vv As Integer
Dim zz As String
Dim ww As String
Dim shi As Shape
Dim shj As Shape
Dim bb As Integer
Dim bbb As Integer
Dim t3 As Integer
Dim rr As String
Dim t4 As Integer
Dim vvv As Integer
Dim fff As Integer
Dim IJ As String
Dim str As Integer
fff = 1
ff = 1
t3 = 0
ttt = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
ttt1 = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
For i = 1 To m
         For bb = Sheets("Diagram2").Shapes.count To 1 Step -1
           If Sheets("Diagram2").Shapes.Item(bb).Name = i Then
             Set shi = Sheets("Diagram2").Shapes.Item(bb)
           End If
         Next
    For j = 1 To m
      If Matrix(i, j).Value = 1 Then
         IJ = i & j
         rr = Matrix(i, j).Value
         t4 = ttt1(i - 1) + 20
         ttt1(i - 1) = t4
         For bbb = Sheets("Diagram2").Shapes.count To 1 Step -1
           If Sheets("Diagram2").Shapes.Item(bbb).Name = j Then
             Set shj = Sheets("Diagram2").Shapes.Item(bbb)
           End If
         Next
         vv = shj.Fill.ForeColor.SchemeColor
         vvv = shi.Fill.ForeColor.SchemeColor
         zz = shj.TextFrame.Characters.Text
        
         If (vv = 6) Then
            ff = 0
         Else
            ff = 1
         End If
           
         p1 = 0
         p2 = 255
         p3 = 0
         If per < 19 Then
            y = y + 10
            perspectiveend = 1
         ElseIf per > 18 And per < 37 Then
            y2 = y2 - 0.01
            perspectiveend = 2
         ElseIf per > 36 And per < 55 Then
           perspectiveend = 3
           twostep1 = twostep1 - 0.001
           twostep2 = twostep2 + 0.001
           count3 = count + 0.005
         ElseIf per > 54 And per < 79 Then
           perspectiveend = 4
           twostep1 = twostep1 - 0.001
           twostep2 = twostep2 + 0.001
           count3 = count + 0.005
         ElseIf per > 78 And per < 94 Then
           perspectiveend = 5
           twostep1 = twostep1 - 0.001
           twostep2 = twostep2 + 0.001
           count3 = count + 0.005
         ElseIf per > 93 And per < 109 Then
           perspectiveend = 6
           twostep1 = twostep1 - 0.001
           twostep2 = twostep2 + 0.001
           count3 = count + 0.005
         ElseIf per > 108 And per < 114 Then
           perspectiveend = 7
           twostep1 = twostep1 - 0.001
           twostep2 = twostep2 + 0.001
           count3 = count + 0.005
         End If
         If per2 > 18 Then
            pTwo = pTwo + 10
         End If
         AddConnectorBetweenShapes3 msoConnectorStraight, shi, shj, p1, x, y, p2, p3, per, per2, y2, pTwo, twostep1, twostep2, count2, count3, perspectivestart, perspectiveend, mm, ff, t3, rr, t4, fff, IJ
         x = x + 1
         If x > 8 Then
           x = 1
         End If
     End If
     per = per + 1
    Next j

    x = 1
    per = 1
    per2 = per2 + 1
    twostep2 = twostep2 - 0.001
          
         If per2 < 19 Then
              perspectivestart = 1
         ElseIf per2 > 18 And per2 < 36 Then
              perspectivestart = 2
         ElseIf per2 > 36 And per2 < 54 Then
              perspectivestart = 3
         ElseIf per2 > 54 And per2 < 78 Then
              perspectivestart = 4
         ElseIf per2 > 78 And per2 < 93 Then
              perspectivestart = 5
         ElseIf per2 > 93 And per2 < 108 Then
              perspectivestart = 6
         ElseIf per2 > 108 And per2 < 113 Then
              perspectivestart = 7
         End If
mm = mm + 1
Next i

         
'for the direction worksheet
Sheets("Diagram2").Activate
Sheets("Diagram2").Shapes.SelectAll
Selection.ShapeRange.Group.Name = "Group99"
Sheets("Diagram2").Shapes("Group99").Copy
Sheets("Diagram2").Shapes("Group99").Copy
Sheets("direction-diagram").Paste

Dim mat As Range
Dim Hd As Range
Dim wr As Worksheet
Dim ik As String
Dim Sh As Shape
Dim group_sh As Shape

Sheets("direction-diagram").Activate
For i = Sheets("direction-diagram").Shapes.count To 1 Step -1
  Sheets("direction-diagram").Shapes(i).Ungroup
Next

For i = 1 To m
    For j = 1 To m
         ik = i & j
         Set wr = Worksheets("3-Direction-Matrix")
         Set mat = wr.Range("G6:DO118")
         Set Hd = wr.Range("A6:E118")
         
         If mat(i, j).Value < 0 Then
         
           t4 = ttt1(i - 1) + 20
           ttt1(i - 1) = t4
           For bbb = Sheets("direction-diagram").Shapes.count To 1 Step -1
             If Sheets("direction-diagram").Shapes.Item(bbb).Name = ik Then
               Set Sh = Sheets("direction-diagram").Shapes.Item(bbb)
               Sh.Line.ForeColor.RGB = RGB(255, 0, 0)
             End If
           Next
           
          ElseIf mat(i, j).Value <> 0 Then
            For bbb = Sheets("direction-diagram").Shapes.count To 1 Step -1
             If Sheets("direction-diagram").Shapes.Item(bbb).Name = ik Then
               Set Sh = Sheets("direction-diagram").Shapes.Item(bbb)
               Sh.Line.ForeColor.RGB = RGB(0, 225, 0)
             End If
            Next
          End If
    Next j
Next i
 
'for relationship strength
Dim mat1 As Range
Dim Hd1 As Range
Dim wr1 As Worksheet
Dim yy As String
Dim distance_bet_ovsal As Double
distance_bet_ovsal = 0
Sheets("direction-diagram").Shapes.SelectAll
Selection.ShapeRange.Group.Name = "Group99"
Sheets("direction-diagram").Shapes("Group99").Copy
Sheets("strength-diagram").Paste
Sheets("strength-diagram").Activate
For i = Sheets("strength-diagram").Shapes.count To 1 Step -1
  Sheets("strength-diagram").Shapes(i).Ungroup
Next

For i = 1 To m
         For bb = Sheets("strength-diagram").Shapes.count To 1 Step -1
           If Sheets("strength-diagram").Shapes.Item(bb).Name = i Then
             Set shi = Sheets("strength-diagram").Shapes.Item(bb)
           End If
         Next
    For j = 1 To m

         Set wr1 = Worksheets("4-Relationship-Strength-Matrix")
         Set mat1 = wr1.Range("G6:DO118")
         Set Hd1 = wr1.Range("A6:E118")
         
         If mat1(i, j).Value <> 0 Then
             yy = mat1(i, j).Value
             If yy = "-3" Then
                t4 = ttt1(i - 1) + 20
                ttt1(i - 1) = t4
                For bbb = Sheets("strength-diagram").Shapes.count To 1 Step -1
                  If Sheets("strength-diagram").Shapes.Item(bbb).Name = j Then
                    Set shj = Sheets("strength-diagram").Shapes.Item(bbb)
                  End If
                Next
                
                p1 = 0
                p2 = 255
                p3 = 0
                
                If per < 19 Then
                   y = y + 10
                   perspectiveend = 1
                ElseIf per > 18 And per < 36 Then
                   y2 = y2 - 0.01
                   perspectiveend = 2
                ElseIf per > 37 And per < 54 Then
                  perspectiveend = 3
                  twostep1 = twostep1 - 0.001
                  twostep2 = twostep2 + 0.001
                  count3 = count + 0.005
                ElseIf per > 55 And per < 78 Then
                  perspectiveend = 4
                  twostep1 = twostep1 - 0.001
                  twostep2 = twostep2 + 0.001
                  count3 = count + 0.005
                ElseIf per > 79 And per < 93 Then
                  perspectiveend = 5
                  twostep1 = twostep1 - 0.001
                  twostep2 = twostep2 + 0.001
                  count3 = count + 0.005
                ElseIf per > 94 And per < 108 Then
                  perspectiveend = 6
                  twostep1 = twostep1 - 0.001
                  twostep2 = twostep2 + 0.001
                  count3 = count + 0.005
                ElseIf per > 109 And per < 113 Then
                  perspectiveend = 7
                  twostep1 = twostep1 - 0.001
                  twostep2 = twostep2 + 0.001
                  count3 = count + 0.005
                End If
                If per2 > 18 Then
                   pTwo = pTwo + 10
                End If
                   AddConnectorBetweenShapes4 msoConnectorStraight, shi, shj, p1, x, y, p2, p3, per, per2, y2, pTwo, twostep1, twostep2, count2, count3, perspectivestart, perspectiveend, mm, ff, t3, -3, t4, fff, distance_bet_ovsal
             ElseIf yy = "-2" Then
                    t4 = ttt1(i - 1) + 20
                    ttt1(i - 1) = t4
                    
                    For bbb = Sheets("strength-diagram").Shapes.count To 1 Step -1
                      If Sheets("strength-diagram").Shapes.Item(bbb).Name = j Then
                        Set shj = Sheets("strength-diagram").Shapes.Item(bbb)
                      End If
                    Next
                    
                    p1 = 0
                    p2 = 255
                    p3 = 0
                    
                    If per < 19 Then
                       y = y + 10
                       perspectiveend = 1
                    ElseIf per > 18 And per < 36 Then
                       y2 = y2 - 0.01
                       perspectiveend = 2
                    ElseIf per > 37 And per < 54 Then
                      perspectiveend = 3
                      twostep1 = twostep1 - 0.001
                      twostep2 = twostep2 + 0.001
                      count3 = count + 0.005
                    ElseIf per > 55 And per < 78 Then
                      perspectiveend = 4
                      twostep1 = twostep1 - 0.001
                      twostep2 = twostep2 + 0.001
                      count3 = count + 0.005
                    ElseIf per > 79 And per < 93 Then
                      perspectiveend = 5
                      twostep1 = twostep1 - 0.001
                      twostep2 = twostep2 + 0.001
                      count3 = count + 0.005
                    ElseIf per > 94 And per < 108 Then
                      perspectiveend = 6
                      twostep1 = twostep1 - 0.001
                      twostep2 = twostep2 + 0.001
                      count3 = count + 0.005
                    ElseIf per > 109 And per < 113 Then
                      perspectiveend = 7
                      twostep1 = twostep1 - 0.001
                      twostep2 = twostep2 + 0.001
                      count3 = count + 0.005
                    End If
                    If per2 > 18 Then
                       pTwo = pTwo + 10
                    End If
                       AddConnectorBetweenShapes4 msoConnectorStraight, shi, shj, p1, x, y, p2, p3, per, per2, y2, pTwo, twostep1, twostep2, count2, count3, perspectivestart, perspectiveend, mm, ff, t3, -2, t4, fff, distance_bet_ovsal
                ElseIf yy = "-1" Then
                   t4 = ttt1(i - 1) + 20
                   ttt1(i - 1) = t4
                   
                   For bbb = Sheets("strength-diagram").Shapes.count To 1 Step -1
                     If Sheets("strength-diagram").Shapes.Item(bbb).Name = j Then
                       Set shj = Sheets("strength-diagram").Shapes.Item(bbb)
                     End If
                   Next
                   
                   p1 = 0
                   p2 = 255
                   p3 = 0
                   
                   If per < 19 Then
                      y = y + 10
                      perspectiveend = 1
                   ElseIf per > 18 And per < 36 Then
                      y2 = y2 - 0.01
                      perspectiveend = 2
                   ElseIf per > 37 And per < 54 Then
                     perspectiveend = 3
                     twostep1 = twostep1 - 0.001
                     twostep2 = twostep2 + 0.001
                     count3 = count + 0.005
                   ElseIf per > 55 And per < 78 Then
                     perspectiveend = 4
                     twostep1 = twostep1 - 0.001
                     twostep2 = twostep2 + 0.001
                     count3 = count + 0.005
                   ElseIf per > 79 And per < 93 Then
                     perspectiveend = 5
                     twostep1 = twostep1 - 0.001
                     twostep2 = twostep2 + 0.001
                     count3 = count + 0.005
                   ElseIf per > 94 And per < 108 Then
                     perspectiveend = 6
                     twostep1 = twostep1 - 0.001
                     twostep2 = twostep2 + 0.001
                     count3 = count + 0.005
                   ElseIf per > 109 And per < 113 Then
                     perspectiveend = 7
                     twostep1 = twostep1 - 0.001
                     twostep2 = twostep2 + 0.001
                     count3 = count + 0.005
                   End If
                   If per2 > 18 Then
                      pTwo = pTwo + 10
                   End If
                     AddConnectorBetweenShapes4 msoConnectorStraight, shi, shj, p1, x, y, p2, p3, per, per2, y2, pTwo, twostep1, twostep2, count2, count3, perspectivestart, perspectiveend, mm, ff, t3, -1, t4, fff, distance_bet_ovsal
            ElseIf yy = "1" Then
                
                yy = mat1(i, j).Value
                t4 = ttt1(i - 1) + 20
                ttt1(i - 1) = t4
                
                For bbb = Sheets("strength-diagram").Shapes.count To 1 Step -1
                  If Sheets("strength-diagram").Shapes.Item(bbb).Name = j Then
                    Set shj = Sheets("strength-diagram").Shapes.Item(bbb)
                  End If
                Next
                
                p1 = 0
                p2 = 255
                p3 = 0
                
                If per < 19 Then
                   y = y + 10
                   perspectiveend = 1
                ElseIf per > 18 And per < 36 Then
                   y2 = y2 - 0.01
                   perspectiveend = 2
                ElseIf per > 37 And per < 54 Then
                  perspectiveend = 3
                  twostep1 = twostep1 - 0.001
                  twostep2 = twostep2 + 0.001
                  count3 = count + 0.005
                ElseIf per > 55 And per < 78 Then
                  perspectiveend = 4
                  twostep1 = twostep1 - 0.001
                  twostep2 = twostep2 + 0.001
                  count3 = count + 0.005
                ElseIf per > 79 And per < 93 Then
                  perspectiveend = 5
                  twostep1 = twostep1 - 0.001
                  twostep2 = twostep2 + 0.001
                  count3 = count + 0.005
                ElseIf per > 94 And per < 108 Then
                  perspectiveend = 6
                  twostep1 = twostep1 - 0.001
                  twostep2 = twostep2 + 0.001
                  count3 = count + 0.005
                ElseIf per > 109 And per < 113 Then
                  perspectiveend = 7
                  twostep1 = twostep1 - 0.001
                  twostep2 = twostep2 + 0.001
                  count3 = count + 0.005
                End If
                If per2 > 18 Then
                   pTwo = pTwo + 10
                End If
                   AddConnectorBetweenShapes4 msoConnectorStraight, shi, shj, p1, x, y, p2, p3, per, per2, y2, pTwo, twostep1, twostep2, count2, count3, perspectivestart, perspectiveend, mm, ff, t3, 1, t4, fff, distance_bet_ovsal
             ElseIf yy = "2" Then
                 t4 = ttt1(i - 1) + 20
                 ttt1(i - 1) = t4
                 
                 For bbb = Sheets("strength-diagram").Shapes.count To 1 Step -1
                   If Sheets("strength-diagram").Shapes.Item(bbb).Name = j Then
                     Set shj = Sheets("strength-diagram").Shapes.Item(bbb)
                   End If
                 Next
                 
                 p1 = 0
                 p2 = 255
                 p3 = 0
                 
                 If per < 19 Then
                    y = y + 10
                    perspectiveend = 1
                 ElseIf per > 18 And per < 36 Then
                    y2 = y2 - 0.01
                    perspectiveend = 2
                 ElseIf per > 37 And per < 54 Then
                   perspectiveend = 3
                   twostep1 = twostep1 - 0.001
                   twostep2 = twostep2 + 0.001
                   count3 = count + 0.005
                 ElseIf per > 55 And per < 78 Then
                   perspectiveend = 4
                   twostep1 = twostep1 - 0.001
                   twostep2 = twostep2 + 0.001
                   count3 = count + 0.005
                 ElseIf per > 79 And per < 93 Then
                   perspectiveend = 5
                   twostep1 = twostep1 - 0.001
                   twostep2 = twostep2 + 0.001
                   count3 = count + 0.005
                 ElseIf per > 94 And per < 108 Then
                   perspectiveend = 6
                   twostep1 = twostep1 - 0.001
                   twostep2 = twostep2 + 0.001
                   count3 = count + 0.005
                 ElseIf per > 109 And per < 113 Then
                   perspectiveend = 7
                   twostep1 = twostep1 - 0.001
                   twostep2 = twostep2 + 0.001
                   count3 = count + 0.005
                 End If
                 If per2 > 18 Then
                    pTwo = pTwo + 10
                 End If
                    AddConnectorBetweenShapes4 msoConnectorStraight, shi, shj, p1, x, y, p2, p3, per, per2, y2, pTwo, twostep1, twostep2, count2, count3, perspectivestart, perspectiveend, mm, ff, t3, 2, t4, fff, distance_bet_ovsal
             ElseIf yy = "3" Then
                 t4 = ttt1(i - 1) + 20
                 ttt1(i - 1) = t4
                 
                 For bbb = Sheets("strength-diagram").Shapes.count To 1 Step -1
                   If Sheets("strength-diagram").Shapes.Item(bbb).Name = j Then
                     Set shj = Sheets("strength-diagram").Shapes.Item(bbb)
                   End If
                 Next
                 
                 p1 = 0
                 p2 = 255
                 p3 = 0
                 
                 If per < 19 Then
                    y = y + 10
                    perspectiveend = 1
                 ElseIf per > 18 And per < 36 Then
                    y2 = y2 - 0.01
                    perspectiveend = 2
                 ElseIf per > 37 And per < 54 Then
                   perspectiveend = 3
                   twostep1 = twostep1 - 0.001
                   twostep2 = twostep2 + 0.001
                   count3 = count + 0.005
                 ElseIf per > 55 And per < 78 Then
                   perspectiveend = 4
                   twostep1 = twostep1 - 0.001
                   twostep2 = twostep2 + 0.001
                   count3 = count + 0.005
                 ElseIf per > 79 And per < 93 Then
                   perspectiveend = 5
                   twostep1 = twostep1 - 0.001
                   twostep2 = twostep2 + 0.001
                   count3 = count + 0.005
                 ElseIf per > 94 And per < 108 Then
                   perspectiveend = 6
                   twostep1 = twostep1 - 0.001
                   twostep2 = twostep2 + 0.001
                   count3 = count + 0.005
                 ElseIf per > 109 And per < 113 Then
                   perspectiveend = 7
                   twostep1 = twostep1 - 0.001
                   twostep2 = twostep2 + 0.001
                   count3 = count + 0.005
                 End If
                 If per2 > 18 Then
                    pTwo = pTwo + 10
                 End If
                    AddConnectorBetweenShapes4 msoConnectorStraight, shi, shj, p1, x, y, p2, p3, per, per2, y2, pTwo, twostep1, twostep2, count2, count3, perspectivestart, perspectiveend, mm, ff, t3, 3, t4, fff, distance_bet_ovsal
                 End If
                 x = x + 1
                 If x > 8 Then
                   x = 1
                 End If
'           ElseIf yy > 3 And yy - 3 Then
'           Call error_handling_r_strength

      End If
      per = per + 1
      distance_bet_ovsal = distance_bet_ovsal + 0.5
    Next j
mm = mm + 1
distance_bet_ovsal = 0
Next i

'for effect strength diagram

Sheets("strength-diagram").Shapes.SelectAll
Selection.ShapeRange.Group.Name = "Group99"
Sheets("strength-diagram").Shapes("Group99").Copy
Sheets("effect-strength-diagram").Paste
Sheets("effect-strength-diagram").Activate
For i = Sheets("effect-strength-diagram").Shapes.count To 1 Step -1
  Sheets("effect-strength-diagram").Shapes(i).Ungroup
Next
Dim es As Range
Dim es_count As Integer
Dim es_sh_count As Integer
Dim esshp As Shape
Dim es_work As Worksheet
Dim esx As Double
Dim esy As Double
Dim es_r As Integer

Set es_work = Worksheets("5-Effect-Strength-Matrix")
Set es = es_work.Range("G119:DO119")
es_count = es.count

For i = 1 To es_count
         For es_sh_count = Sheets("effect-strength-diagram").Shapes.count To 1 Step -1
           If Sheets("effect-strength-diagram").Shapes.Item(es_sh_count).Name = i Then
             Set esshp = Sheets("effect-strength-diagram").Shapes.Item(es_sh_count)
           End If
         Next
         
            If es(i).Value > 0 Then
              For est = 1 To es(i).Value
                es_r = 1
                esshp.Fill.ForeColor.SchemeColor = 3
                  esx = esshp.Left + 10 - t3
                  esy = esshp.Top - 20
                Call CreateAutoshapes1(esy, esx, es_r)
                t3 = t3 - 20
               Next est
               t3 = 0
            ElseIf es(i) < 0 Then
              For est = es(i).Value To -1
                es_r = 0
                esshp.Fill.ForeColor.SchemeColor = 2
                esx = esshp.Left + 10 - t3
                esy = esshp.Top - 20
                Call CreateAutoshapes1(esy, esx, es_r)
                t3 = t3 - 20
              Next est
               t3 = 0
            End If
Next i




Dim shp1 As Shape
Dim shp_next As Shape
Dim shp_prev As Shape
Dim bbs As Integer
Dim bbbs As Integer
Dim description As String
Dim description2 As String
Dim worksheet_des As Worksheet
Dim desstr As String
Dim effect_strength As Integer

Set worksheet_des = Sheets("2 Relationship Matrix")
Set matrix1 = worksheet_des.Range("G6:DO118")
Set header1 = worksheet_des.Range("A6:E118")

For i = 1 To m
'    Set shp1 = Sheets("effect-strength-diagram").Shapes.Item(i)
   description = "Has an effect on: "
   description2 = "Caused by: "
   
   For bbs = Sheets("effect-strength-diagram").Shapes.count To 1 Step -1
         If Sheets("effect-strength-diagram").Shapes.Item(bbs).Name = i Then
           Set shp1 = Sheets("effect-strength-diagram").Shapes.Item(bbs)
           For j = 1 To m
             If matrix1(j, i).Value = 1 Then
                 For bbbs = Sheets("effect-strength-diagram").Shapes.count To 1 Step -1
                    If Sheets("effect-strength-diagram").Shapes.Item(bbbs).Name = j Then
                         Set shp_prev = Sheets("effect-strength-diagram").Shapes.Item(bbbs)
                        
                         description2 = description2 + header(j, 3).Value + header(j, 4).Value
                         
                         If mat(j, i).Value < 0 Then
                            description2 = description2 + "(Negative cause)"
                            
                            If mat1(j, i).Value = -3 Then
                                  description2 = description2 + "relation strength = -3" + vbNewLine
                            ElseIf mat1(j, i).Value = -2 Then
                                  description2 = description2 + "relation strength = -2" + vbNewLine
                            ElseIf mat1(j, i).Value = -1 Then
                                  description2 = description2 + "relation strength = -1" + vbNewLine
                            ElseIf mat1(j, i).Value = 1 Then
                                  description2 = description2 + "relation strength = 1" + vbNewLine
                            ElseIf mat1(j, i).Value = 2 Then
                                  description2 = description2 + "relation strength = 2" + vbNewLine
                            ElseIf mat1(j, i).Value = 3 Then
                                  description2 = description2 + "relation strength = 3" + vbNewLine
                            End If
                            
                         ElseIf mat(j, i).Value > 0 Then
                            description2 = description2 + "(Positive cause)"
                            If mat1(j, i).Value = -3 Then
                                  description2 = description2 + "relation strength = -3" + vbNewLine
                            ElseIf mat1(j, i).Value = -2 Then
                                  description2 = description2 + "relation strength = -2" + vbNewLine
                            ElseIf mat1(j, i).Value = -1 Then
                                  description2 = description2 + "relation strength = -1" + vbNewLine
                            ElseIf mat1(j, i).Value = 1 Then
                                  description2 = description2 + "relation strength = 1" + vbNewLine
                            ElseIf mat1(j, i).Value = 2 Then
                                  description2 = description2 + "relation strength = 2" + vbNewLine
                            ElseIf mat1(j, i).Value = 3 Then
                                  description2 = description2 + "relation strength = 3" + vbNewLine
                            End If
                         
                         End If
                    End If
                 Next
             End If
           Next j
           
           For j = 1 To m
             If matrix1(i, j).Value = 1 Then
                 For bbbs = Sheets("effect-strength-diagram").Shapes.count To 1 Step -1
                    If Sheets("effect-strength-diagram").Shapes.Item(bbbs).Name = j Then
                         Set shp_next = Sheets("effect-strength-diagram").Shapes.Item(bbbs)
                         description = description + " " + header(j, 3).Value + header(j, 4).Value + ","
                         
                         If mat(i, j).Value < 0 Then
                            description = description + "(Negative effect)"
                            
                            If mat1(i, j).Value = -3 Then
                                  description = description + "relation strength = -3" + vbNewLine
                            ElseIf mat1(i, j).Value = -2 Then
                                  description = description + "relation strength = -2" + vbNewLine
                            ElseIf mat1(i, j).Value = -1 Then
                                  description = description + "relation strength = -1" + vbNewLine
                            ElseIf mat1(i, j).Value = 1 Then
                                  description = description + "relation strength = 1" + vbNewLine
                            ElseIf mat1(i, j).Value = 2 Then
                                  description = description + "relation strength = 2" + vbNewLine
                            ElseIf mat1(i, j).Value = 3 Then
                                  description = description + "relation strength = 3" + vbNewLine
                            End If
                         End If
                         If mat(i, j).Value > 0 Then
                            description = description + "(Posetive effect)"
                            
                            If mat1(i, j).Value = -3 Then
                                  description = description + "relation strength = -3" + vbNewLine
                            ElseIf mat1(i, j).Value = -2 Then
                                  description = description + "relation strength = -2" + vbNewLine
                            ElseIf mat1(i, j).Value = -1 Then
                                  description = description + "relation strength = -1" + vbNewLine
                            ElseIf mat1(i, j).Value = 1 Then
                                  description = description + "relation strength = 1" + vbNewLine
                            ElseIf mat1(i, j).Value = 2 Then
                                  description = description + "relation strength = 2" + vbNewLine
                            ElseIf mat1(i, j).Value = 3 Then
                                  description = description + "relation strength = 3" + vbNewLine
                            End If
                         End If
                    End If
                 Next
             End If
           Next j
           
           'effect strength
           
           effect_strength = es(i).Value
           description = description + "Effect strength = " + CStr(effect_strength)
           Sheets("effect-strength-diagram").Hyperlinks.Add shp1, "", "", ScreenTip:=description2 + vbNewLine + description
         End If
   Next
'   If (StrComp(descritpition = "", vbTextCompare) = 0) = False Then
'      'Sheets("effect-strength-diagram").Hyperlinks.Add shp1, "", "", ScreenTip:=description
'   End If
Next i
End Sub

Sub error_handling_r_strength()
  MsgBox "Invalid input input in relation strength sheet "
End Sub
Sub Build()
Sheets("Diagram2").Shapes.SelectAll
Selection.ShapeRange.Group.Name = "Group99"
Sheets("Diagram2").Shapes("Group99").Copy
Sheets("direction-diagram").Paste
End Sub
Function Modulo(a, b)
    Modulo = a - (b * (a \ b))
End Function

Sub CreateAutoshapes(x As Integer, y As Integer, e As String, z As Integer)
  Dim i As Integer
  Dim t As Integer
  Dim shp As Shape
  Set shp = Sheets("Diagram2").Shapes.AddShape(msoShapeRoundedRectangle, y, x, 250, 130)

  shp.Name = z
  shp.TextFrame.Characters.Text = e
  shp.TextFrame.Characters.Font.Size = 15
  shp.TextFrame.Characters.Font.Bold = True
  
  shp.TextFrame.Characters.Font.ColorIndex = 0
  shp.Fill.ForeColor.SchemeColor = 1
End Sub

Sub CreateAutoshapes1(x As Double, y As Double, kk As Integer)
  Dim i As Integer
  Dim t As Integer
  Dim shp As Shape
  Set shp = Sheets("effect-strength-diagram").Shapes.AddShape(msoShapeOval, y, x, 20, 20)
  shp.TextFrame.Characters.Text = "*"
  shp.TextFrame.VerticalAlignment = xlVAlignCenter
  shp.TextFrame.HorizontalAlignment = xlHAlignCenter
  shp.TextFrame.Characters.Font.Size = 20
  If kk = 1 Then
     shp.Fill.ForeColor.SchemeColor = 3
  ElseIf kk = 0 Then
     shp.Fill.ForeColor.SchemeColor = 2
  End If
  
End Sub

Sub CreateAutoshapes4(x As Integer, y As Integer, e As String)
  Dim shp As Shape
  Set shp = Sheets("Diagram2").Shapes.AddShape(msoShapeRectangle, y, x, 420, 30)
  shp.Fill.ForeColor.SchemeColor = 0
  shp.TextFrame.Characters.Text = e
  shp.TextFrame.Characters.Font.ColorIndex = 2
  shp.TextFrame.Characters.Font.Size = 20
  
End Sub

Sub CreateAutoshapes3(x As Integer, y As Integer, kk As Integer)
  Dim i As Integer
  Dim t As Integer
  Dim shp As Shape
  Set shp = Sheets("Diagram2").Shapes.AddShape(msoShapeOval, y, x, 20, 20)
  If kk = 0 Then
     shp.Fill.ForeColor.SchemeColor = 2
  Else
     shp.Fill.ForeColor.SchemeColor = 4
  End If
End Sub

Sub CreateAutoshapes5(x As Integer, y As Integer, kk As Integer, rrr As String)
  Dim shp As Shape
  Set shp = Sheets("strength-diagram").Shapes.AddShape(msoShapeOval, y, x, 20, 20)
  shp.TextFrame.Characters.Text = rrr
  If kk = 0 Then
       shp.TextFrame.Characters.Font.ColorIndex = 4
       shp.Fill.ForeColor.SchemeColor = 1
       shp.TextFrame.VerticalAlignment = xlVAlignCenter
       shp.TextFrame.HorizontalAlignment = xlHAlignCenter
  Else
       shp.TextFrame.Characters.Font.ColorIndex = 4
       shp.Fill.ForeColor.SchemeColor = 1
       shp.TextFrame.VerticalAlignment = xlVAlignCenter
       shp.TextFrame.HorizontalAlignment = xlHAlignCenter
  End If
End Sub

Sub AddConnectorBetweenShapes3(ConnectorType As MsoConnectorType, oBeginShape As Shape, oEndShape As Shape, pos1 As Integer, x As Integer, y As Integer, pos2 As Integer, pos3 As Integer, per As Integer, per2 As Integer, y2 As Double, pTwo As Integer, twostep1 As Double, twostep2 As Double, count2 As Double, count3 As Double, perspectivestart As Integer, perspectiveend As Integer, mm As Integer, ss As Integer, t3 As Integer, rrr As String, t4 As Integer, fff As Integer, IJ As String)
With oBeginShape
    bx = .Left + .Width
    by = .Top + (.Height / 2)
End With
With oEndShape
    ex = .Left + .Width
    ey = .Top + (.Height / 2)
End With
Dim xxx As Integer
Dim yyy As Integer
Dim xxx1 As Integer
Dim yyy1 As Integer
If nn <> mm Then
   nn = mm
   xxx = oEndShape.Left + t3
   yyy = oEndShape.Top - 20
   xxx1 = bx + (ex - bx) / 50
   yyy1 = by + (ey - by) / 50
Else
   xxx = oEndShape.Left + t3
   yyy = oEndShape.Top - 20
   xxx1 = bx + (ex - bx) / 50
   yyy1 = by + (ey - by) / 50
End If

If perspectivestart = 1 And perspectiveend = 1 Then
     With Sheets("Diagram2")
        With .Shapes.AddConnector(msoConnectorCurve, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorCurve
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Name = IJ
            .Line.Weight = 5#
            .Line.EndArrowheadWidth = msoArrowheadWide
            .Line.EndArrowheadLength = msoArrowheadLong
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
            
        End With
      End With
     
ElseIf perspectivestart = 2 And perspectiveend = 2 Then
    With Sheets("Diagram2")
        With .Shapes.AddConnector(msoConnectorCurve, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 2
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorCurve
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Name = IJ
            .Line.Weight = 5#
            .Line.EndArrowheadWidth = msoArrowheadWide
            .Line.EndArrowheadLength = msoArrowheadLong
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
        End With
     End With

ElseIf perspectivestart = 3 And perspectiveend = 3 Then
    With Sheets("Diagram2")
         With .Shapes.AddConnector(msoConnectorCurve, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorCurve
            .Line.Weight = 5#
            .Line.EndArrowheadWidth = msoArrowheadWide
            .Line.EndArrowheadLength = msoArrowheadLong
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Name = IJ
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
         End With
   End With
ElseIf perspectivestart = 4 And perspectiveend = 4 Then
    With Sheets("Diagram2")
        With .Shapes.AddConnector(msoConnectorCurve, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorCurve
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Name = IJ
            .Line.Weight = 5#
            .Line.EndArrowheadWidth = msoArrowheadWide
            .Line.EndArrowheadLength = msoArrowheadLong
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
           
        End With
   End With

ElseIf perspectivestart = 5 And perspectiveend = 5 Then
    With Sheets("Diagram2")
        With .Shapes.AddConnector(msoConnectorCurve, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorCurve
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.Weight = 5#
            .Line.EndArrowheadWidth = msoArrowheadWide
            .Line.EndArrowheadLength = msoArrowheadLong
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Name = IJ
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
        End With
   End With
ElseIf perspectivestart = 6 And perspectiveend = 6 Then
    With Sheets("Diagram2")
        With .Shapes.AddConnector(msoConnectorCurve, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorCurve
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Name = IJ
            .Line.Weight = 5#
            .Line.EndArrowheadWidth = msoArrowheadWide
            .Line.EndArrowheadLength = msoArrowheadLong
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
        End With
   End With
ElseIf perspectivestart = 7 And perspectiveend = 7 Then
    With Sheets("Diagram2")
        With .Shapes.AddConnector(msoConnectorCurve, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorCurve
            .Name = IJ
            .Line.Weight = 5#
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.EndArrowheadWidth = msoArrowheadWide
            .Line.EndArrowheadLength = msoArrowheadLong
            .Line.ForeColor.RGB = RGB(0, 0, 0)
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
           
        End With
   End With
Else


With Sheets("Diagram2")
   With .Shapes.AddConnector(msoConnectorStraight, bx, by, bx + 100, by + 100)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorStraight
            .Line.Weight = 5#
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.EndArrowheadWidth = msoArrowheadWide
            .Line.EndArrowheadLength = msoArrowheadLong
            .Name = IJ
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(0, 0, 0)
    End With
End With
End If
'Call CreateAutoshapes5(yyy1, xxx1, fff, rrr)
'Call CreateAutoshapes3(yyy, xxx, ss)
End Sub

Sub AddConnectorBetweenShapes4(ConnectorType As MsoConnectorType, oBeginShape As Shape, oEndShape As Shape, pos1 As Integer, x As Integer, y As Integer, pos2 As Integer, pos3 As Integer, per As Integer, per2 As Integer, y2 As Double, pTwo As Integer, twostep1 As Double, twostep2 As Double, count2 As Double, count3 As Double, perspectivestart As Integer, perspectiveend As Integer, mm As Integer, ss As Integer, t3 As Integer, rrr As String, t4 As Integer, fff As Integer, dis_bet_ova As Double)

With oBeginShape
    bx = .Left + .Width
    by = .Top + (.Height / 2)
End With
With oEndShape
    ex = .Left + .Width
    ey = .Top + (.Height / 2)
End With
Dim xxx As Integer
Dim yyy As Integer
Dim xxx1 As Integer
Dim yyy1 As Integer
If nn <> mm Then
   nn = mm
   xxx = oEndShape.Left + t3
   yyy = oEndShape.Top - 20
   xxx1 = bx + (ex - bx) / 50
   yyy1 = by + (ey - by) / 50
Else
   xxx = oEndShape.Left + t3
   yyy = oEndShape.Top - 20
   xxx1 = bx + (ex - bx) / 50
   yyy1 = by + (ey - by) / 50
   
End If
Call CreateAutoshapes5(yyy1, xxx1, fff, rrr)
'Call CreateAutoshapes3(yyy, xxx, ss)
End Sub
   
Sub AddConnectorBetweenShapes2(ConnectorType As MsoConnectorType, oBeginShape As Shape, oEndShape As Shape, pos1 As Integer, x As Integer, y As Integer, pos2 As Integer, pos3 As Integer, per As Integer, per2 As Integer, y2 As Double, pTwo As Integer, twostep1 As Double, twostep2 As Double, count2 As Double, count3 As Double, perspectivestart As Integer, perspectiveend As Integer, mm As Integer, ss As Integer, t3 As Integer, rrr As Integer, t4 As Integer, fff As Integer)
Dim xxx As Integer
Dim yyy As Integer

With oBeginShape
    bx = .Left + .Width + 50
    by = .Top + .Height + 50
End With
With oEndShape
    bx = .Left + .Width + 50
    by = .Top + .Height + 50
End With

If nn <> mm Then
   nn = mm
   tt = tt + 10
   xxx = oEndShape.Left + tt
   yyy = oEndShape.Top - 20
Else
   xxx = oEndShape.Left + tt
   yyy = oEndShape.Top - 20
   
End If

If perspectivestart = 1 Then
  If perspectiveend = 1 Then
    With Sheets("Diagram2")
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 2
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
        End With
   End With
   
   ElseIf perspectiveend = 2 Then
    With Sheets("Diagram2")
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y2
            l = .Left
            t = .Top
        End With
   End With
   
   Call CreateAutoshapes3(yyy, xxx, ss)
   ElseIf perspectiveend = 3 Then
    With Sheets("Diagram2")
    
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(2) = twostep2
        End With
        
   End With
   ElseIf perspectiveend = 4 Then
   Call CreateAutoshapes3(yyy, xxx, ss)
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
   ElseIf perspectiveend = 5 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
   ElseIf perspectiveend = 6 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
   ElseIf perspectiveend = 7 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  End If
ElseIf perspectivestart = 2 Then
    If perspectiveend = 2 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 2
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 3 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y2
            l = .Left
            t = .Top
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 4 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 5 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 6 Then
  With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 7 Then
  With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
  End If
  Call CreateAutoshapes3(yyy, xxx, ss)
ElseIf perspectivestart = 3 Then
    If perspectiveend = 3 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 2
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 4 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y2
            l = .Left
            t = .Top
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
   ElseIf perspectiveend = 5 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 6 Then
   With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
  Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 7 Then
  With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  End If
  
ElseIf perspectivestart = 4 Then
    If perspectiveend = 4 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 2
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 5 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y2
            l = .Left
            t = .Top
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 6 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 7 Then
  With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
   End If
   
   ElseIf perspectivestart = 5 Then
    If perspectiveend = 5 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 2
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 6 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y2
            l = .Left
            t = .Top
        End With
    'Next j
   End With
  Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 7 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 1
            .ConnectorFormat.EndConnect oEndShape, 4
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            '.Adjustments.Item(1) = twostep1
            .Adjustments.Item(2) = twostep2
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
   End If
ElseIf perspectivestart = 6 Then
    If perspectiveend = 6 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 2
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
        End With
    'Next j
   End With
  Call CreateAutoshapes3(yyy, xxx, ss)
  ElseIf perspectiveend = 7 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 4
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y2
            l = .Left
            t = .Top
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
   End If
   ElseIf perspectivestart = 7 Then
    If perspectiveend = 7 Then
    With Sheets("Diagram2")
    'For j = 2 To 3 'mainshape.ConnectionSiteCount
        With .Shapes.AddConnector(msoConnectorElbow, bx, by, bx + 200, by + 200)
            .ConnectorFormat.BeginConnect oBeginShape, 2
            .ConnectorFormat.EndConnect oEndShape, 2
            .ConnectorFormat.Type = msoConnectorElbow
            .Line.EndArrowheadStyle = msoArrowheadTriangle
            .Line.ForeColor.RGB = RGB(pos1, pos2, pos3)
            .Adjustments.Item(1) = y
            l = .Left
            t = .Top
        End With
    'Next j
   End With
   Call CreateAutoshapes3(yyy, xxx, ss)
   End If
End If
End Sub




