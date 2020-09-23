Attribute VB_Name = "Module1"
Option Explicit

Type PersonInfo
    BookID As String * 5
    BookName As String * 40
    LenderName As String * 40
    Date As String * 10
    Returned As String * 1
    
End Type

Public Sub Shadow(f As Form, c As Control, shWidth As Integer, Color As String)
Dim oldWidth As Integer
Dim oldScale As Integer
oldWidth = f.DrawWidth
oldScale = f.ScaleMode
f.ScaleMode = 3
f.DrawWidth = 1
f.Line (c.Left + shWidth, c.Top + shWidth)-Step(c.Width - 1, c.Height - 1), Color, BF
f.DrawWidth = oldWidth
f.ScaleMode = oldScale
End Sub

