VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private spr As Sprite
Private sprState As Integer
Private pos As Range

Public Sub Class_Initialize()
    sprState = 1
End Sub

Public Sub Init(s As Sprite)
    Set spr = s
End Sub

Public Property Let position(r As Range)
    Set pos = r
End Property

Public Property Get position() As Range
    Set position = pos
End Property

Public Sub MovePositionRelative(rows As Long, columns As Long)
    If pos.row + rows > 0 And pos.Column + columns > 0 Then
        Application.ScreenUpdating = False
        ClearSprite
        Set pos = pos.Offset(rows, columns)
        spr.Paste sprState, pos
        Application.ScreenUpdating = True
    End If
End Sub

Public Sub MovePosition(r As Range)
    Application.ScreenUpdating = False
    ClearSprite
    Set pos = r
    spr.Paste sprState, pos
    Application.ScreenUpdating = True
End Sub

Public Sub UpdatePosition()
    Application.ScreenUpdating = False
    ClearSprite
    spr.Paste sprState, pos
    Application.ScreenUpdating = True
End Sub

Public Function Overlaps(elem As Element) As Boolean
    Overlaps = Application.Intersect(SpriteRange, elem.SpriteRange) <> Null
End Function

Public Function OverlapsAny(coll As Collection) As Collection
    Set OverlapsAny = New Collection
    
    Dim elem As Element
    For Each elem In coll
        If Overlaps(elem) Then
            OverlapsAny.Add elem
        End If
    Next elem
End Function

Private Sub ClearSprite()
    If Not pos Is Nothing Then
        SpriteRange.ClearFormats
    End If
End Sub

Public Property Get SpriteRange() As Range
    Set SpriteRange = GameSheet.Range(pos, pos.Offset(spr.height(sprState), spr.width(sprState)))
End Property


