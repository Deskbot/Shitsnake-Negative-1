VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String
Private subSprites As Collection

Public Sub Init(n As String)
    pName = n
    Set subSprites = New Collection
End Sub

Public Property Get name() As String
    name = pName
End Property

Public Property Get subSprite(index As Integer) As Range
    Set subSprite = subSprites.Item(index)
End Property

Public Property Get height(index As Integer) As Long
    height = subSprites(index).rows.Count
End Property

Public Property Get width(index As Integer) As Long
    width = subSprites(index).columns.Count
End Property

Public Sub AddSubSprite(r As Range)
    subSprites.Add r
End Sub

Public Sub Paste(index As Integer, r As Range)
    subSprites(index).Copy
    r.PasteSpecial Paste:=xlPasteFormats
End Sub
