VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private allCells As Range
Private sn As Snake
Public bounds As Range
Public leftColBound As Long
Public rightColBound As Long
Public topRowBound As Long
Public bottomRowBound As Long

Public Sub Class_Initialize()
    Set allCells = GameSheet.Range(GameSheet.cells(1, 1), GameSheet.cells(GameSheet.rows.Count, GameSheet.columns.Count))
    allCells.ClearFormats
    allCells.ClearContents
    'allCells.Font.size = 1
    'this does nothing because the formatting is pasted from the sprites page, which needs font to be properly sized
    
    'display settings
    With Data.settings
        If .Exists("bg-col") Then
            With allCells.FormatConditions
                .Delete
                .Add(xlCellValue, xlEqual, "=""""").Interior.Color = Data.settings("bg-col")
            End With
        End If
        If .Exists("bg-img") Then
            GameSheet.SetBackgroundPicture (Data.filepath & Data.settings("bg-img"))
        End If
    End With
    
    GameSheet.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

Public Sub Start()
    Set bounds = GameSheet.Range(GameSheet.cells(15, 15), GameSheet.cells(115, 225))
    leftColBound = 15
    rightColBound = 225
    bottomRowBound = 115
    topRowBound = 15
    
    'data settings
    Dim col As Long
    Dim row As Long
    
    Dim wallSprite As Sprite
    Dim spriteWidth As Long
    Dim spriteHeight As Long
    Set wallSprite = Data.sprites("Wall")
    spriteWidth = wallSprite.width(1)
    spriteHeight = wallSprite.height(1)
    
    'For row = 0 To (bottomRowBound - topRowBound) / spriteHeight
    '    For col = 0 To (rightColBound - leftColBound) / spriteWidth
    '        GameSheet.cells(row * spriteHeight, col * spriteWidth).Value = "shit"
    '    Next col
    'Next cell
    
    ' set stage
    Application.ScreenUpdating = False
    
    Set sn = Factory.NewSnake(Me, GameSheet.cells(45, 95))
    
    AddShit (GameSheet.cells(85, 95))
    AddShit (GameSheet.cells(85, 105))
    AddShit (GameSheet.cells(85, 115))
    
    For col = 0 To (rightColBound - leftColBound) / spriteWidth
        wallSprite.Paste 1, GameSheet.cells(topRowBound, leftColBound + col * spriteWidth)
        wallSprite.Paste 1, GameSheet.cells(bottomRowBound, leftColBound + col * spriteWidth)
    Next col
    For row = 0 To (bottomRowBound - topRowBound) / spriteHeight
        wallSprite.Paste 1, GameSheet.cells(topRowBound + row * spriteHeight, leftColBound)
        wallSprite.Paste 1, GameSheet.cells(topRowBound + row * spriteHeight, rightColBound)
    Next row
    
    Application.ScreenUpdating = True
    
End Sub

Public Sub AddShit(r As Range)
    r.value = "shit"
    Data.sprites("Shit").Paste 1, r
End Sub

Public Sub Over()
    MsgBox "You crashed! Game Over! Your score was: " & sn.size & "!"
End Sub

Public Property Get player() As Snake
    Set player = sn
End Property
