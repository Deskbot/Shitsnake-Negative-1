VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Snake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private head As Range
Private tail As Range
Private thisGame As Game
Private snakeSprite As Sprite
Private shitSprite As Sprite
Private snakeSubSprite As Dictionary
Private shitCount As Integer
Private segments As Variant
Private direction As Integer ' clockwise 0,1,2,3 0=forward
Public size As Integer

Public Sub Class_Initialize()
    Set snakeSprite = Data.sprites("Snake")
    Set shitSprite = Data.sprites("Shit")
    Set snakeSubSprite = New Dictionary
    Set segments = CreateObject("System.Collections.Queue")
    direction = 5
    With snakeSubSprite
        .Add "head", 1
        .Add "seg", 2
    End With
End Sub

Public Sub Init(g As Game, r As Range)
    Set thisGame = g
    
    Set head = r
    snakeSprite.Paste snakeSubSprite("head"), head
    
    Dim i As Long
    Dim segR As Range
    Dim numOfBodySegs As Integer
    
    numOfBodySegs = 3
    For i = 1 To numOfBodySegs
        Set segR = head.Offset(0, -(-i + numOfBodySegs + 1) * snakeSprite.width(snakeSubSprite("head")))
        segments.Enqueue segR
        snakeSprite.Paste snakeSubSprite("seg"), segR
    Next i
    segments.Enqueue head
End Sub

Public Sub moveLeft()
    If head.Column = 1 Or head.Offset(0, -1).Column < thisGame.leftColBound Then
        thisGame.Over
    ElseIf direction <> 1 Then
        Dim nHead As Integer
        nHead = snakeSubSprite("head")
        snakeSprite.Paste snakeSubSprite("seg"), head
        Set head = head.Offset(0, -snakeSprite.width(nHead))
        segments.Enqueue head
        snakeSprite.Paste nHead, head
        
        Dim seg As Range
        
        If IsOnShit Then
            shitCount = shitCount + 4
            size = size + 1
        Else
            Set seg = segments.Dequeue
            If shitCount > 0 Then
                shitCount = shitCount - 1
                thisGame.AddShit seg
            Else
                PosToArea(seg).ClearFormats
            End If
        End If
        
        direction = 3
    End If
End Sub

Public Sub moveRight()
    If head.Column = GameSheet.columns.Count Or head.Offset(0, 1).Column > thisGame.rightColBound Then
        thisGame.Over
    ElseIf direction <> 3 Then
        Dim nHead As Integer
        nHead = snakeSubSprite("head")
        snakeSprite.Paste snakeSubSprite("seg"), head
        Set head = head.Offset(0, snakeSprite.width(nHead))
        segments.Enqueue head
        snakeSprite.Paste snakeSubSprite("head"), head
        
        Dim seg As Range
        
        If IsOnShit Then
            shitCount = shitCount + 4
            size = size + 1
        Else
            Set seg = segments.Dequeue
            If shitCount > 0 Then
                shitCount = shitCount - 1
                thisGame.AddShit seg
            Else
                PosToArea(seg).ClearFormats
            End If
        End If
        
        direction = 1
    End If
End Sub

Public Sub moveUp()
    If head.Column = GameSheet.columns.Count Or head.Offset(-1, 0).row < thisGame.topRowBound Then
        thisGame.Over
    ElseIf direction <> 2 Then
        Dim nHead As Integer
        nHead = snakeSubSprite("head")
        snakeSprite.Paste snakeSubSprite("seg"), head
        Set head = head.Offset(-snakeSprite.height(nHead), 0)
        segments.Enqueue head
        snakeSprite.Paste snakeSubSprite("head"), head
        
        Dim seg As Range
        
        If IsOnShit Then
            shitCount = shitCount + 4
            size = size + 1
        Else
            Set seg = segments.Dequeue
            If shitCount > 0 Then
                shitCount = shitCount - 1
                thisGame.AddShit seg
            Else
                PosToArea(seg).ClearFormats
            End If
        End If
        
        direction = 0
    End If
End Sub

Public Sub moveDown()
    If head.Column = GameSheet.columns.Count Or head.Offset(1, 0).row > thisGame.bottomRowBound Then
        thisGame.Over
    ElseIf direction <> 0 Then
        Dim nHead As Integer
        nHead = snakeSubSprite("head")
        snakeSprite.Paste snakeSubSprite("seg"), head
        Set head = head.Offset(snakeSprite.height(nHead), 0)
        segments.Enqueue head
        snakeSprite.Paste snakeSubSprite("head"), head
        
        Dim seg As Range
        
        If IsOnShit Then
            shitCount = shitCount + 4
            size = size + 1
        Else
            Set seg = segments.Dequeue
            If shitCount > 0 Then
                shitCount = shitCount - 1
                thisGame.AddShit seg
            Else
                PosToArea(seg).ClearFormats
            End If
        End If
        
        direction = 2
    End If
End Sub

Public Function IsOnShit() As Boolean
    IsOnShit = head.value = "shit"
End Function

Public Function PosToArea(r As Range) As Range
    Dim i As Integer
    i = snakeSubSprite("head")
    Set PosToArea = GameSheet.Range(r, r.Offset(snakeSprite.height(i), snakeSprite.width(i)))
End Function

