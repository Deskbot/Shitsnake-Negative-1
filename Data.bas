Attribute VB_Name = "Data"
Option Explicit

Public filepath As String
Public settings As Dictionary
Public sprites As Dictionary
Public spritesCount As Integer

Public Sub CompileData()
    filepath = Application.ActiveWorkbook.Path & "\"
    Data.CompileSettings
    Data.CompileSprites
End Sub

Private Sub CompileSettings()
    Dim startCell As Range
    Dim settingsRange As Range
    Dim c As Range
    
    Set settings = New Dictionary
    
    With SettingsSheet
        Set startCell = .cells(1, 1)
        
        If startCell.Offset(1, 0) = "" Then 'there's only 1 setting
            settings.Add startCell.value, startCell.Offset(0, 1).value 'the cell to the right
        Else
            Set settingsRange = .Range(startCell, startCell.End(xlDown))
            
            For Each c In settingsRange
                settings.Add c.value, c.Offset(0, 1).value 'the cell to the right
            Next c
        End If
    End With
    
End Sub

Private Sub CompileSprites()
    Dim spriteName As String
    Dim spriteIndex As Range
    
    Set sprites = New Dictionary
    
    Set spriteIndex = SpritesSheet.cells(1, 1)
    
    Do While spriteIndex.row <> SpritesSheet.rows.Count
        Dim s As Sprite
        spriteName = spriteIndex.value
        Set s = Factory.NewSprite(spriteName)
        
        Dim startOfSubSprite As Range
        
        Set startOfSubSprite = spriteIndex.End(xlToRight)
        
        Do While startOfSubSprite.Column <> SpritesSheet.columns.Count
            Dim SpriteRange As Range
            Set SpriteRange = SpritesSheet.Range(startOfSubSprite, startOfSubSprite.End(xlToRight).End(xlDown))
        
            'add sprite to Sprites
            
            s.AddSubSprite SpriteRange
            
            Set startOfSubSprite = startOfSubSprite.End(xlToRight).End(xlToRight)
        Loop
        
        sprites.Add spriteName, s
        
        'next sprite
        Set spriteIndex = spriteIndex.End(xlDown)
    Loop
End Sub

