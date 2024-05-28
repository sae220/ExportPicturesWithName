Attribute VB_Name = "Module11"
Dim lastMessage As String

Dim pictures As Collection
Dim textBoxes As Collection

Enum Corner
    TopLeft
    BottomLeft
    BottomRight
    TopRight
End Enum

' すべてのスライドにある画像を保存する
Sub exportPicturesWithName()
    Dim folderPath As String
    
    Dim sld As Slide
    Dim shp As Shape
    
    Dim picture As Shape
    Dim textBox As Shape
    Dim nearestTextBox As Shape
    
    Dim filePath As String
    
    ' 画像を出力するフォルダを選択する
    MsgBox "画像を出力するフォルダを選択してください"
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show = 0 Then
            MsgBox "キャンセルボタンが押されました", vbExclamation
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With
    
    For Each sld In ActivePresentation.Slides
        
        ' 画像とテキストボックスを取得する（グローバル変数に格納）
        Set pictures = New Collection
        Set textBoxes = New Collection
        Call classifyShapes(sld.Shapes)
        
        ' テキストボックスがないスライドはスキップする
        If textBoxes.Count = 0 Then
            If pictures.Count <> 0 Then
                lastMessage = lastMessage & vbCrLf & "スライド" & CStr(sld.SlideIndex) & "にはテキストボックスがないためスキップされました"
            End If
            GoTo ContinueSlideLoop
        End If
        
        ' 一番近いテキストボックスをリセット
        Set nearestTextBox = Nothing
        
        For Each picture In pictures
            
            ' 一番近いテキストボックスを探す
            For Each textBox In textBoxes
                If nearestTextBox Is Nothing Then
                    Set nearestTextBox = textBox
                Else
                    If calcSquareDistanceBetweenShapes(picture, textBox) < calcSquareDistanceBetweenShapes(picture, nearestTextBox) Then
                        Set nearestTextBox = textBox
                    End If
                End If
            Next
            
            ' ファイル名を調整する
            filePath = folderPath & "\" & removeForbiddenCharacters(nearestTextBox.TextFrame2.TextRange.Text) & ".png"
            filePath = convertUnique(filePath)
            
            ' パス名が256文字以上であればスキップする
            If Len(filePath) >= 256 Then
                lastMessage = lastMessage & vbCrLf & "スライド" & CStr(sld.SlideIndex) & "、オブジェクト「" & picture.Name & "」はパスには長すぎるのでスキップされました"
                GoTo ContinuePictureLoop
            End If
            
            ' 出力する
            Call exportPictureInOriginalSize(picture, filePath)
ContinuePictureLoop:
        Next
ContinueSlideLoop:
    Next
    
    MsgBox "完了しました" & lastMessage
End Sub

' Shapeオブジェクトを画像かテキストボックスか分類する（グループの場合は再帰的に処理する）
Function classifyShapes(shps As Object) As Collection()  ' Shapes | GroupShapes
    For Each shp In shps
        Select Case shp.Type
            Case msoPicture
                pictures.Add shp
            Case msoTextBox
                textBoxes.Add shp
            Case msoGroup
                Call classifyShapes(shp.GroupItems)
        End Select
    Next
End Function

' 画像とテキストボックスの間の距離（の平方）を計算する（デフォルトは画像の右上、テキストボックスの左下）
Function calcSquareDistanceBetweenShapes(picture As Shape, textBox As Shape, Optional pictureCorner As Corner = Corner.TopRight, Optional textBoxCorner As Corner = Corner.BottomLeft) As Single
    Dim position() As Single
    Dim pictureHorizontalPosition As Single
    Dim pictureVerticalPosition As Single
    Dim textBoxHorizontalPosition As Single
    Dim textBoxVerticalPosition As Single
    
    position() = getPosition(picture, pictureCorner)
    pictureHorizontalPosition = position(0)
    pictureVerticalPosition = position(1)
    position() = getPosition(textBox, textBoxCorner)
    textBoxHorizontalPosition = position(0)
    textBoxVerticalPosition = position(1)
    
    calcSquareDistanceBetweenShapes = (textBoxHorizontalPosition - pictureHorizontalPosition) ^ 2 + (textBoxVerticalPosition - pictureVerticalPosition) ^ 2
End Function

' オブジェクトの角の座標を取得する
Function getPosition(shp As Shape, crnr As Corner) As Single()
    Dim result(1) As Single
    Dim horizontalPosition As Single
    Dim verticalPosition As Single
    
    Select Case crnr
        Case Corner.TopLeft
            horizontalPosition = shp.Left
            verticalPosition = shp.Top
        Case Corner.BottomLeft
            horizontalPosition = shp.Left
            verticalPosition = shp.Top + shp.Height
        Case Corner.BottomRight
            horizontalPosition = shp.Left + shp.Width
            verticalPosition = shp.Top + shp.Height
        Case Corner.TopRight
            horizontalPosition = shp.Left + shp.Width
            verticalPosition = shp.Top
    End Select
    
    result(0) = horizontalPosition
    result(1) = verticalPosition
    getPosition = result()
End Function

' 文字列から改行などの禁止文字を抜く
Function removeForbiddenCharacters(text_ As String, Optional replacement As String = "") As String
    Dim forbiddenCharacter As Variant  ' String
    Dim forbiddenCharacters() As Variant  ' String
    forbiddenCharacters = Array(vbCr, "<", ">", ":", """", "/", "\", "|", "?", "*")
    For Each forbiddenCharacter In forbiddenCharacters
        text_ = Replace(text_, forbiddenCharacter, replacement)
    Next
    removeForbiddenCharacters = text_
End Function

' 同じ名前のファイルがある場合はインデックスをつけるか変更するかして再帰的に処理する
Function convertUnique(path As String) As String

    ' ファイルが存在しないならそのまま
    If Not (CreateObject("Scripting.FileSystemObject").FileExists(path)) Then
        convertUnique = path
        Exit Function
    End If
    
    ' 存在するなら正規表現で分解して変更してユニークにする
    With CreateObject("VBScript.RegExp")
        Dim mc As Object  ' MatchCollection
        
        .Pattern = "_(\d+)\.png$"  ' _123.pngみたいな感じ
        Set mc = .Execute(path)
        
        If mc.Count = 0 Then
            path = Left(path, Len(path) - 4) & "_1.png"
        Else
            Dim m As Object  ' Match
            Dim index As Integer
            
            Set m = mc(0)
            index = CInt(m.SubMatches(0))
            path = Left(path, m.FirstIndex) & "_" & CStr(index + 1) & ".png"
        End If
    End With
    convertUnique = convertUnique(path)
End Function

' 画像を本来のサイズで出力する
Sub exportPictureInOriginalSize(picture As Shape, path As String)
    Dim pictureHeight As Single
    Dim pictureWidth As Single
    
    ' 元の大きさを保管する
    pictureHeight = picture.Height
    pictureWidth = picture.Width
    
    ' 本来のサイズにする
    picture.ScaleHeight 1#, msoTrue
    
    ' エクスポートする
    On Error Resume Next
        picture.Export path, ppShapeFormatPNG
    On Error GoTo 0
    
    ' 元の大きさにに戻す
    picture.Height = pictureHeight
    pictureWidth = pictureWidth
End Sub

