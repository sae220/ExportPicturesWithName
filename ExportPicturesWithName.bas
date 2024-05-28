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

' ���ׂẴX���C�h�ɂ���摜��ۑ�����
Sub exportPicturesWithName()
    Dim folderPath As String
    
    Dim sld As Slide
    Dim shp As Shape
    
    Dim picture As Shape
    Dim textBox As Shape
    Dim nearestTextBox As Shape
    
    Dim filePath As String
    
    ' �摜���o�͂���t�H���_��I������
    MsgBox "�摜���o�͂���t�H���_��I�����Ă�������"
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        If .Show = 0 Then
            MsgBox "�L�����Z���{�^����������܂���", vbExclamation
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With
    
    For Each sld In ActivePresentation.Slides
        
        ' �摜�ƃe�L�X�g�{�b�N�X���擾����i�O���[�o���ϐ��Ɋi�[�j
        Set pictures = New Collection
        Set textBoxes = New Collection
        Call classifyShapes(sld.Shapes)
        
        ' �e�L�X�g�{�b�N�X���Ȃ��X���C�h�̓X�L�b�v����
        If textBoxes.Count = 0 Then
            If pictures.Count <> 0 Then
                lastMessage = lastMessage & vbCrLf & "�X���C�h" & CStr(sld.SlideIndex) & "�ɂ̓e�L�X�g�{�b�N�X���Ȃ����߃X�L�b�v����܂���"
            End If
            GoTo ContinueSlideLoop
        End If
        
        ' ��ԋ߂��e�L�X�g�{�b�N�X�����Z�b�g
        Set nearestTextBox = Nothing
        
        For Each picture In pictures
            
            ' ��ԋ߂��e�L�X�g�{�b�N�X��T��
            For Each textBox In textBoxes
                If nearestTextBox Is Nothing Then
                    Set nearestTextBox = textBox
                Else
                    If calcSquareDistanceBetweenShapes(picture, textBox) < calcSquareDistanceBetweenShapes(picture, nearestTextBox) Then
                        Set nearestTextBox = textBox
                    End If
                End If
            Next
            
            ' �t�@�C�����𒲐�����
            filePath = folderPath & "\" & removeForbiddenCharacters(nearestTextBox.TextFrame2.TextRange.Text) & ".png"
            filePath = convertUnique(filePath)
            
            ' �p�X����256�����ȏ�ł���΃X�L�b�v����
            If Len(filePath) >= 256 Then
                lastMessage = lastMessage & vbCrLf & "�X���C�h" & CStr(sld.SlideIndex) & "�A�I�u�W�F�N�g�u" & picture.Name & "�v�̓p�X�ɂ͒�������̂ŃX�L�b�v����܂���"
                GoTo ContinuePictureLoop
            End If
            
            ' �o�͂���
            Call exportPictureInOriginalSize(picture, filePath)
ContinuePictureLoop:
        Next
ContinueSlideLoop:
    Next
    
    MsgBox "�������܂���" & lastMessage
End Sub

' Shape�I�u�W�F�N�g���摜���e�L�X�g�{�b�N�X�����ނ���i�O���[�v�̏ꍇ�͍ċA�I�ɏ�������j
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

' �摜�ƃe�L�X�g�{�b�N�X�̊Ԃ̋����i�̕����j���v�Z����i�f�t�H���g�͉摜�̉E��A�e�L�X�g�{�b�N�X�̍����j
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

' �I�u�W�F�N�g�̊p�̍��W���擾����
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

' �����񂩂���s�Ȃǂ̋֎~�����𔲂�
Function removeForbiddenCharacters(text_ As String, Optional replacement As String = "") As String
    Dim forbiddenCharacter As Variant  ' String
    Dim forbiddenCharacters() As Variant  ' String
    forbiddenCharacters = Array(vbCr, "<", ">", ":", """", "/", "\", "|", "?", "*")
    For Each forbiddenCharacter In forbiddenCharacters
        text_ = Replace(text_, forbiddenCharacter, replacement)
    Next
    removeForbiddenCharacters = text_
End Function

' �������O�̃t�@�C��������ꍇ�̓C���f�b�N�X�����邩�ύX���邩���čċA�I�ɏ�������
Function convertUnique(path As String) As String

    ' �t�@�C�������݂��Ȃ��Ȃ炻�̂܂�
    If Not (CreateObject("Scripting.FileSystemObject").FileExists(path)) Then
        convertUnique = path
        Exit Function
    End If
    
    ' ���݂���Ȃ琳�K�\���ŕ������ĕύX���ă��j�[�N�ɂ���
    With CreateObject("VBScript.RegExp")
        Dim mc As Object  ' MatchCollection
        
        .Pattern = "_(\d+)\.png$"  ' _123.png�݂����Ȋ���
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

' �摜��{���̃T�C�Y�ŏo�͂���
Sub exportPictureInOriginalSize(picture As Shape, path As String)
    Dim pictureHeight As Single
    Dim pictureWidth As Single
    
    ' ���̑傫����ۊǂ���
    pictureHeight = picture.Height
    pictureWidth = picture.Width
    
    ' �{���̃T�C�Y�ɂ���
    picture.ScaleHeight 1#, msoTrue
    
    ' �G�N�X�|�[�g����
    On Error Resume Next
        picture.Export path, ppShapeFormatPNG
    On Error GoTo 0
    
    ' ���̑傫���ɂɖ߂�
    picture.Height = pictureHeight
    pictureWidth = pictureWidth
End Sub

