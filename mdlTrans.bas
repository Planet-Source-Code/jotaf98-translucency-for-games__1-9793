Attribute VB_Name = "mdlTrans"

'+--------------------------------------+
'| The function "PaintTrans" draws an   |
'|image to another one, but applies a   |
'|translucency effect at the same time. |
'| Use it wherever you want, but please |
'|give me some credit for it ;)         |
'|                                      |
'|(Check the Readme for instructions)   |
'|                                      |
'|Made by Jotaf98 (João F. S. Henriques)|
'|E-mail me at jotaf98@hotmail.com      |
'+--------------------------------------+


'+--------------------------------------+
'| PS: This function works best if the  |
'|Target's AutoRedraw is set to True.   |
'+--------------------------------------+

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim Red As Long
Dim Green As Long
Dim Blue As Long

Dim Red2 As Long
Dim Green2 As Long
Dim Blue2 As Long

Dim RedF As Long
Dim GreenF As Long
Dim BlueF As Long

Dim RGB1 As Long
Dim RGB2 As Long

Public Sub PaintTrans(Target As Long, Source As Long, X As Long, Y As Long, Width As Long, Height As Long)
    Dim TrnsX As Long
    Dim TrnsY As Long
    
    For TrnsY = 0 To Height - 1
        For TrnsX = 0 To Width - 1
            RGB2 = GetPixel(Source, TrnsX, TrnsY)
            
            If RGB2 <> 0 Then
                'Initializing:
                GetRGBs RGB2
                Red2 = Red
                Green2 = Green
                Blue2 = Blue
                
                RGB1 = GetPixel(Target, TrnsX + X, TrnsY + Y)
                GetRGBs RGB1
                
                'Calculations:
                RedF = (Red + Red2) / 2
                GreenF = (Green + Green2) / 2
                BlueF = (Blue + Blue2) / 2
                
                'Display pixel:
                SetPixel Target, TrnsX + X, TrnsY + Y, RGB(RedF, GreenF, BlueF)
            End If
        Next TrnsX
    Next TrnsY
End Sub

Public Sub GetRGBs(RGBVal As Long)
    
    ' This function extracts the Red, Green and Blue
    'values from a color to 3 variables with their
    'names. You may keep it as well, but please give
    'me some credit for it too ;)
    'By the way, this function isn't very exact, so
    'you may end up with the green and blue values
    'not as they should (plus or minus 1 or 2).
    
    If RGBVal = 16777215 Then
        Red = 255
        Green = 255
        Blue = 255
        Exit Sub
    End If

    Red = RGBVal And 255
    Green = RGBVal / 256 And 255
    Blue = RGBVal / 65536 And 255
    
End Sub

