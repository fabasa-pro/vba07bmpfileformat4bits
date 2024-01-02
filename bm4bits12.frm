VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bm4bits12 
   Caption         =   "Imagem BMP 4 bits"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2910
   OleObjectBlob   =   "bm4bits12.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "bm4bits12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Licenciado sob a licença MIT.
' Copyright (C) 2012 - 2024 @Fabasa-Pro. Todos os direitos reservados.
' Consulte LICENSE.TXT na raiz do projeto para obter informações.

Option Explicit

Private Sub CommandButton1_Click()

    ' Declarações gerais:
    
    Dim HX As String    ' Dados (hexadecimal)
    Dim BT As String    ' Bytes
    Dim i As Integer    ' Índices]
    
    ' Primeira estrutura 'Bitmap File Header' contém informações sobre o tipo,
    ' tamanho e layout de um bitmap e ocupa 14 bytes (padrão).
    
    HX = HX & "424D"        ' BitmapFileType         WORD               4D42 = 19778, 42 = 66 4D = 77 "BM"       O tipo de arquivo ("BM").
    HX = HX & "46010000"    ' BitmapFileSize         DOUBLE WORD    00000146 = 14 + 12 + 48 + 252 = 326 bytes    O tamanho do arquivo bitmap.
    HX = HX & "0000"        ' BitmapFileReserved1    WORD               0000 = 0 byte                            Reservados (0 byte)
    HX = HX & "0000"        ' BitmapFileReserved2    WORD               0000 = 0 byte                            Reservados (0 byte)
    HX = HX & "4A000000"    ' BitmapFileOffBits      DOUBLE WORD    0000004A = 14 + 12 + 48 = 74 bytes           O deslocamento desde o início da estrutura BITMAPFILEHEADER até os bits de bitmap.
    
    ' Segunda estrutura 'Bitmap Core Header' é semelhante à primeira, porém
    ' contém dados reduzidos, apenas informações sobre as dimensões e formato de
    ' cores de um bitmap e ocupa 12 bytes (padrão).
    
    HX = HX & "0C000000"    ' BitmapCoreSize         DOUBLE WORD    0000000C = 12 bytes     Especifica o número de bytes exigidos pela estrutura.
    HX = HX & "1200"        ' BitmapCoreWidth        WORD           00000012 = 18 pixels    Especifica a largura do bitmap.
    HX = HX & "1500"        ' BitmapCoreHeight       WORD           00000015 = 21 pixels    Especifica a altura do bitmap.
    HX = HX & "0100"        ' BitmapCorePlanes       WORD               0001 = 1 plano      Especifica o número de planos para o dispositivo de destino. (1 plano)
    HX = HX & "0400"        ' BitmapCoreBitCoun      WORD               0004 = 4 bpp        Especifica o número de bits por pixel.
    
    ' Terceira estrutura 'Palette' só será necessária para bitmaps menores que
    ' 24 bits, quando não for possível inserir as cores RGB ou ARGB de cada
    ' pixel diretamente no bitmap e, como nosso bitmap tem 4 bit e utiliza o
    ' cabeçalho Core/RGB, ela ocupa 16 cores * 3 bytes = 48 bytes.
    
    HX = HX & "000000"      ' 0 Black              000000 = RGB(000, 000, 000)
    HX = HX & "000080"      ' 1 Maroon             800000 = RGB(128, 000, 000)
    HX = HX & "008000"      ' 2 Green              008000 = RGB(000, 128, 000)
    HX = HX & "008080"      ' 3 Olive              808000 = RGB(128, 128, 000)
    HX = HX & "800000"      ' 4 Navy               000080 = RGB(000, 000, 128)
    HX = HX & "800080"      ' 5 Purple             800080 = RGB(128, 000, 128)
    HX = HX & "808000"      ' 6 Teal               008080 = RGB(000, 128, 128)
    HX = HX & "808080"      ' 7 Gray               808080 = RGB(128, 128, 128)
    HX = HX & "C0C0C0"      ' 8 Silver             C0C0C0 = RGB(192, 192, 192)
    HX = HX & "0000FF"      ' 9 Red                FF0000 = RGB(255, 000, 000)
    HX = HX & "00FF00"      ' A Lime               00FF00 = RGB(000, 255, 000)
    HX = HX & "00FFFF"      ' B Yellow             FFFF00 = RGB(255, 255, 000)
    HX = HX & "FF0000"      ' C Blue               0000FF = RGB(000, 000, 255)
    HX = HX & "FF00FF"      ' D Magenta/Fuchsia    FF00FF = RGB(255, 000, 255)
    HX = HX & "FFFF00"      ' E Cyan/Aqua          00FFFF = RGB(000, 255, 255)
    HX = HX & "FFFFFF"      ' F White              FFFFFF = RGB(255, 255, 255)
       
    ' Quarta estrutura 'Bitmap' contém todos os pixels extrudados em uma matriz
    ' de coluna e linha, onde temos linhas de 0 a 20 = 21 de altura e 18 na
    ' largura, em partes de 32 bits, por esse motivo completamos com 0 (zero)
    ' até obter os completos 32 bits, ela ocupa 21 linhas * 12 bytes = 252 bytes.
        
    '         32 bits         32 bits         32 bits
    '     --------------- --------------- ---------------
    '  0: F F F F F F F F F F F F F F F F F F 0 0 0 0 0 0
    '  1: F F F F F F F F F F F F F F F F F F 0 0 0 0 0 0
    '  2: F F F F F F F 0 0 0 0 F F F F F F F 0 0 0 0 0 0
    '  3: F F F F F 0 0 B B B B 0 0 F F F F F 0 0 0 0 0 0
    '  4: F F F F 0 B B B B B B B B 0 F F F F 0 0 0 0 0 0
    '  5: F F F F 0 B B B B B B B B 0 F F F F 0 0 0 0 0 0
    '  6: F F F 0 B F B B B B B B B B 0 F F F 0 0 0 0 0 0
    '  7: F F F 0 F F F F B B B B B F 0 F F F 0 0 0 0 0 0
    '  8: F F F 0 F F F F F F F F F F 0 F F F 0 0 0 0 0 0
    '  9: F F F F 0 F F 0 F F 0 F F 0 F F F F 0 0 0 0 0 0
    ' 10: F F F F F 0 F 0 F F 0 F 0 F F F F F 0 0 0 0 0 0
    ' 11: F F F F 0 9 0 F F F F 0 9 0 F F F F 0 0 0 0 0 0
    ' 12: F F F 0 F 9 9 0 0 0 0 9 9 F 0 F F F 0 0 0 0 0 0
    ' 13: F F 0 F F 0 9 9 9 9 9 9 0 F F 0 F F 0 0 0 0 0 0
    ' 14: F F 0 F F 0 9 9 9 9 9 9 0 F F 0 F F 0 0 0 0 0 0
    ' 15: F F F 0 0 E 0 0 0 0 0 0 E 0 0 F F F 0 0 0 0 0 0
    ' 16: F F F F 0 E E E E E E E E 0 F F F F 0 0 0 0 0 0
    ' 17: F F F F 0 C C C 0 0 C C C 0 F F F F 0 0 0 0 0 0
    ' 18: F F F F F 0 0 0 F F 0 0 0 F F F F F 0 0 0 0 0 0
    ' 19: F F F F F F F F F F F F F F F F F F 0 0 0 0 0 0
    ' 20: F F F F F F F F F F F F F F F F F F 0 0 0 0 0 0

    HX = HX & "FFFFFFFFFFFFFFFFFF000000"    ' 20 :                                     0 0 0 0 0 0
    HX = HX & "FFFFFFFFFFFFFFFFFF000000"    ' 19 :                                     0 0 0 0 0 0
    HX = HX & "FFFFF000FF000FFFFF000000"    ' 18 :           0 0 0     0 0 0           0 0 0 0 0 0
    HX = HX & "FFFF0CCC00CCC0FFFF000000"    ' 17 :         0 C C C 0 0 C C C 0         0 0 0 0 0 0
    HX = HX & "FFFF0EEEEEEEE0FFFF000000"    ' 16 :         0 E E E E E E E E 0         0 0 0 0 0 0
    HX = HX & "FFF00E000000E00FFF000000"    ' 15 :       0 0 E 0 0 0 0 0 0 E 0 0       0 0 0 0 0 0
    HX = HX & "FF0FF09999990FF0FF000000"    ' 14 :     0     0 9 9 9 9 9 9 0     0     0 0 0 0 0 0
    HX = HX & "FF0FF09999990FF0FF000000"    ' 13 :     0     0 9 9 9 9 9 9 0     0     0 0 0 0 0 0
    HX = HX & "FFF0F99000099F0FFF000000"    ' 12 :       0   9 9 0 0 0 0 9 9   0       0 0 0 0 0 0
    HX = HX & "FFFF090FFFF090FFFF000000"    ' 11 :         0 9 0         0 9 0         0 0 0 0 0 0
    HX = HX & "FFFFF0F0FF0F0FFFFF000000"    ' 10 :           0   0     0   0           0 0 0 0 0 0
    HX = HX & "FFFF0FF0FF0FF0FFFF000000"    '  9 :         0     0     0     0         0 0 0 0 0 0
    HX = HX & "FFF0FFFFFFFFFF0FFF000000"    '  8 :       0                     0       0 0 0 0 0 0
    HX = HX & "FFF0FFFFBBBBBF0FFF000000"    '  7 :       0         B B B B B   0       0 0 0 0 0 0
    HX = HX & "FFF0BFBBBBBBBB0FFF000000"    '  6 :       0 B   B B B B B B B B 0       0 0 0 0 0 0
    HX = HX & "FFFF0BBBBBBBB0FFFF000000"    '  5 :         0 B B B B B B B B 0         0 0 0 0 0 0
    HX = HX & "FFFF0BBBBBBBB0FFFF000000"    '  4 :         0 B B B B B B B B 0         0 0 0 0 0 0
    HX = HX & "FFFFF00BBBB00FFFFF000000"    '  3 :           0 0 B B B B 0 0           0 0 0 0 0 0
    HX = HX & "FFFFFFF0000FFFFFFF000000"    '  2 :               0 0 0 0               0 0 0 0 0 0
    HX = HX & "FFFFFFFFFFFFFFFFFF000000"    '  1 :                                     0 0 0 0 0 0
    HX = HX & "FFFFFFFFFFFFFFFFFF000000"    '  0 :                                     0 0 0 0 0 0
        
    ' Salvar arquivo bitmap 16 cores (*.bmp;*.dib).
    
    Open Project.ThisDocument.Path & "\~$bm4bits12.bmp" For Binary Access Write As #1
        For i = 0 To Len(HX) - 1 Step 2
            BT = BT & Chr(Val("&H" & Mid(HX, i + 1, 2)))
        Next
        Put #1, , BT
    Close #1
    
    ' Visualizar o arquivo bitmap.
    
    Me.Image1.Picture = LoadPicture(Project.ThisDocument.Path & "\~$bm4bits12.bmp")
    
End Sub
