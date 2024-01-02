VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bm4bits40 
   Caption         =   "Imagem BMP 4 bits"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2910
   OleObjectBlob   =   "bm4bits40.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "bm4bits40"
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
    Dim i As Integer    ' Índices
    
    ' Primeira estrutura 'Bitmap File Header' contém informações sobre o tipo,
    ' tamanho e layout de um bitmap e ocupa 14 bytes (padrão).
    
    HX = HX & "424D"        ' BitmapFileType         WORD               4D42 = 19778, 42 = 66 4D = 77 "BM"       O tipo de arquivo ("BM").
    HX = HX & "72010000"    ' BitmapFileSize         DOUBLE WORD    00000172 = 14 + 40 + 64 + 252 = 370 bytes    O tamanho do arquivo bitmap.
    HX = HX & "0000"        ' BitmapFileReserved1    WORD               0000 = 0 byte                            Reservados (0 byte)
    HX = HX & "0000"        ' BitmapFileReserved2    WORD               0000 = 0 byte                            Reservados (0 byte)
    HX = HX & "76000000"    ' BitmapFileOffBits      DOUBLE WORD    00000076 = 14 + 40 + 64 = 118 bytes          O deslocamento desde o início da estrutura BITMAPFILEHEADER até os bits de bitmap.
    
    ' Segunda estrutura 'Bitmap Info Header' é semelhante à primeira, porém
    ' contém dados reduzidos, apenas informações sobre as dimensões e formato de
    ' cores de um bitmap e ocupa 40 bytes (padrão).
    
    HX = HX & "28000000"    ' BitmapInfoSize             DOUBLE WORD    00000028 = 40 bytes     Especifica o número de bytes exigidos pela estrutura.
    HX = HX & "12000000"    ' BitmapInfoWidth            LONG           00000012 = 18 pixels    Especifica a largura do bitmap.
    HX = HX & "15000000"    ' BitmapInfoHeight           LONG           00000015 = 21 pixels    Especifica a altura do bitmap.
    HX = HX & "0100"        ' BitmapInfoPlanes           WORD               0001 = 1 plano      Especifica o número de planos para o dispositivo de destino. (1 plano)
    HX = HX & "0400"        ' BitmapInfoBitCount         WORD               0004 = 4 bpp        Especifica o número de bits por pixel.
    HX = HX & "00000000"    ' BitmapInfoCompression      DOUBLE WORD    00000000 = 0 nenhuma    Especifica o formato de vídeo compactado. (0 nenhuma)
    HX = HX & "FC000000"    ' BitmapInfoSizeImage        DOUBLE WORD    000000FC = 252 bytes    Especifica o tamanho da imagem.
    HX = HX & "00000000"    ' BitmapInfoXPelsPerMeter    LONG           00000000 = 0 ppm        Especifica a resolução horizontal do dispositivo de destino para o bitmap. (0 ppm)
    HX = HX & "00000000"    ' BitmapInfoYPelsPerMeter    LONG           00000000 = 0 ppm        Especifica a resolução vertical do dispositivo de destino para o bitmap. (0 ppm)
    HX = HX & "00000000"    ' BitmapInfoClrUsed          DOUBLE WORD    00000000 = 0 atributo   Especifica o número de índices de cores na tabela de cores que são realmente usados pelo bitmap. (0 attribute)
    HX = HX & "00000000"    ' BitmapInfoClrImportant     DOUBLE WORD    00000000 = 0 atributo   Especifica o número de índices de cores que são considerados importantes para exibir o bitmap. (0 attribute)
    
    ' Terceira estrutura 'Palette' só será necessária para bitmaps menores que
    ' 24 bits, quando não for possível inserir as cores RGB ou ARGB de cada
    ' pixel diretamente no bitmap e, como nosso bitmap tem 4 bit e utiliza o
    ' cabeçalho Info/RGB, ela ocupa 16 cores * 4 bytes = 64 bytes.
    
    HX = HX & "00000000"    ' 0 Black              00000000 = ARGB(000, 000, 000, 000)
    HX = HX & "00008000"    ' 1 Maroon             00800000 = ARGB(000, 128, 000, 000)
    HX = HX & "00800000"    ' 2 Green              00008000 = ARGB(000, 000, 128, 000)
    HX = HX & "00808000"    ' 3 Olive              00808000 = ARGB(000, 128, 128, 000)
    HX = HX & "80000000"    ' 4 Navy               00000080 = ARGB(000, 000, 000, 128)
    HX = HX & "80008000"    ' 5 Purple             00800080 = ARGB(000, 128, 000, 128)
    HX = HX & "80800000"    ' 6 Teal               00008080 = ARGB(000, 000, 128, 128)
    HX = HX & "80808000"    ' 7 Gray               00808080 = ARGB(000, 128, 128, 128)
    HX = HX & "C0C0C000"    ' 8 Silver             00C0C0C0 = ARGB(000, 192, 192, 192)
    HX = HX & "0000FF00"    ' 9 Red                00FF0000 = ARGB(000, 255, 000, 000)
    HX = HX & "00FF0000"    ' A Lime               0000FF00 = ARGB(000, 000, 255, 000)
    HX = HX & "00FFFF00"    ' B Yellow             00FFFF00 = ARGB(000, 255, 255, 000)
    HX = HX & "FF000000"    ' C Blue               000000FF = ARGB(000, 000, 000, 255)
    HX = HX & "FF00FF00"    ' D Magenta/Fuchsia    00FF00FF = ARGB(000, 255, 000, 255)
    HX = HX & "FFFF0000"    ' E Cyan/Aqua          0000FFFF = ARGB(000, 000, 255, 255)
    HX = HX & "FFFFFF00"    ' F White              00FFFFFF = ARGB(000, 255, 255, 255)
       
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
    
    Open Project.ThisDocument.Path & "\~$bm4bits40.bmp" For Binary Access Write As #1
        For i = 0 To Len(HX) - 1 Step 2
            BT = BT & Chr(Val("&H" & Mid(HX, i + 1, 2)))
        Next
        Put #1, , BT
    Close #1
    
    ' Visualizar o arquivo bitmap.
    
    Me.Image1.Picture = LoadPicture(Project.ThisDocument.Path & "\~$bm4bits40.bmp")
    
End Sub
