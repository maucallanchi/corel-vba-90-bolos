Attribute VB_Name = "mdmRutinasOpcionales"
Option Explicit
    'Cantidad de Boxes en A3
    Global Const NumRowsA3 As Integer = 6
    Global Const NumColsA3 As Integer = 4
    'Estructura del Box
    Global Const NumColsBox As Integer = 9
    Global Const NumRowsBox As Integer = 3
    'BOX_largo=90mm / BOX_alto=31.62 => Casilla_ancho: 10mm / Casilla_alto = 10.54
    Global Const Casilla_ancho As Double = 10
    Global Const Casilla_alto As Double = 10.54
    'Estructura del Array
    Global Const NumSillyCols = 2 'Si el #tira y el #Carton estan en las dos primeras columnas.
    Global Const NumSillyRows = 1 ' Si existe cabecera
    
Sub InvertirOrdenPaginas()
    Dim i, NumPagTotal As Double
    
    NumPagTotal = ActiveDocument.Pages.Count
    'Loop que permite mover las ultimas paginas siempre al inicio
    For i = 1 To NumPagTotal - 1
        ActiveDocument.Pages(NumPagTotal).MoveTo i
    Next i
End Sub

Sub AgregarTextoQR()
    Dim shDATA_OBJECT1 As Shape
    Dim x As Integer
    Dim s As Shape
    Dim sr As ShapeRange
    Dim PosRelativaX, PosRelativaY As Double
    Dim TextoQR As String
    
    TextoQR = frmMain.txt_TextoQR.Value
    
    'Se recorreran todas las páginas para agregar Texto QR
    For x = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(x).Activate
        ActivePage.Layers("Layer12").Shapes.All.Delete
        'Se recorreran los textos de Layer5 (NumCartilla) para extraer su posición y usarla para la ubicaci{on de TextoQR
        ActivePage.Layers("Layer5").Activate
        Set sr = ActiveLayer.Shapes.All
        For Each s In sr
            PosRelativaX = s.PositionX
            PosRelativaY = s.PositionY
            Set shDATA_OBJECT1 = ActivePage.Layers("Layer12").CreateArtisticText(PosRelativaX + 44.8, PosRelativaY - 2.563, TextoQR, , , "Century Schoolbook", 10, cdrTrue, , , cdrCenterAlignment)
        Next s
    Next x
    
    'Guardadndo en archivo INI
    Dim sINI_FILE As String
    Dim sReturn As String
    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "TextoQR", sINI_FILE, frmMain.txt_TextoQR.Value)
    
End Sub

Sub CambiarTextoContactanos()
    Dim x As Integer
    Dim s As Shape
    Dim sr As ShapeRange
    Dim TextoContactanos As String
    
    TextoContactanos = frmMain.txt_Contactanos.Value
    
    'Se recorreran todas las páginas
    For x = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(x).Activate
        ActiveDocument.MasterPage.Layers("Layer2").Editable = True
        ActiveDocument.MasterPage.Layers("Layer2").Activate
        Set sr = ActiveLayer.Shapes.All
        'Se recorreran todos los textos Contactanos
        For Each s In sr
            s.Text.Story.Text = TextoContactanos
        Next s
        ActiveDocument.MasterPage.Layers("Layer2").Editable = False
    Next x
    
    'Guardadndo en archivo INI
    Dim sINI_FILE As String
    Dim sReturn As String
    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "Contactanos", sINI_FILE, frmMain.txt_Contactanos.Value)
    
End Sub

Sub CambiarNivelColor()
    Dim s As Shape, sr As ShapeRange
    Dim x, ValorNegro As Integer
    Dim l As Layer, lr As Layers
    Dim ValorQR As String
    Dim ColorCMYK As Variant
        
    ValorNegro = Val(frmMain.txt_NivelNegro.Value)
    ValorQR = frmMain.txt_NivelNegroQR.Value
    ColorCMYK = Split(ValorQR, ",")
    
    'Para cambio de % color negro:
    '- Numeros y letras (requieren cambiar relleno):
    '-- Capas Master: Layer2, Layer3
    '-- Capas normales: Layer4B, Layer5, Layer6, Layer7, Layer8, Layer12, Layer14
    '- Formas(requieren cambiar Pluma):
    '-- Capas Master: Layer1, Layer4C
    '-- Capas normales: Layer4A, "Layer4B", Layer13
        
    For x = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(x).Activate
        Set lr = ActivePage.AllLayers
        
        For Each l In lr
            l.Activate
            
            'Hace cambios en las Capas del tipo cdrTextShape que son NO-Master
            If l.Name = "Layer4B" Or l.Name = "Layer5" Or l.Name = "Layer6" Or l.Name = "Layer7" Or l.Name = "Layer8" Or l.Name = "Layer12" Or l.Name = "Layer14" Then
                Set sr = ActiveLayer.Shapes.All
                If sr.Count <> 0 Then
                    l.Editable = True
                    sr.CreateSelection
                    ActiveSelection.Fill.UniformColor.CMYKAssign 0, 0, 0, ValorNegro
                    l.Editable = False
                End If
                
            'Hace cambios en las Capas del tipo cdrCurveShape que son NO-Master
            ElseIf l.Name = "Layer4A" Or l.Name = "Layer4B" Then
                Set sr = ActiveLayer.Shapes.All
                If sr.Count <> 0 Then
                    l.Editable = True
                    sr.CreateSelection
                    sr.SetOutlineProperties Color:=CreateCMYKColor(0, 0, 0, ValorNegro)
                    l.Editable = False
                End If
                
            'Hace cambios en la Capa que contiene las imagenes PNG del CODIGO QR********
            ElseIf l.Name = "Layer13" Then
                Set sr = ActiveLayer.Shapes.All
                If sr.Count <> 0 Then
                    l.Editable = True
                    sr.CreateSelection
                    sr.SetOutlineProperties Color:=CreateCMYKColor(ColorCMYK(0), ColorCMYK(1), ColorCMYK(2), ColorCMYK(3))
                    l.Editable = False
                End If
            End If
            
            If x = 1 Then
                'Hace cambios en las Capas del tipo cdrTextShape que son MASTER
                If l.Name = "Layer2" Or l.Name = "Layer3" Then
                    ActiveDocument.MasterPage.Layers(l.Name).Editable = True
                    ActiveDocument.MasterPage.Layers(l.Name).Activate
                    Set sr = ActiveLayer.Shapes.All
                    If sr.Count <> 0 Then
                        sr.CreateSelection
                        ActiveSelection.Fill.UniformColor.CMYKAssign 0, 0, 0, ValorNegro
                    End If
                    ActiveDocument.MasterPage.Layers(l.Name).Editable = False
                'Hace cambios en las Capas del tipo cdrCurveShape que son MASTER
                ElseIf l.Name = "Layer1" Or l.Name = "Layer4C" Then
                    ActiveDocument.MasterPage.Layers(l.Name).Editable = True
                    ActiveDocument.MasterPage.Layers(l.Name).Activate
                    Set sr = ActiveLayer.Shapes.All
                    If sr.Count <> 0 Then
                        sr.CreateSelection
                        sr.SetOutlineProperties Color:=CreateCMYKColor(0, 0, 0, ValorNegro)
                    End If
                    ActiveDocument.MasterPage.Layers(l.Name).Editable = False
                End If
                    
            End If
        Next l
    Next x

    'Guardadndo en archivo INI
    Dim sINI_FILE As String
    Dim sReturn As String
    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "NivelNegro", sINI_FILE, frmMain.txt_NivelNegro.Value)
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "NivelNegroQR", sINI_FILE, frmMain.txt_NivelNegroQR.Value)

End Sub

Sub AgregarFondoBox()
    Dim s As Shape, sr As ShapeRange
    Dim x As Integer
    Dim strColor As String
    Dim ColorCMYK As Variant
        
    strColor = frmMain.txt_FondoBox.Value
    ColorCMYK = Split(strColor, ",")

    For x = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(x).Activate
        ActivePage.Layers("Layer4A").Activate
        
        
        Set sr = ActiveLayer.Shapes.All
        
        If sr.Count <> 0 Then
            ActivePage.Layers("Layer4A").Editable = True
            sr.CreateSelection
            ActiveSelection.Fill.UniformColor.CMYKAssign ColorCMYK(0), ColorCMYK(1), ColorCMYK(2), ColorCMYK(3)
            ActivePage.Layers("Layer4A").Editable = False
        End If
        
        If x = 1 Then
            'Hace cambios en las Capas Master con texto o formas
            ActiveDocument.Pages(1).Activate
            ActiveDocument.MasterPage.Layers("Layer1").Editable = True
            
            ActiveDocument.MasterPage.Layers("Layer1").Activate
            Set sr = ActiveLayer.Shapes.All
            sr.CreateSelection
            ActiveSelection.Fill.UniformColor.CMYKAssign ColorCMYK(0), ColorCMYK(1), ColorCMYK(2), ColorCMYK(3)
            
            ActiveDocument.MasterPage.Layers("Layer1").Editable = False
        End If
        
    Next x

    'Guardadndo en archivo INI
    Dim sINI_FILE As String
    Dim sReturn As String
    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "FondoBox", sINI_FILE, frmMain.txt_FondoBox.Value)
End Sub

Sub AgregarFondoA3()
    Dim strColor As String
    Dim ColorCMYK As Variant
    Dim shDATA_OBJECT1 As Shape
    Dim sr As ShapeRange
        
    strColor = frmMain.txt_FondoA3.Value
    ColorCMYK = Split(strColor, ",")

    ActiveDocument.MasterPage.Layers("Layer0").Editable = True
    ActiveDocument.MasterPage.Layers("Layer0").Shapes.All.Delete
    
    Set shDATA_OBJECT1 = ActiveDocument.MasterPage.Layers("Layer0").CreateRectangle2(ActivePage.LeftX, ActivePage.BottomY, ActivePage.SizeWidth, ActivePage.SizeHeight)
    shDATA_OBJECT1.Fill.UniformColor.CMYKAssign ColorCMYK(0), ColorCMYK(1), ColorCMYK(2), ColorCMYK(3)
    
    ActiveDocument.MasterPage.Layers("Layer0").Editable = False

    'Guardadndo en archivo INI
    Dim sINI_FILE As String
    Dim sReturn As String
    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "FondoA3", sINI_FILE, frmMain.txt_FondoA3.Value)
End Sub

Sub AgregarTramadoA3()
    Dim strColor As String
    Dim ColorCMYK As Variant
    Dim s As Shape, sr As ShapeRange
    Dim shDATA_OBJECT1 As Shape
    Dim i As Integer
        
    strColor = frmMain.txt_TramadoA3.Value
    ColorCMYK = Split(strColor, ",")
    
    ActiveDocument.Unit = cdrMillimeter
    
    ActiveDocument.Pages(1).Activate
    ActiveDocument.MasterPage.Layers("Layer0A").Editable = True
    ActiveDocument.MasterPage.Layers("Layer0A").Activate
    
    If strColor = "0,0,0,0" Then
        ActiveDocument.MasterPage.Layers("Layer0A").Shapes.All.Delete
    Else
        Set sr = ActiveLayer.Shapes.All
        
        'Genera lineas divisorias Verticales en la hoja A3 si no existen
        If sr.Count = 0 Then
             i = 0
            Do While (i * 1.542) < ActivePage.SizeWidth
                Set shDATA_OBJECT1 = ActiveDocument.MasterPage.Layers("Layer0A").CreateLineSegment(ActivePage.LeftX + i * 1.542, ActivePage.BottomY, ActivePage.LeftX + i * 1.542, ActivePage.BottomY + ActivePage.SizeHeight)
                shDATA_OBJECT1.Outline.SetProperties 0.25
                i = i + 1
            Loop
            Set sr = ActiveLayer.Shapes.All
        End If
        
        'Aplica color a las lineas verticales
        sr.SetOutlineProperties Color:=CreateCMYKColor(ColorCMYK(0), ColorCMYK(1), ColorCMYK(2), ColorCMYK(3))
        
    End If
    
    ActiveDocument.MasterPage.Layers("Layer0A").Editable = False
    
    'Guardadndo en archivo INI
    Dim sINI_FILE As String
    Dim sReturn As String
    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "TramadoA3", sINI_FILE, frmMain.txt_TramadoA3.Value)
End Sub

' USO UNICO PARA GENERAR BASE DE DATOS DE LETRAS CON BASE EN CSV DE NUMEROS
Sub GeneraCSVconArrayDeLetras()
    Dim i, j As Long
    Dim ArrayLetras() As String
    Dim Prueba As String
    Dim strLocationFile2 As String
    
    'Genera ArrayLetras
    Prueba = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,Ñ,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,BB,CC,DD,EE,FF,GG,HH,II,JJ,KK,LL,MM,NN,ÑÑ,OO,PP,QQ,RR,SS,TT,UU,VV,WW,XX,YY,ZZ,AAA,BBB,CCC,DDD,EEE,FFF,GGG,HHH,III,JJJ,KKK,LLL,MMM,NNN,ÑÑÑ,OOO,PPP,QQQ,RRR,SSS,TTT,UUU,VVV,WWW,XXX,YYY,ZZZ,a,b,c,d,e,f,g,h,i"
    ArrayLetras = Split(Prueba, ",")
    
    'Intercambio de Numeros por Letras en MyArray()
    For i = NumSillyRows To NumRowsCSV - 1
        For j = NumSillyCols To NumColsCSV - 1
            If MyArray(i, j) <> "" Then
                MyArray(i, j) = ArrayLetras(MyArray(i, j) - 1)
            End If
        Next j
    Next i
    
    'Exporta MyArray() de letras a un archivo CSV
    strLocationFile2 = "C:\Users\Marlon\Documents\prueba.csv" 'Ubicacion del archivo de salida
    Open strLocationFile2 For Output As #2
    For i = 0 To NumRowsCSV - 1
        Prueba = MyArray(i, 0)
        For j = 1 To NumColsCSV - 1
            Prueba = Prueba & "," & MyArray(i, j)
        Next j
        Print #2, Prueba
    Next i
    
    MsgBox "Archivo guardado exitosamente", vbOKOnly
    
    Close #2

End Sub

Sub AgregarIcono()

    'PRINT MyArray() IN COREL DRAW
    Dim NumRowsCSVxColsA3, NumBoxActual, NumCasillaActual, NumPagTotal As Long
    Dim PosRelativaX, PosRelativaY As Double
    Dim p, q, r, s, x As Integer
    Dim shDATA_OBJECT1 As Shape

    Dim UbicacionBD, UbicacionIcono, FontLetras As String
    Dim TamañoLetras, AnchoPlumaLetras, EstiramientoLetras As Double
    Dim PosX, PosY As Double
    Dim BoxOffsetX, BoxOffsetY As Double
    Dim Contacto As String
    Dim NumRowsCSV, NumPagInicio, NumPagFinal As Long
    Dim DocPagInicio, DocPagFinal As Integer
    
    Dim NumColsCSV As Integer
    Dim MyArray() As String
    
    UbicacionIcono = frmMain.txt_UbicacionIcono.Value
    'Validación de la ruta del archivo ICONO
    If Len(Dir(UbicacionIcono)) = 0 Then
        MsgBox Prompt:="El Archivo ICONO no se encontró en la ruta indicada.", Buttons:=vbOKOnly, Title:="Validación del archivo ICONO"
        Exit Sub
    End If
    
    'Posicion de la 1era casilla del Box (Posicicion inicial X e Y)
    If frmMain.txt_PosInicialX.Value = "" Or frmMain.txt_PosInicialY.Value = "" Then
        MsgBox Prompt:="Debe revisar el valor de Posicion para X o Y", Buttons:=vbOKOnly
        End
    Else
        PosX = Val(frmMain.txt_PosInicialX.Value)
        PosY = Val(frmMain.txt_PosInicialY.Value)
    End If
    
    'Espacio entre Boxes
    If frmMain.txt_BoxOffsetX.Value = "" Or frmMain.txt_BoxOffsetY.Value = "" Then
        MsgBox Prompt:="Debe revisar el valor de Offset para X o Y", Buttons:=vbOKOnly
        End
    Else
        BoxOffsetX = Val(frmMain.txt_BoxOffsetX.Value)
        BoxOffsetY = Val(frmMain.txt_BoxOffsetY.Value)
    End If
     
    ActiveDocument.Unit = cdrMillimeter

    '***************************************************************
    'Creando los Iconos en las cartillas
    '***************************************************************
    ActiveDocument.Pages(1).Activate
    ActiveDocument.MasterPage.Layers("Layer1A").Editable = True
    ActiveDocument.MasterPage.Layers("Layer1A").Activate
    'Se recorreran las 4 columnas por cada hoja A3
    For p = 0 To NumColsA3 - 1
        'Se recorrerán los 6 filas por cada Tira
        For q = 0 To NumRowsA3 - 1
            'Se recorrera cada Fila por cada Box
            For r = 0 To NumRowsBox - 1
                'Se recorrerá cada Columna por cada Fila
                For s = 0 To NumColsBox - 1
                        'Posicion relativa (x,y)
                        PosRelativaX = PosX + 0.35 + s * Casilla_ancho + p * BoxOffsetX
                        PosRelativaY = PosY - 0.35 + Casilla_alto - r * Casilla_alto - q * BoxOffsetY
                        'Importando Icono
                        ActiveDocument.ActiveLayer.Import (UbicacionIcono)
                        ActiveSelection.SetPosition PosRelativaX, PosRelativaY
                Next s
            Next r
        Next q
    Next p
    
    ActiveDocument.MasterPage.Layers("Layer1A").Editable = False
    
    'Guardadndo en archivo INI
    Dim sINI_FILE As String
    Dim sReturn As String
    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "UbicacionIcono", sINI_FILE, frmMain.txt_UbicacionIcono.Value)
            
End Sub

Sub BorrarContenidoIconos()
    Dim l As Layer, lr As Layers
    Dim i As Integer
    
    ActiveDocument.MasterPage.Layers("Layer1A").Shapes.All.Delete
    
End Sub


Sub CambiarColorNumCart()
    Dim strColor As String
    Dim ColorCMYK As Variant
    Dim s As Shape, sr As ShapeRange
    Dim shDATA_OBJECT1 As Shape
    Dim x As Integer
    Dim l As Layer
        
    strColor = frmMain.txt_ColorNumCart.Value
    ColorCMYK = Split(strColor, ",")
        
    For x = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(x).Activate
        ActivePage.Layers("Layer4B").Activate
        Set sr = ActiveLayer.Shapes.All
        If sr.Count <> 0 Then
            ActivePage.Layers("Layer4B").Editable = True
            sr.CreateSelection
            ActiveSelection.Fill.UniformColor.CMYKAssign ColorCMYK(0), ColorCMYK(1), ColorCMYK(2), ColorCMYK(3)
            ActivePage.Layers("Layer4B").Editable = False
        End If
    Next x

    'Guardadndo en archivo INI
    Dim sINI_FILE As String
    Dim sReturn As String
    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "ColorNumCart", sINI_FILE, frmMain.txt_ColorNumCart.Value)

End Sub

Sub CambiarNivelColorNegro()        'REEMPLAZADA POR CambiarNivelColor()
    Dim s As Shape, sr As ShapeRange
    Dim x, ValorNegro, ValorNegroQR As Integer
    Dim l As Layer, lr As Layers
        
    ValorNegro = Val(frmMain.txt_NivelNegro.Value)
    ValorNegroQR = Val(frmMain.txt_NivelNegroQR.Value)
    
    'Para cambio de % color negro:
    '- Numeros y letras (requieren cambiar relleno):
    '-- Capas Master: Layer2, Layer3
    '-- Capas normales: Layer4B, Layer5, Layer6, Layer7, Layer8, Layer12, Layer14
    '- Formas(requieren cambiar Pluma):
    '-- Capas Master: Layer1, Layer4C
    '-- Capas normales: Layer4A, "Layer4B", Layer13
        
    For x = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(x).Activate
        Set lr = ActivePage.AllLayers
        
        For Each l In lr
            l.Activate
            
            'Hace cambios en las Capas del tipo cdrTextShape que son NO-Master
            If l.Name = "Layer4B" Or l.Name = "Layer5" Or l.Name = "Layer6" Or l.Name = "Layer7" Or l.Name = "Layer8" Or l.Name = "Layer12" Or l.Name = "Layer14" Then
                Set sr = ActiveLayer.Shapes.All
                If sr.Count <> 0 Then
                    l.Editable = True
                    sr.CreateSelection
                    ActiveSelection.Fill.UniformColor.CMYKAssign 0, 0, 0, ValorNegro
                    l.Editable = False
                End If
                
            'Hace cambios en las Capas del tipo cdrCurveShape que son NO-Master
            ElseIf l.Name = "Layer4A" Or l.Name = "Layer4B" Then
                Set sr = ActiveLayer.Shapes.All
                If sr.Count <> 0 Then
                    l.Editable = True
                    sr.CreateSelection
                    sr.SetOutlineProperties Color:=CreateCMYKColor(0, 0, 0, ValorNegro)
                    l.Editable = False
                End If
                
            'Hace cambios en la Capa que contiene las imagenes PNG del codigo QR
            ElseIf l.Name = "Layer13" Then
                Set sr = ActiveLayer.Shapes.All
                If sr.Count <> 0 Then
                    l.Editable = True
                    sr.CreateSelection
                    sr.SetOutlineProperties Color:=CreateCMYKColor(0, 0, 0, ValorNegroQR)
                    l.Editable = False
                End If
            End If
            
            If x = 1 Then
                'Hace cambios en las Capas del tipo cdrTextShape que son MASTER
                If l.Name = "Layer2" Or l.Name = "Layer3" Then
                    ActiveDocument.MasterPage.Layers(l.Name).Editable = True
                    ActiveDocument.MasterPage.Layers(l.Name).Activate
                    Set sr = ActiveLayer.Shapes.All
                    If sr.Count <> 0 Then
                        sr.CreateSelection
                        ActiveSelection.Fill.UniformColor.CMYKAssign 0, 0, 0, ValorNegro
                    End If
                    ActiveDocument.MasterPage.Layers(l.Name).Editable = False
                'Hace cambios en las Capas del tipo cdrCurveShape que son MASTER
                ElseIf l.Name = "Layer1" Or l.Name = "Layer4C" Then
                    ActiveDocument.MasterPage.Layers(l.Name).Editable = True
                    ActiveDocument.MasterPage.Layers(l.Name).Activate
                    Set sr = ActiveLayer.Shapes.All
                    If sr.Count <> 0 Then
                        sr.CreateSelection
                        sr.SetOutlineProperties Color:=CreateCMYKColor(0, 0, 0, ValorNegro)
                    End If
                    ActiveDocument.MasterPage.Layers(l.Name).Editable = False
                End If
                    
            End If
        Next l
    Next x

    'Guardadndo en archivo INI
    Dim sINI_FILE As String
    Dim sReturn As String
    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "NivelNegro", sINI_FILE, frmMain.txt_NivelNegro.Value)
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "NivelNegroQR", sINI_FILE, frmMain.txt_NivelNegroQR.Value)

End Sub

Sub OcultarCapas()
    Dim x As Integer
    Dim lr As Layers
    Dim l As Layer
    Dim sr As ShapeRange
    
    'Se recorreran todas las páginas para agregar Texto QR
    For x = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(x).Activate
        Set lr = ActivePage.AllLayers
        For Each l In lr
            If l.Name = "Layer4B" Or l.Name = "Layer5" Or l.Name = "Layer6" Then
                ActivePage.Layers(l.Name).Editable = True
                ActivePage.Layers(l.Name).Activate
                Set sr = ActiveLayer.Shapes.All
                'If sr.Count <> 0 Then
                    ActivePage.Layers(l.Name).Visible = False
                    ActivePage.Layers(l.Name).Printable = False
                'End If
                ActivePage.Layers(l.Name).Editable = False
            End If
        Next l
    Next x
    
    ActiveDocument.Pages(1).Activate
    Set lr = ActivePage.AllLayers
    For Each l In lr
        If l.Name = "Layer1" Or l.Name = "Layer2" Or l.Name = "Layer4C" Then
            ActiveDocument.MasterPage.Layers(l.Name).Editable = True
            ActiveDocument.MasterPage.Layers(l.Name).Activate
            Set sr = ActiveDocument.MasterPage.Layers(l.Name).Shapes.All
            'If sr.Count <> 0 Then
                ActiveDocument.MasterPage.Layers(l.Name).Visible = False
                ActiveDocument.MasterPage.Layers(l.Name).Printable = False
            'End If
            ActiveDocument.MasterPage.Layers(l.Name).Editable = False
        End If
    Next l
    
End Sub

Sub MostrarCapas()
    Dim x As Integer
    Dim lr As Layers
    Dim l As Layer
    Dim sr As ShapeRange
    
    'Se recorreran todas las páginas para agregar Texto QR
    For x = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(x).Activate
        Set lr = ActivePage.AllLayers
        For Each l In lr
            If l.Name = "Layer4B" Or l.Name = "Layer5" Or l.Name = "Layer6" Then
                ActivePage.Layers(l.Name).Editable = True
                ActivePage.Layers(l.Name).Activate
                Set sr = ActiveLayer.Shapes.All
                'If sr.Count <> 0 Then
                    ActivePage.Layers(l.Name).Visible = True
                    ActivePage.Layers(l.Name).Printable = True
                'End If
                ActivePage.Layers(l.Name).Editable = False
            End If
        Next l
    Next x
    
    ActiveDocument.Pages(1).Activate
    Set lr = ActivePage.AllLayers
    For Each l In lr
        If l.Name = "Layer1" Or l.Name = "Layer2" Or l.Name = "Layer4C" Then
            ActiveDocument.MasterPage.Layers(l.Name).Editable = True
            ActiveDocument.MasterPage.Layers(l.Name).Activate
            Set sr = ActiveDocument.MasterPage.Layers(l.Name).Shapes.All
            'If sr.Count <> 0 Then
                ActiveDocument.MasterPage.Layers(l.Name).Visible = True
                ActiveDocument.MasterPage.Layers(l.Name).Printable = True
            'End If
            ActiveDocument.MasterPage.Layers(l.Name).Editable = False
        End If
    Next l
    
End Sub

Sub CambiaTamañoFuenteNumCartilla()

    Dim x As Integer
    Dim lr As Layers
    Dim l As Layer
    Dim sr As ShapeRange
    Dim s As Shape
    Dim PosX, PosY As Double
    
    'Se recorreran todas las páginas para cambiar el tamaño de la fuente de Numero Cartilla
    For x = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(x).Activate
        Set lr = ActivePage.AllLayers
        For Each l In lr
            If l.Name = "Layer4B" Then
                l.Activate
                Set sr = ActiveLayer.Shapes.All
                If sr.Count <> 0 Then
                    l.Editable = True
                    'sr.CreateSelection
                    For Each s In sr
                        PosX = s.PositionX
                        PosY = s.PositionY
                        s.Stretch 0.5, 1
                        s.PositionX = PosX + 1
                        s.PositionY = PosY
                    Next s
                    l.Editable = False
                End If
            End If
        Next l
    Next x

End Sub
