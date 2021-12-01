Attribute VB_Name = "mdmRutinasObligatorias"
Option Explicit
    'Cantidad de Boxes en A3
    Global Const NumRowsA3 As Integer = 6
    Global Const NumColsA3 As Integer = 4

    Public NumSerie As String
    Public NumRowsCSV As String
    Public Codificacion As String
    Public SmallArray(23, 2)
    Public MiniArray(3, 2)

Sub AplicarCambiosObligatorios()
    
    ValidarDatos
    DetectarBaseDatos
    BorrarContenidoCapasObligatoriasPrevias
    CrearTextoNumSerieEnA3
    CrearCodificacionEnA3

End Sub
Sub ValidarDatos()
    'Estableciendo el tipo de codificacion del codigo de barras y el Numero de Serie
    Codificacion = frmMain.cbo_Codificacion.Value
    NumSerie = frmMain.txt_NumSerie.Value
    'Now check the entered data
    If NumSerie = "" Or Len(NumSerie) <> 7 Then
        MsgBox Prompt:="Numero de Serie no tiene 7 digitos.", Buttons:=vbOKOnly, Title:="Validación del campo NumSerie"
        End
    End If
    
    'FALTA VALIDAR QUE LAYER #5 EXISTE
    ActiveDocument.Unit = cdrMillimeter
    
End Sub

Sub DetectarBaseDatos()
    Dim UbicacionBD As String
    UbicacionBD = frmMain.txt_UbicacionBD.Value
    
    '***************************************************************
    'Obtener Numero Filas CSV
    '***************************************************************
    Dim objFSO, txsInput
    Const ForReading = 1
    
    'Open TXT object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set txsInput = objFSO.OpenTextFile(UbicacionBD, ForReading)
    'Skip lines one by one
    Do While txsInput.AtEndOfStream <> True
        txsInput.SkipLine ' or strTemp = txsInput.ReadLine
    Loop
    'Saving number of lines as a rows
    NumRowsCSV = txsInput.Line
    'Cleanup
    Set objFSO = Nothing

End Sub

Function BuscarRegistroBaseDatosQR(ByVal Fila As Long) As String
    Dim BaseDatosQR As String
    BaseDatosQR = frmMain.txt_BaseDatosQR.Value
    
    Dim objFSO, txsInput
    Const ForReading = 1
    'Open TXT object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set txsInput = objFSO.OpenTextFile(BaseDatosQR, ForReading)
    'Skip lines one by one
    Do While txsInput.AtEndOfStream <> True
        
        If txsInput.Line = Fila Then
            BuscarRegistroBaseDatosQR = txsInput.ReadLine
            Exit Do
        End If
        
        txsInput.SkipLine ' or strTemp = txsInput.ReadLine
    Loop
    
    'Cleanup
    Set objFSO = Nothing
    
End Function
Sub BorrarContenidoCapasObligatoriasPrevias()
    Dim l As Layer, lr As Layers
    Dim i As Integer
      
    ActiveDocument.MasterPage.Layers("Layer3").Shapes.All.Delete
    
    For i = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(i).Activate
        If ActivePage.Layers("Layer11").Shapes.Count >= 1 Then
            ActiveDocument.Pages(i).Delete
        Else
            ActivePage.Layers("Layer7").Shapes.All.Delete
            ActivePage.Layers("Layer8").Shapes.All.Delete
            ActivePage.Layers("Layer9").Shapes.All.Delete
            ActivePage.Layers("Layer10").Shapes.All.Delete
            ActivePage.Layers("Layer11").Shapes.All.Delete
            ActivePage.Layers("Layer12").Shapes.All.Delete
            ActivePage.Layers("Layer13").Shapes.All.Delete
            ActivePage.Layers("Layer14").Shapes.All.Delete
        End If
    Next i
    
End Sub
Sub CrearTextoNumSerieEnA3()
    Dim shDATA_OBJECT1 As Shape
    Dim s As Shape
    Dim sr As ShapeRange
    Dim ValorNegro As String
    
    ValorNegro = Val(frmMain.txt_NivelNegroCodificacion.Value)
    
    'Se seleccionan todos los elementos de Pagina1/Layer5
    ActiveDocument.Pages(1).Activate
    ActivePage.Layers("Layer5").Activate
    Set sr = ActiveLayer.Shapes.All
    
    'Se recorren los textos de Layer5 (NumCartilla) para extraer su valor y posición
    For Each s In sr
        'Imprime NumSerie (Layer3 Master) en posición relativa segun los NumCartilla (Layer5)
        Set shDATA_OBJECT1 = ActiveDocument.MasterPage.Layers("Layer3").CreateArtisticText(s.PositionX + 74.341, s.PositionY - 2.563, NumSerie, , , "Arial", 14, cdrFalse, , , cdrCenterAlignment)
    Next s
    
    ActiveDocument.MasterPage.Layers("Layer3").Editable = True
    ActiveDocument.MasterPage.Layers("Layer3").Activate
    Set sr = ActiveLayer.Shapes.All
    sr.CreateSelection
    ActiveSelection.Fill.UniformColor.CMYKAssign 0, 0, 0, ValorNegro
    ActiveDocument.MasterPage.Layers("Layer3").Editable = False
    
End Sub
Sub CrearCodificacionEnA3()
    Dim shDATA_OBJECT1 As Shape
    Dim x, i As Integer
    Dim s As Shape
    Dim sr As ShapeRange
    Dim PosRelativaX, PosRelativaY As Double
    Dim NumCartilla As Long
    Dim strCode As String
    Dim strBarCode As String
    Dim clBarChars As Clase1
    Dim ValorNegro As Integer
    Dim grupoQR As Integer
    
    grupoQR = Val(frmMain.txt_grupoQR.Value)
    
    ValorNegro = Val(frmMain.txt_NivelNegroCodificacion.Value)
    
    'Seteando la clase para Code128
    Set clBarChars = New Clase1
    
    'Se recorreran todas las páginas para agregar NumSerie y Codificacion
    For x = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(x).Activate
        ActivePage.Layers("Layer5").Activate
        Set sr = ActiveLayer.Shapes.All
        i = 0   'Contador para SmallArray(23, 2) que se usa cuando Codificacion="QR Retira"
        
        For Each s In sr
            'Se recorreran los textos de Layer5 (NumCartilla) para extraer su valor y posición
            NumCartilla = Val(Right(s.Text.Story.Text, Len(s.Text.Story.Text) - 8))
            strCode = NumSerie & CStr(Format(NumCartilla, "00000"))
            PosRelativaX = s.PositionX
            PosRelativaY = s.PositionY
                        
            'Imprime strCode segun Codificacion
            If Codificacion = "EAN13" Then
                strBarCode = EAN_13(strCode)
                Set shDATA_OBJECT1 = ActivePage.Layers("Layer8").CreateArtisticText(PosRelativaX + 44.8, PosRelativaY - 7.563 - 1.5 + 0.5, strBarCode, , , "Code EAN13", 36, cdrFalse, , , cdrCenterAlignment)
                shDATA_OBJECT1.Stretch 1, 0.4 'Instruccion que permite escalar verticalmente
            
            ElseIf Codificacion = "Code128" Then
                strBarCode = clBarChars.Code128_Str(strCode)
                Set shDATA_OBJECT1 = ActivePage.Layers("Layer7").CreateArtisticText(PosRelativaX + 44.8, PosRelativaY - 6.563, strBarCode, , , "Code 128", 26, cdrFalse, , , cdrCenterAlignment)
                shDATA_OBJECT1.Stretch 1, 0.6 'Instruccion que permite escalar verticalmente el barcode
            
            Else
                'Se genera Array con TextoQR, PosX y PosY que luego se procesará.
                'SmallArray(i, 0) = strCode
                SmallArray(i, 0) = NumCartilla + (Val(grupoQR) - 1) * (NumRowsCSV - 1)
                SmallArray(i, 1) = PosRelativaX + 30.25
                SmallArray(i, 2) = PosRelativaY - 7.25
                i = i + 1
                If x = ActiveDocument.Pages.Count And i = 1 Then
                    frmMain.txt_LastQRtext.Value = SmallArray(0, 0) & "  " & BuscarRegistroBaseDatosQR(SmallArray(0, 0))
                End If
            End If
        Next s
        
        'Asignación de Color
        If Codificacion = "EAN13" Then
            ActivePage.Layers("Layer8").Activate
            Set sr = ActiveLayer.Shapes.All
            sr.CreateSelection
            ActiveSelection.Fill.UniformColor.CMYKAssign 0, 0, 0, ValorNegro
        ElseIf Codificacion = "Code128" Then
            ActivePage.Layers("Layer7").Activate
            Set sr = ActiveLayer.Shapes.All
            sr.CreateSelection
            ActiveSelection.Fill.UniformColor.CMYKAssign 0, 0, 0, ValorNegro
        End If
        
               
        If Codificacion = "QRCode" Then
            'Se recorreran todos los objetos de la capa NumTira
            ActivePage.Layers("Layer6").Activate
            Set sr = ActiveLayer.Shapes.All
            i = 0   'Contador para MiniArray()
            
            For Each s In sr
            'Se genera MiniArray() con NumTira, su PosX y su PosY .
                MiniArray(i, 0) = s.Text.Story.Text
                MiniArray(i, 1) = s.RotationCenterX
                MiniArray(i, 2) = s.RotationCenterY
                i = i + 1
            Next s
            
            GeneraRetiraQRenA3

        End If
        
    Next x
    
    SampleINIRespaldoObligatorias
    
    MsgBox Prompt:="Se agregó el N° de Serie y la Codificación en todas las páginas", Buttons:=vbOKOnly
    
End Sub
Sub GeneraRetiraQRenA3()
    Dim PosicionQR_X, PosicionQR_Y As Double
    Dim strFile_Path As String
    Dim i As Integer
    Dim pg As Page
    Dim shDATA_OBJECT1 As Shape
    Dim ValorNegro As String
    Dim sr As ShapeRange
    Dim ValorQR As String
    Dim ColorCMYK As Variant
    Dim AjustePosQR As Variant
    Dim AjusteTamañoQR As Variant
    Dim CarpetaLocalQR As String
    
    AjustePosQR = Split(frmMain.txt_AjustePosQR.Value, ",")
    AjusteTamañoQR = Split(frmMain.txt_AjusteTamañoQR.Value, ",")
    CarpetaLocalQR = frmMain.txt_CarpetaLocalQR.Value
        
    ValorNegro = Val(frmMain.txt_NivelNegroCodificacion.Value)
    ValorQR = frmMain.txt_NivelNegroQRCodificacion.Value
    ColorCMYK = Split(ValorQR, ",")
    
    ValorNegro = Val(frmMain.txt_NivelNegroCodificacion.Value)
    'ValorNegroQR = Val(frmMain.txt_NivelNegroQRCodificacion.Value)
    
    'Se crea nueva pagina Retira y se agrega un cuadrado de fondo A3 de color blanco
    Set pg = ActiveDocument.InsertPages(1, False, ActivePage.Index)
    Set shDATA_OBJECT1 = ActivePage.Layers("Layer11").CreateRectangle2(ActivePage.LeftX, ActivePage.BottomY, ActivePage.SizeWidth, ActivePage.SizeHeight)
    shDATA_OBJECT1.Fill.UniformColor.CMYKAssign 0, 0, 0, 0

    'Recorre las 24 cartillas de la hoja A3
    For i = 0 To NumRowsA3 * NumColsA3 - 1
        'strFile_Path = DownloadFile("http://api.qrserver.com/v1/create-qr-code/?data=" & SmallArray(i, 0) & "&ecc=L&size=71x71&format=png")
        'If the program does not obtain a QR barcode exit.
        'If strFile_Path = "" Then
        '    MsgBox "Existe un problema con la conexión a Internet o el servidor http://api.qrserver.com no está disponible."
        '    Exit Sub
        'End If
            
        'Se configura posición del QR de acuerdo con SmallArray() pero se invierten las coordenadas del eje X
        PosicionQR_X = SmallArray(23 - i, 1) + AjustePosQR(0)
        PosicionQR_Y = SmallArray(i, 2) + AjustePosQR(1)
        
        'Se configura posición del QR de acuerdo con SmallArray() pero se invierten las coordenadas del eje Y
        'PosicionQR_X = SmallArray(i, 1)
        'PosicionQR_Y = SmallArray(23 - i, 2)
            
        'Se imprime el codigoQR en Layer13
        'ActivePage.Layers("Layer13").Import (strFile_Path)
        ActivePage.Layers("Layer13").Import (CarpetaLocalQR & SmallArray(i, 0) & ".png")
        ActiveSelection.Stretch AjusteTamañoQR(0), AjusteTamañoQR(1)
        ActiveSelection.SetPosition PosicionQR_X, PosicionQR_Y
    Next i
    'Asignando color a todas las imagenes QR en hoja A3
    ActivePage.Layers("Layer13").Activate
    Set sr = ActiveLayer.Shapes.All
    sr.CreateSelection
    sr.SetOutlineProperties Color:=CreateCMYKColor(ColorCMYK(0), ColorCMYK(1), ColorCMYK(2), ColorCMYK(3))
    
    'Se imprime Texto Nº de Tira que va en la Retira A3
    For i = 0 To 3
        Set shDATA_OBJECT1 = ActivePage.Layers("Layer14").CreateArtisticText(MiniArray(3 - i, 1) + 80 + 15, MiniArray(i, 2), MiniArray(i, 0), , , "Century Schoolbook", 10, cdrTrue, , , cdrCenterAlignment)
        shDATA_OBJECT1.Rotate 270
    Next i
    'Asignando color al NªTira
    ActivePage.Layers("Layer14").Activate
    Set sr = ActiveLayer.Shapes.All
    sr.CreateSelection
    ActiveSelection.Fill.UniformColor.CMYKAssign 0, 0, 0, ValorNegro
    
     ActivePage.Layers("Layer11").Editable = False
     ActivePage.Layers("Layer13").Editable = False
     ActivePage.Layers("Layer14").Editable = False
End Sub

Sub SampleINIRespaldoObligatorias()

    Dim sINI_FILE As String
    Dim sReturn As String

    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"

    ' Write to the ini file
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "NumSerie", sINI_FILE, frmMain.txt_NumSerie.Value)
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "Codificacion", sINI_FILE, frmMain.cbo_Codificacion.Value)
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "NivelNegro", sINI_FILE, frmMain.txt_NivelNegroCodificacion)
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "NivelNegroQR", sINI_FILE, frmMain.txt_NivelNegroQRCodificacion)
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "grupoQR", sINI_FILE, frmMain.txt_grupoQR.Value)
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "CarpetaLocalQR", sINI_FILE, frmMain.txt_CarpetaLocalQR)
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "BaseDatosQR", sINI_FILE, frmMain.txt_BaseDatosQR)
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "AjustePosQR", sINI_FILE, frmMain.txt_AjustePosQR)
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "AjusteTamañoQR", sINI_FILE, frmMain.txt_AjusteTamañoQR)

 End Sub

Sub CambiarNivelColorSoloCodificacion()
    Dim s As Shape, sr As ShapeRange
    Dim x, ValorNegro As Integer
    Dim l As Layer, lr As Layers
    Dim ValorQR As String
    Dim ColorCMYK As Variant
        
    ValorNegro = Val(frmMain.txt_NivelNegroCodificacion.Value)
    ValorQR = frmMain.txt_NivelNegroQRCodificacion.Value
    ColorCMYK = Split(ValorQR, ",")
    
    'Para cambio de % color negro SOLO en la Codificacion:
    '- Numeros y letras (requieren cambiar relleno):
    '-- Capas Master: Layer3
    '-- Capas normales: Layer7, Layer8, Layer12, Layer14
    '- Formas(requieren cambiar Pluma):
    '-- Capas Master:
    '-- Capas normales: Layer13
        
    For x = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(x).Activate
        Set lr = ActivePage.AllLayers
        
        For Each l In lr
            'Hace cambios en las Capas del tipo cdrTextShape que son NO-Master
            If l.Name = "Layer7" Or l.Name = "Layer8" Or l.Name = "Layer12" Or l.Name = "Layer14" Then
                l.Activate
                l.Editable = True
                Set sr = ActiveLayer.Shapes.All
                If sr.Count <> 0 Then
                    sr.CreateSelection
                    ActiveSelection.Fill.UniformColor.CMYKAssign 0, 0, 0, ValorNegro
                End If
                'l.Editable = False
            'Hace cambios en la Capa que contiene las imagenes PNG del codigo QR
            ElseIf l.Name = "Layer13" Then
                l.Activate
                l.Editable = True
                Set sr = ActiveLayer.Shapes.All
                If sr.Count <> 0 Then
                    sr.CreateSelection
                    sr.SetOutlineProperties Color:=CreateCMYKColor(ColorCMYK(0), ColorCMYK(1), ColorCMYK(2), ColorCMYK(3))
                End If
                'l.Editable = False
            End If
            
            If x = 1 Then
                'Hace cambios en las Capas del tipo cdrTextShape que son MASTER
                If l.Name = "Layer3" Then
                    ActiveDocument.MasterPage.Layers(l.Name).Editable = True
                    ActiveDocument.MasterPage.Layers(l.Name).Activate
                    Set sr = ActiveLayer.Shapes.All
                    If sr.Count <> 0 Then
                        sr.CreateSelection
                        ActiveSelection.Fill.UniformColor.CMYKAssign 0, 0, 0, ValorNegro
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
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "NivelNegro", sINI_FILE, frmMain.txt_NivelNegroCodificacion.Value)
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "NivelNegroQR", sINI_FILE, frmMain.txt_NivelNegroQRCodificacion.Value)

End Sub


Sub CambiarNivelColorNegroSoloCodificacion()        'REEMPLAZADO POR CambiarNivelColorSoloCodificacion()
    Dim s As Shape, sr As ShapeRange
    Dim x, ValorNegro, ValorNegroQR As Integer
    Dim l As Layer, lr As Layers
        
    ValorNegro = Val(frmMain.txt_NivelNegroCodificacion.Value)
    ValorNegroQR = Val(frmMain.txt_NivelNegroQRCodificacion.Value)
    
    'Para cambio de % color negro SOLO en la Codificacion:
    '- Numeros y letras (requieren cambiar relleno):
    '-- Capas Master: Layer3
    '-- Capas normales: Layer7, Layer8, Layer12, Layer14
    '- Formas(requieren cambiar Pluma):
    '-- Capas Master:
    '-- Capas normales: Layer13
        
    For x = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(x).Activate
        Set lr = ActivePage.AllLayers
        
        For Each l In lr
            'Hace cambios en las Capas del tipo cdrTextShape que son NO-Master
            If l.Name = "Layer7" Or l.Name = "Layer8" Or l.Name = "Layer12" Or l.Name = "Layer14" Then
                l.Activate
                l.Editable = True
                Set sr = ActiveLayer.Shapes.All
                If sr.Count <> 0 Then
                    sr.CreateSelection
                    ActiveSelection.Fill.UniformColor.CMYKAssign 0, 0, 0, ValorNegro
                End If
                l.Editable = False
            'Hace cambios en la Capa que contiene las imagenes PNG del codigo QR
            ElseIf l.Name = "Layer13" Then
                l.Activate
                l.Editable = True
                Set sr = ActiveLayer.Shapes.All
                If sr.Count <> 0 Then
                    sr.CreateSelection
                    sr.SetOutlineProperties Color:=CreateCMYKColor(0, 0, 0, ValorNegroQR)
                End If
                l.Editable = False
            End If
            
            If x = 1 Then
                'Hace cambios en las Capas del tipo cdrTextShape que son MASTER
                If l.Name = "Layer3" Then
                    ActiveDocument.MasterPage.Layers(l.Name).Editable = True
                    ActiveDocument.MasterPage.Layers(l.Name).Activate
                    Set sr = ActiveLayer.Shapes.All
                    If sr.Count <> 0 Then
                        sr.CreateSelection
                        ActiveSelection.Fill.UniformColor.CMYKAssign 0, 0, 0, ValorNegro
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
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "NivelNegro", sINI_FILE, frmMain.txt_NivelNegroCodificacion.Value)
    sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "NivelNegroQR", sINI_FILE, frmMain.txt_NivelNegroQRCodificacion.Value)

End Sub

