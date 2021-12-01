Attribute VB_Name = "old_mdmRutinasBase"
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

Sub ReiniciarPaginasyCapas()
    Dim l As Layer, lr As Layers
    Dim i As Integer
    
    For i = ActiveDocument.Pages.Count To 1 Step -1
        If i = 1 Then
            Set lr = ActivePage.AllLayers
            For Each l In lr
                If l.IsSpecialLayer = False Or l.Name = "Layer0" Or l.Name = "Layer0A" Or l.Name = "Layer1" Or l.Name = "Layer2" Or l.Name = "Layer3" Then
                    l.Delete
                End If
            Next l
        Else
            ActiveDocument.ActivePage.Delete
        End If
    Next i
    
    ActivePage.CreateLayer "Layer0" 'Fondo Tira A3
    ActivePage.CreateLayer "Layer0A" 'Tramado de lineas verticales en Hoja A3
    ActivePage.CreateLayer "Layer1" 'Rectangulo y Lineas del Box
    ActivePage.CreateLayer "Layer2" 'Texto Contactanos
    ActivePage.CreateLayer "Layer3" 'Texto Numero de Serie
    ActivePage.CreateLayer "Layer4" 'Texto Números de las cartillas de bingo
    ActivePage.CreateLayer "Layer4A" 'Iconos de las cartillas de bingo
    ActivePage.CreateLayer "Layer5" 'Texto Nº de Cartilla
    ActivePage.CreateLayer "Layer6" 'Texto Nº de Tira
    ActivePage.CreateLayer "Layer7" 'Codigo de barras CODE128
    ActivePage.CreateLayer "Layer8" 'Codigo de barras EAN13
    ActivePage.CreateLayer "Layer9" 'Reservado
    ActivePage.CreateLayer "Layer10" 'Reservado
    ActivePage.CreateLayer "Layer11" 'Fondo de la Retira A3
    ActivePage.CreateLayer "Layer12" 'Texto QR que va en la Tira A3
    ActivePage.CreateLayer "Layer13" 'Codigo QR que va en Retira A3
    ActivePage.CreateLayer "Layer14" 'Texto Nº de Tira que va en la Retira A3
    
    ActiveDocument.MasterPage.Background = cdrPageBackgroundNone
    
    'Guardadndo en archivo INI los valores iniciales de TramadoA3, FondoA3 y FondoBox
    Dim sINI_FILE As String
    Dim sReturn As String
    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "TramadoA3", sINI_FILE, "0,0,0,0")
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "FondoA3", sINI_FILE, "0,0,0,0")
    sReturn = sManageSectionEntry(iniWrite, "Opcionales", "FondoBox", sINI_FILE, "0,0,0,0")
    
    frmMain.txt_FondoA3.Value = "0,0,0,0"
    frmMain.txt_FondoBox.Value = "0,0,0,0"
    frmMain.txt_TramadoA3 = "0,0,0,0"
    
End Sub

Sub GenerarBloquesBase()

    Dim UbicacionBD, UbicacionIcono, FontLetras As String
    Dim TamañoLetras, AnchoPlumaLetras, EstiramientoLetras As Double
    Dim PosX, PosY As Double
    Dim BoxOffsetX, BoxOffsetY As Double
    Dim Contacto As String
    Dim NumRowsCSV, NumPagInicio, NumPagFinal As Long
    
    Dim NumColsCSV As Integer
    Dim MyArray()
    
    'Ubicacion de BD
    UbicacionBD = frmMain.txt_UbicacionBD.Value
    If UbicacionBD = "" Then
        MsgBox Prompt:="Debe indicar una ubicación de Archivo BD válida.", Buttons:=vbOKOnly, Title:="Validación del campo Ubicación BD"
        End
    ElseIf Len(Dir(UbicacionBD)) = 0 Then
        MsgBox Prompt:="El Archivo BD no se encontró en la ruta indicada.", Buttons:=vbOKOnly, Title:="Validación del campo Ubicación BD"
        End
    End If
    
    'Atributos para la impresion de Cartillas de Letras
    FontLetras = frmMain.txt_FontLetras.Value
    TamañoLetras = Val(frmMain.txt_TamañoLetras.Value)
    AnchoPlumaLetras = Val(frmMain.txt_AnchoPlumaLetras.Value)
    EstiramientoLetras = Val(frmMain.txt_EstiramientoLetras.Value)
    
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
    
    'Textos variables
    Contacto = frmMain.txt_ContactanosBase.Value
    
    'Rango de Paginas a generar por bloque
    If frmMain.txt_PaginaInicial.Value = "" Or frmMain.txt_PaginaFinal.Value = "" Then
        MsgBox Prompt:="Debe revisar el rango de paginas elegido.", Buttons:=vbOKOnly
        End
    Else
        NumPagInicio = Val(frmMain.txt_PaginaInicial.Value)
        NumPagFinal = Val(frmMain.txt_PaginaFinal.Value)
    End If
     
    ActiveDocument.Unit = cdrMillimeter
    
    '***************************************************************
    'CREAR BOXES A3
    '***************************************************************
    Dim shDATA_OBJECT1 As Shape
    Dim PosRelativaX As Double
    Dim PosRelativaY As Double
    Dim p, q, r, s, x As Integer

    'Se recorreran las 4 columnas (tiras) por cada hoja A3
    For p = 0 To NumColsA3 - 1
        'Se recorrerán los 6 filas por cada Tira
        For q = 0 To NumRowsA3 - 1
            'Se imprime el marco del box
            Set shDATA_OBJECT1 = ActivePage.Layers("Layer1").CreateRectangle2(PosX + p * BoxOffsetX, PosY - 21.08 - q * BoxOffsetY, Casilla_ancho * 9, Casilla_alto * 3)
            shDATA_OBJECT1.Outline.SetProperties 0.5
            shDATA_OBJECT1.Fill.UniformColor.CMYKAssign 0, 0, 0, 0
            
            'Se recorrerá cada fila del Box y se generará rectangulos por cada casilla
            'For r = 0 To NumRowsBox - 1
                'Se recorrerá cada Columna por cada Fila
                'For s = 0 To NumColsBox - 1
                    'PosRelativaX = PosX + s * Casilla_ancho + p * BoxOffsetX
                    'PosRelativaY = PosY - r * Casilla_alto - q * BoxOffsetY
                    'Creando los recuadros del Box
                    'Set shDATA_OBJECT1 = ActivePage.Layers("Layer1").CreateRectangle2(PosRelativaX, PosRelativaY, Casilla_ancho, Casilla_alto)
                    'shDATA_OBJECT1.Outline.SetProperties 0.25
                'Next s
            'Next r
            
            'Se imprime lneas divisorias horizontales del Box
            Set shDATA_OBJECT1 = ActivePage.Layers("Layer1").CreateLineSegment(PosX + p * BoxOffsetX, PosY - q * BoxOffsetY, PosX + p * BoxOffsetX + Casilla_ancho * NumColsBox, PosY - q * BoxOffsetY)
            shDATA_OBJECT1.Outline.SetProperties 0.25
            Set shDATA_OBJECT1 = ActivePage.Layers("Layer1").CreateLineSegment(PosX + p * BoxOffsetX, PosY - q * BoxOffsetY - Casilla_alto, PosX + p * BoxOffsetX + Casilla_ancho * NumColsBox, PosY - q * BoxOffsetY - Casilla_alto)
            shDATA_OBJECT1.Outline.SetProperties 0.25
            
            'Se imprime lineas divisorias Verticales del Box
            For s = 1 To NumColsBox - 1
                PosRelativaX = PosX + s * Casilla_ancho + p * BoxOffsetX
                PosRelativaY = PosY + Casilla_alto - q * BoxOffsetY
                Set shDATA_OBJECT1 = ActivePage.Layers("Layer1").CreateLineSegment(PosRelativaX, PosRelativaY, PosRelativaX, PosRelativaY - Casilla_alto * NumRowsBox)
                shDATA_OBJECT1.Outline.SetProperties 0.25
            Next s
            
        Next q
    Next p
    
    '***************************************************************
    'Crear Texto Contacto En A3
    '***************************************************************
    'Se recorreran las 4 columnas (tiras) por cada hoja A3
    For p = 0 To NumColsA3 - 1
        'Se recorrerán los 6 filas por cada Tira
        For q = 0 To NumRowsA3 - 1
            'Imprimiendo el texto "Contacto"
            Set shDATA_OBJECT1 = ActivePage.Layers("Layer2").CreateArtisticText(PosX + 92.6 + p * BoxOffsetX, PosY - 5.657 - q * BoxOffsetY, Contacto, , , "Calibri", 7, cdrFalse, , , cdrCenterAlignment)
            shDATA_OBJECT1.Rotate (90)
        Next q
    Next p

    '***************************************************************
    'Configurar Capas Master(
    '***************************************************************
    ActivePage.Layers("Layer0").Editable = False
    ActivePage.Layers("Layer0A").Editable = False
    ActivePage.Layers("Layer1").Editable = False
    ActivePage.Layers("Layer2").Editable = False
    ActivePage.Layers("Layer3").Editable = False

    ActivePage.Layers("Layer0").Master = True
    ActivePage.Layers("Layer0A").Master = True
    ActivePage.Layers("Layer1").Master = True
    ActivePage.Layers("Layer2").Master = True
    ActivePage.Layers("Layer3").Master = True

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
    
    '***************************************************************
    'Llenar Mi Array
    '***************************************************************

    'LLENAR MyArray() CON INFO DEL CSV
    Dim i, j As Integer       'Integer va hasta 32,767
    Dim TXT As String
    Dim LineArray As Variant

    'Directory Address. A valid file number in the range 1 to 255
    Open UbicacionBD For Input As #1
    'Recorre cada fila del CSV
    For i = 0 To NumRowsCSV - 1
        'Read line into variable.
        Line Input #1, TXT
        'Remove linebreak from TXT
        TXT = Replace(TXT, vbLf, "")
        TXT = Replace(TXT, vbCrLf, "")
        'Split TXT into a 1d Array
        LineArray = Split(TXT, ",")

        'Redimensionar MyArray() en funcion al total de Filas y Columnas
        If i = 0 Then
            NumColsCSV = UBound(LineArray) + 1
            ReDim MyArray(NumRowsCSV - 1, NumColsCSV - 1)
        End If
            
        'Saving LineArray into MyArray
        For j = 0 To NumColsCSV - 1 'or UBound(LineArray)
            MyArray(i, j) = LineArray(j)
        Next j
    Next i
    'Closing file
    Close #1
    
    '***************************************************************
    'CrearPaginasConCartillas()
    '***************************************************************
    
    'PRINT MyArray() IN COREL DRAW
    Dim NumRowsCSVxColsA3, NumBoxActual, NumCasillaActual, NumPagTotal As Long
    'Dim PosRelativaX, PosRelativaY As Double
    Dim NumTira As String
    'Dim p, q, r, s, x As Integer
    'Dim shDATA_OBJECT1 As Shape
    Dim pg As Page
    Dim PosXprovisional As Double
    Dim OrigSelection As ShapeRange

    'Calculando la cantidad de Cartillas totales en una sola columna
    NumRowsCSVxColsA3 = (NumRowsCSV - NumSillyRows) / NumColsA3
    'Calculando la cantidad de PaginasA3
    NumPagTotal = NumRowsCSVxColsA3 / NumRowsA3
       
    'Se recorrerán las X páginas del Bloque seleccionado
    For x = NumPagInicio To NumPagFinal
        'Crea una página nueva en el documento
        If x > NumPagInicio And x <= NumPagFinal Then
                Set pg = ActiveDocument.InsertPages(1, False, ActivePage.Index)
        End If
        'Se recorreran las 4 columnas por cada hoja A3
        For p = 0 To NumColsA3 - 1
            'Se recorrerán los 6 filas por cada Tira
            For q = 0 To NumRowsA3 - 1
                
                'Imprimiendo las casilla del Box
                NumBoxActual = NumSillyRows + q + NumRowsA3 * (x - 1) + p * NumRowsCSVxColsA3
                
                'Se recorrera cada Fila por cada Box
                For r = 0 To NumRowsBox - 1
                    'Se recorrerá cada Columna por cada Fila
                    For s = 0 To NumColsBox - 1
                        NumCasillaActual = NumSillyCols + s + NumColsBox * r
                        'Imprime solo si el valor de la casilla es diferente a ""
                        If MyArray(NumBoxActual, NumCasillaActual) <> "" Then
                            'Posicion relativa (x,y)
                            PosRelativaX = PosX + Casilla_ancho / 2 + s * Casilla_ancho + p * BoxOffsetX
                            PosRelativaY = PosY + Casilla_alto / 4 - r * Casilla_alto - q * BoxOffsetY
                            'Seleccionando la fuente y tamaño:
                            
                            If frmMain.chk_BaseLetras.Value = True Then
                                Set shDATA_OBJECT1 = ActivePage.Layers("Layer4").CreateArtisticText(PosRelativaX, PosRelativaY, MyArray(NumBoxActual, NumCasillaActual), , , "DotumChe", TamañoLetras, cdrFalse, , , cdrCenterAlignment)
                                PosXprovisional = shDATA_OBJECT1.RotationCenterX
                                Set OrigSelection = ActiveSelectionRange
                                OrigSelection.SetOutlineProperties AnchoPlumaLetras ', OutlineStyles(0)
                                shDATA_OBJECT1.Stretch EstiramientoLetras, 1
                                shDATA_OBJECT1.Move (PosXprovisional - shDATA_OBJECT1.RotationCenterX + 0.1), 0#
                            Else
                                Set shDATA_OBJECT1 = ActivePage.Layers("Layer4").CreateArtisticText(PosRelativaX, PosRelativaY, MyArray(NumBoxActual, NumCasillaActual), , , "Century Schoolbook", 24, cdrTrue, , , cdrCenterAlignment)
                            End If
                            
                        End If
                    Next s
                Next r
                
                'Imprimiendo el Nº de Cartilla
                Set shDATA_OBJECT1 = ActivePage.Layers("Layer5").CreateArtisticText(PosX + 2.259 + p * BoxOffsetX, PosY + 11.994 - 0.05 - q * BoxOffsetY, "CART Nº " & NumBoxActual, , , "Century Schoolbook", 10, cdrTrue, , , cdrLeftAlignment)
                
                'Imprimiendo el Nº de Tiras
                If q = 0 Then
                    NumTira = x + NumPagTotal * p
                    Set shDATA_OBJECT1 = ActivePage.Layers("Layer6").CreateArtisticText(PosX - 2.649 + p * BoxOffsetX, PosY - 4.825, NumTira, , , "Century Schoolbook", 12, cdrTrue, , , cdrCenterAlignment)
                    shDATA_OBJECT1.Rotate 90
                End If
                
            Next q
        Next p
        
    Next x
    
    ActivePage.Layers("Layer4").Editable = False
    ActivePage.Layers("Layer5").Editable = False
    ActivePage.Layers("Layer6").Editable = False
    
    SampleINIRespaldoBase
    
End Sub

Sub SampleINIRespaldoBase()

    Dim sINI_FILE As String
    Dim sReturn As String

    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"

    ' Write to the ini file
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "UbicacionBD", sINI_FILE, frmMain.txt_UbicacionBD.Value)
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "UbicacionIcono", sINI_FILE, frmMain.txt_UbicacionIcono.Value)
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "Contactanos", sINI_FILE, frmMain.txt_ContactanosBase.Value)
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "PosInicialX", sINI_FILE, frmMain.txt_PosInicialX.Value)
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "PosInicialY", sINI_FILE, frmMain.txt_PosInicialY.Value)
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "BoxOffsetX", sINI_FILE, frmMain.txt_BoxOffsetX.Value)
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "BoxOffsetY", sINI_FILE, frmMain.txt_BoxOffsetY.Value)
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "PaginaInicial", sINI_FILE, frmMain.txt_PaginaInicial.Value)
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "PaginaFinal", sINI_FILE, frmMain.txt_PaginaFinal.Value)
    
    If frmMain.chk_BaseLetras.Value = True Then
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "BaseLetras", sINI_FILE, "True")
    Else
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "BaseLetras", sINI_FILE, "False")
    End If
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "FontLetras", sINI_FILE, frmMain.txt_FontLetras.Value)
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "TamañoLetras", sINI_FILE, frmMain.txt_TamañoLetras.Value)
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "AnchoPlumaLetras", sINI_FILE, frmMain.txt_AnchoPlumaLetras.Value)
    sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "EstiramientoLetras", sINI_FILE, frmMain.txt_EstiramientoLetras.Value)

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
    
    'Ubicacion de BD
    UbicacionBD = frmMain.txt_UbicacionBD.Value
    If UbicacionBD = "" Then
        MsgBox Prompt:="Debe indicar una ubicación de Archivo BD válida.", Buttons:=vbOKOnly, Title:="Validación del campo Ubicación BD"
        End
    ElseIf Len(Dir(UbicacionBD)) = 0 Then
        MsgBox Prompt:="El Archivo BD no se encontró en la ruta indicada.", Buttons:=vbOKOnly, Title:="Validación del campo Ubicación BD"
        End
    End If
    
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
    
    'Rango de Paginas a generar por bloque
    If frmMain.txt_PaginaInicial.Value = "" Or frmMain.txt_PaginaFinal.Value = "" Then
        MsgBox Prompt:="Debe revisar el rango de paginas elegido.", Buttons:=vbOKOnly
        End
    Else
        NumPagInicio = Val(frmMain.txt_PaginaInicial.Value)
        NumPagFinal = Val(frmMain.txt_PaginaFinal.Value)
    End If
    
    'Rango de Paginas a generar por Documento
    If frmMain.txt_DocPaginaInicial.Value = "" Or frmMain.txt_DocPaginaFinal.Value = "" Then
        MsgBox Prompt:="Debe revisar el rango de paginas elegido.", Buttons:=vbOKOnly
        End
    Else
        DocPagInicio = Val(frmMain.txt_DocPaginaInicial.Value)
        DocPagFinal = Val(frmMain.txt_DocPaginaFinal.Value)
    End If
    
     
    ActiveDocument.Unit = cdrMillimeter
    
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
    
    '***************************************************************
    'Llenar Mi Array
    '***************************************************************

    'LLENAR MyArray() CON INFO DEL CSV
    Dim i, j As Integer       'Integer va hasta 32,767
    Dim TXT As String
    Dim LineArray As Variant

    'Directory Address. A valid file number in the range 1 to 255
    Open UbicacionBD For Input As #1
    'Recorre cada fila del CSV
    For i = 0 To NumRowsCSV - 1
        'Read line into variable.
        Line Input #1, TXT
        'Remove linebreak from TXT
        TXT = Replace(TXT, vbLf, "")
        TXT = Replace(TXT, vbCrLf, "")
        'Split TXT into a 1d Array
        LineArray = Split(TXT, ",")

        'Redimensionar MyArray() en funcion al total de Filas y Columnas
        If i = 0 Then
            NumColsCSV = UBound(LineArray) + 1
            ReDim MyArray(NumRowsCSV - 1, NumColsCSV - 1)
        End If
            
        'Saving LineArray into MyArray
        For j = 0 To NumColsCSV - 1 'or UBound(LineArray)
            MyArray(i, j) = LineArray(j)
        Next j
    Next i
    'Closing file
    Close #1

    '***************************************************************
    'Creando los Iconos en las cartillas
    '***************************************************************

    'Calculando la cantidad de Cartillas totales en una sola columna
    NumRowsCSVxColsA3 = (NumRowsCSV - NumSillyRows) / NumColsA3
    'Calculando la cantidad de PaginasA3
    NumPagTotal = NumRowsCSVxColsA3 / NumRowsA3
       
    'Se recorrerán las X páginas del Documento
    For x = DocPagInicio To DocPagFinal
        ActiveDocument.Pages(x).Activate
        ActivePage.Layers("Layer4A").Activate
        'Se recorreran las 4 columnas por cada hoja A3
        For p = 0 To NumColsA3 - 1
            'Se recorrerán los 6 filas por cada Tira
            For q = 0 To NumRowsA3 - 1
                'Imprimiendo la casilla del Box
                NumBoxActual = NumSillyRows + q + NumRowsA3 * (x + NumPagInicio - 2) + p * NumRowsCSVxColsA3
                'Se recorrera cada Fila por cada Box
                For r = 0 To NumRowsBox - 1
                    'Se recorrerá cada Columna por cada Fila
                    For s = 0 To NumColsBox - 1
                        NumCasillaActual = NumSillyCols + s + NumColsBox * r
                        'Imprime ICONO solo si el valor de la casilla es igual a "" y si Ruta existe
                        If MyArray(NumBoxActual, NumCasillaActual) = "" And Len(Dir(UbicacionIcono)) <> 0 Then
                            'Posicion relativa (x,y)
                            PosRelativaX = PosX + 0.35 + s * Casilla_ancho + p * BoxOffsetX
                            PosRelativaY = PosY - 0.35 + Casilla_alto - r * Casilla_alto - q * BoxOffsetY
                            'Seleccionando capa , importando y ubicando el icono
                            ActiveDocument.ActivePage.Layers("Layer4A").Activate
                            ActiveDocument.ActiveLayer.Import (UbicacionIcono)
                            ActiveSelection.SetPosition PosRelativaX, PosRelativaY
                            
                            'PosRelativaX = PosX + s * Casilla_ancho + p * BoxOffsetX
                            'PosRelativaY = PosY - r * Casilla_alto - q * BoxOffsetY
                            'Set shDATA_OBJECT1 = ActivePage.Layers("Layer4A").CreateRectangle2(PosRelativaX, PosRelativaY, Casilla_ancho, Casilla_alto)
                        End If
                    Next s
                Next r
            Next q
        Next p
    Next x
    
    ActivePage.Layers("Layer4A").Editable = False
    
    SampleINIRespaldoBase
            
End Sub

Sub BorrarContenidoIconos()
    Dim l As Layer, lr As Layers
    Dim i As Integer
    
    For i = ActiveDocument.Pages.Count To 1 Step -1
        ActiveDocument.Pages(i).Activate
        ActivePage.Layers("Layer4A").Shapes.All.Delete
    Next i
    
End Sub

Sub CrearFondoTiraA3()
    Dim shDATA_OBJECT1 As Shape
    Dim i As Integer
    Dim sr As ShapeRange

    'Se crea nueva pagina Retira y se agrega un cuadrado de fondo A3 de color blanco
    ActiveDocument.Pages(1).Activate
    Set shDATA_OBJECT1 = ActivePage.Layers("Layer0").CreateRectangle2(ActivePage.LeftX, ActivePage.BottomY, ActivePage.SizeWidth, ActivePage.SizeHeight)
    shDATA_OBJECT1.Fill.UniformColor.CMYKAssign 0, 0, 0, 0
    
    i = 0
    'Se imprime lineas divisorias Verticales en la hoja A3
    Do While (i * 1.542) < ActivePage.SizeWidth
        Set shDATA_OBJECT1 = ActivePage.Layers("Layer0A").CreateLineSegment(ActivePage.LeftX + i * 1.542, ActivePage.BottomY, ActivePage.LeftX + i * 1.542, ActivePage.BottomY + ActivePage.SizeHeight)
        shDATA_OBJECT1.Outline.SetProperties 0.25
        i = i + 1
    Loop
    
    Set sr = ActivePage.Layers("Layer0A").Shapes.All
    sr.SetOutlineProperties Color:=CreateCMYKColor(0, 0, 0, 0)
     
End Sub
