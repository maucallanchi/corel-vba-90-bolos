VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Control de Pre-Impresion de Cartillas de 90 bolos V3.22"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4980
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chk_BaseLetras_Click()
    If chk_BaseLetras.Value = True Then
        txt_FontLetras.Visible = True
        lbl_FontLetras.Visible = True
        txt_TamañoLetras.Visible = True
        lbl_TamañoLetras.Visible = True
        txt_AnchoPlumaLetras.Visible = True
        lbl_AnchoPlumaLetras.Visible = True
        txt_EstiramientoLetras.Visible = True
        lbl_EstiramientoLetras.Visible = True
    Else
        txt_FontLetras.Visible = False
        lbl_FontLetras.Visible = False
        txt_TamañoLetras.Visible = False
        lbl_TamañoLetras.Visible = False
        txt_AnchoPlumaLetras.Visible = False
        lbl_AnchoPlumaLetras.Visible = False
        txt_EstiramientoLetras.Visible = False
        lbl_EstiramientoLetras.Visible = False
    End If
End Sub

Private Sub cmdAgregarIcono_Click()
    ActiveDocument.BeginCommandGroup
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    mdmRutinasOpcionales.AgregarIcono
    
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    Application.CorelScript.RedrawScreen
    
    MsgBox Prompt:="Se agregaron los iconos exitosamente", Buttons:=vbOKOnly
End Sub

Private Sub cmdAplicarTextoQR_Click()
    mdmRutinasOpcionales.AgregarTextoQR
End Sub
Private Sub cmdCambiarContactanos_Click()
    mdmRutinasOpcionales.CambiarTextoContactanos
End Sub
Private Sub cmdFondoA3_Click()
    mdmRutinasOpcionales.AgregarFondoA3
End Sub
Private Sub cmdFondoBox_Click()
    ActiveDocument.BeginCommandGroup
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    mdmRutinasOpcionales.AgregarFondoBox
    
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    Application.CorelScript.RedrawScreen
End Sub
Private Sub cmdMostrarQRRetira_Click()
    mdmRutinasObligatorias.MostrarQRRetira
End Sub
Private Sub cmdMostrarQRTira_Click()
    mdmRutinasObligatorias.MostrarQRTira
End Sub

Private Sub cmdMostrarCapas_Click()
    ActiveDocument.BeginCommandGroup
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    mdmRutinasOpcionales.MostrarCapas
    
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    Application.CorelScript.RedrawScreen
    
    Beep
    MsgBox Prompt:="Se activaron las capas", Buttons:=vbExclamation
    
End Sub

Private Sub cmdNivelNegro_Click()
    ActiveDocument.BeginCommandGroup
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    mdmRutinasOpcionales.CambiarNivelColor
    
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    'Application.CorelScript.RedrawScreen
    
    Beep
    MsgBox Prompt:="Se aplico el %negro a todas las capas", Buttons:=vbExclamation
End Sub

Private Sub cmdNivelNegroCodificacion_Click()
    ActiveDocument.BeginCommandGroup
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    mdmRutinasObligatorias.CambiarNivelColorSoloCodificacion
    
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    Application.CorelScript.RedrawScreen
    
    Beep
    MsgBox Prompt:="Se aplico el Color a las capas de Codificación", Buttons:=vbExclamation
End Sub

Private Sub cmdOcultarCapas_Click()
    ActiveDocument.BeginCommandGroup
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    mdmRutinasOpcionales.OcultarCapas
    
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    Application.CorelScript.RedrawScreen
    
    Beep
    MsgBox Prompt:="Se activaron las capas", Buttons:=vbExclamation
End Sub

Private Sub cmdRestablecer_Click()
    mdmRutinasObligatorias.BorrarContenidoCapasObligatoriasPrevias
End Sub
Private Sub cmdColapsarExpandir_Click()
    If frmMain.Height = 278.25 Then
        frmMain.Height = 48
    Else
        frmMain.Height = 278.25
    End If
End Sub
Private Sub cmdGenerarBloquesBase_Click()
    mdmRutinasBase1.ReiniciarPaginasyCapas
    
    ActiveDocument.BeginCommandGroup
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    mdmRutinasBase1.GenerarBloquesBase
    
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    'Application.CorelScript.RedrawScreen
    
    MsgBox Prompt:="Se generó el bloque exitosamente", Buttons:=vbOKOnly
End Sub
Private Sub cmdInvertirOrden_Click()
    mdmRutinasOpcionales.InvertirOrdenPaginas
    MsgBox Prompt:="Se cambio el orden de todas las páginas", Buttons:=vbOKOnly
End Sub
Private Sub cmdReiniciarTodo_Click()
    mdmRutinasBase1.ReiniciarPaginasyCapas
    MsgBox Prompt:="Se borraron todas las páginas", Buttons:=vbOKOnly
End Sub
Private Sub cmdAplicarSeleccion_Click()
    ActiveDocument.BeginCommandGroup
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    mdmRutinasObligatorias.AplicarCambiosObligatorios
    
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    'Application.CorelScript.RedrawScreen
End Sub
Private Sub cmdTramadoA3_Click()
    ActiveDocument.BeginCommandGroup
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False

    mdmRutinasOpcionales.AgregarTramadoA3
    
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    'Application.CorelScript.RedrawScreen
End Sub

Private Sub cmdBorrarIconos_Click()
    mdmRutinasOpcionales.BorrarContenidoIconos
    MsgBox Prompt:="Se borraron todos los iconos", Buttons:=vbOKOnly
End Sub


Private Sub mdmColorNumCart_Click()
    mdmRutinasOpcionales.CambiarColorNumCart
End Sub


Private Sub UserForm_Initialize()
    Dim sINI_FILE As String
    Dim sReturn As String
    
    'Se agregan opciones al ComboBox "Codificacion"
    cbo_Codificacion.AddItem "EAN13"
    cbo_Codificacion.AddItem "Code128"
    cbo_Codificacion.AddItem "QRCode"
    
    'Valida si el archivo Corel ya está guardado
    If Len(ActiveDocument.FileName) = 0 Then
        MsgBox Prompt:="Debe darle un nombre al Archivo Corel antes de usar el Programa.", Buttons:=vbOKOnly, Title:="Validación Nombre Archivo Corel"
        End
    End If
    
    'Crea ruta y nombre de archivo .INI
    sINI_FILE = ActiveDocument.FilePath & Left(ActiveDocument.FileName, Len(ActiveDocument.FileName) - 4) & ".ini"
    
    If Len(Dir(sINI_FILE)) = 0 Then
        'File Does Not exist
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "UbicacionBD", sINI_FILE, "C:\Users\Marlon\Documents\BGC18000.csv")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "Contactanos", sINI_FILE, "AMAE (+51) 955 621 606")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "PosInicialX", sINI_FILE, "12")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "PosInicialY", sINI_FILE, "275")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "BoxOffsetX", sINI_FILE, "106.23")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "BoxOffsetY", sINI_FILE, "48.1")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "PaginaInicial", sINI_FILE, "1")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "PaginaFinal", sINI_FILE, "2")
        
        sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "NumSerie", sINI_FILE, "0003240")
        sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "Codificacion", sINI_FILE, "EAN13")
        sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "NivelNegro", sINI_FILE, "100")
        sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "NivelNegroQR", sINI_FILE, "0,0,0,100")
        sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "grupoQR", sINI_FILE, "1")
        sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "CarpetaLocalQR", sINI_FILE, "C:\Users\Marlon\Documents\OUTPUT\")
        sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "BaseDatosQR", sINI_FILE, "C:\Users\Marlon\Documents\QR Puntos Bingo editado.txt")
        sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "AjustePosQR", sINI_FILE, "0,0")
        sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "AjusteTamañoQR", sINI_FILE, "0.6,0.6")
        
        sReturn = sManageSectionEntry(iniWrite, "Opcionales", "TextoQR", sINI_FILE, "Texto QR")
        sReturn = sManageSectionEntry(iniWrite, "Opcionales", "Contactanos", sINI_FILE, "AMAE PRUEBA")
        sReturn = sManageSectionEntry(iniWrite, "Opcionales", "NivelNegro", sINI_FILE, "100")
        sReturn = sManageSectionEntry(iniWrite, "Opcionales", "NivelNegroQR", sINI_FILE, "0,0,0,100")
        sReturn = sManageSectionEntry(iniWrite, "Opcionales", "FondoBox", sINI_FILE, "0,10,0,0")
        sReturn = sManageSectionEntry(iniWrite, "Opcionales", "FondoA3", sINI_FILE, "0,50,0,0")
        sReturn = sManageSectionEntry(iniWrite, "Opcionales", "TramadoA3", sINI_FILE, "0,50,0,0")
        sReturn = sManageSectionEntry(iniWrite, "Opcionales", "UbicacionIcono", sINI_FILE, "C:\Users\Marlon\Documents\DREAM_LOGO-BINGO_LETRAS.svg")
        sReturn = sManageSectionEntry(iniWrite, "Opcionales", "ColorNumCart", sINI_FILE, "0,0,0,100")
        
        
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "BaseLetras", sINI_FILE, "False")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "CompatibleIconos", sINI_FILE, "False")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "FontLetras", sINI_FILE, "DotumChe")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "TamañoLetras", sINI_FILE, "20")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "AnchoPlumaLetras", sINI_FILE, "0.18")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "EstiramientoLetras", sINI_FILE, "0.7")
        
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "FontNumeros", sINI_FILE, "Century Schoolbook")
        sReturn = sManageSectionEntry(iniWrite, "Modulo_Base", "TamanoNumeros", sINI_FILE, "24")
        
        sReturn = sManageSectionEntry(iniWrite, "Obligatorias", "FontNumSerie", sINI_FILE, "Arial")
        
    End If
    
    'Leyendo archivo INI y actualizando valores del Formulario
    frmMain.txt_UbicacionBD.Value = sManageSectionEntry(iniRead, "Modulo_Base", "UbicacionBD", sINI_FILE)
    frmMain.txt_ContactanosBase.Value = sManageSectionEntry(iniRead, "Modulo_Base", "Contactanos", sINI_FILE)
    frmMain.txt_PosInicialX.Value = sManageSectionEntry(iniRead, "Modulo_Base", "PosInicialX", sINI_FILE)
    frmMain.txt_PosInicialY.Value = sManageSectionEntry(iniRead, "Modulo_Base", "PosInicialY", sINI_FILE)
    frmMain.txt_BoxOffsetX.Value = sManageSectionEntry(iniRead, "Modulo_Base", "BoxOffsetX", sINI_FILE)
    frmMain.txt_BoxOffsetY.Value = sManageSectionEntry(iniRead, "Modulo_Base", "BoxOffsetY", sINI_FILE)
    frmMain.txt_PaginaInicial.Value = sManageSectionEntry(iniRead, "Modulo_Base", "PaginaInicial", sINI_FILE)
    frmMain.txt_PaginaFinal.Value = sManageSectionEntry(iniRead, "Modulo_Base", "PaginaFinal", sINI_FILE)
    
    frmMain.txt_NumSerie.Value = sManageSectionEntry(iniRead, "Obligatorias", "NumSerie", sINI_FILE)
    frmMain.cbo_Codificacion.Value = sManageSectionEntry(iniRead, "Obligatorias", "Codificacion", sINI_FILE)
    frmMain.txt_NivelNegroCodificacion.Value = sManageSectionEntry(iniRead, "Obligatorias", "NivelNegro", sINI_FILE)
    frmMain.txt_NivelNegroQRCodificacion.Value = sManageSectionEntry(iniRead, "Obligatorias", "NivelNegroQR", sINI_FILE)
    frmMain.txt_grupoQR.Value = sManageSectionEntry(iniRead, "Obligatorias", "grupoQR", sINI_FILE)
    frmMain.txt_CarpetaLocalQR.Value = sManageSectionEntry(iniRead, "Obligatorias", "CarpetaLocalQR", sINI_FILE)
    frmMain.txt_BaseDatosQR.Value = sManageSectionEntry(iniRead, "Obligatorias", "BaseDatosQR", sINI_FILE)
    frmMain.txt_AjustePosQR.Value = sManageSectionEntry(iniRead, "Obligatorias", "AjustePosQR", sINI_FILE)
    frmMain.txt_AjusteTamañoQR.Value = sManageSectionEntry(iniRead, "Obligatorias", "AjusteTamañoQR", sINI_FILE)
    
    frmMain.txt_TextoQR.Value = sManageSectionEntry(iniRead, "Opcionales", "TextoQR", sINI_FILE)
    frmMain.txt_Contactanos.Value = sManageSectionEntry(iniRead, "Opcionales", "Contactanos", sINI_FILE)
    frmMain.txt_NivelNegro.Value = sManageSectionEntry(iniRead, "Opcionales", "NivelNegro", sINI_FILE)
    frmMain.txt_NivelNegroQR.Value = sManageSectionEntry(iniRead, "Opcionales", "NivelNegroQR", sINI_FILE)
    frmMain.txt_FondoBox.Value = sManageSectionEntry(iniRead, "Opcionales", "FondoBox", sINI_FILE)
    frmMain.txt_FondoA3.Value = sManageSectionEntry(iniRead, "Opcionales", "FondoA3", sINI_FILE)
    frmMain.txt_TramadoA3.Value = sManageSectionEntry(iniRead, "Opcionales", "TramadoA3", sINI_FILE)
    frmMain.txt_UbicacionIcono.Value = sManageSectionEntry(iniRead, "Opcionales", "UbicacionIcono", sINI_FILE)
    frmMain.txt_ColorNumCart.Value = sManageSectionEntry(iniRead, "Opcionales", "ColorNumCart", sINI_FILE)
    
    frmMain.chk_BaseLetras.Value = sManageSectionEntry(iniRead, "Modulo_Base", "BaseLetras", sINI_FILE)
    frmMain.chk_CompatibilidadIconos.Value = sManageSectionEntry(iniRead, "Modulo_Base", "CompatibleIconos", sINI_FILE)
    frmMain.txt_FontLetras.Value = sManageSectionEntry(iniRead, "Modulo_Base", "FontLetras", sINI_FILE)
    frmMain.txt_TamañoLetras.Value = sManageSectionEntry(iniRead, "Modulo_Base", "TamañoLetras", sINI_FILE)
    frmMain.txt_AnchoPlumaLetras.Value = sManageSectionEntry(iniRead, "Modulo_Base", "AnchoPlumaLetras", sINI_FILE)
    frmMain.txt_EstiramientoLetras.Value = sManageSectionEntry(iniRead, "Modulo_Base", "EstiramientoLetras", sINI_FILE)
    
    frmMain.txt_FontNumeros.Value = sManageSectionEntry(iniRead, "Modulo_Base", "FontNumeros", sINI_FILE)
    frmMain.txt_TamanoNumeros.Value = sManageSectionEntry(iniRead, "Modulo_Base", "TamanoNumeros", sINI_FILE)
    
    frmMain.txt_FontNumSerie.Value = sManageSectionEntry(iniRead, "Obligatorias", "FontNumSerie", sINI_FILE)
    
    
End Sub
