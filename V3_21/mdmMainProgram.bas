Attribute VB_Name = "mdmMainProgram"
Sub MainProgram()
    frmMain.Show vbModeless
End Sub

'Release Notes
' Version 3.21:
' Se corrige detalles sobre chk_CompatibilidadIconos y chk_BaseLetras

' Version 3.20:
' Se agrega rutinas Opcionales para Mostrar y Ocultar las capas Layer4B, Layer5 y Layer6 + Layer1, Layer2 y Layer4C.
' Se agrega Botones al formulario con opci{on de Optimizacion de procesamiento.

' Version 3.19:
' Se mueve NUmTira 15mm a la derecha en la Retira QR
' Se agrega campo en Formularo llamado GrupoQR que permite indicar que bloque usar dentro del catalogo de QR de acuerdo con la BD de cartillas usada.

' Version 3.18:
' Ajuste de formulario + nuevos campos de texto para .ini en pestaña Codificacion
' Se cambia size de NumTira de 12 a 10 en Tira y Retira
' Se crea rutina DetectarBaseDatos y funcion BuscarRegistroBaseDatosQR

' Versión 3.17:
' Se agrega un Ajuste de posicion para imagenes QR y Ajuste de Tamaño del PNG.
' Se cambia funcionalidad de rutina GeneraRetiraQRenA3 para importe las imagenes QR de fuente local y ya no de internet

' Versión 3.16:
' Se cambia rutina CambiarNivelColorNegro por CambiarNivelColor
' Se cambia en el formulario la entrada de color negro QR para ahora este en formato CYMK
' Versi{on 3.15:
' Se agrega texto y boton nuevos en Formulario para controlar el color de los numeros de cartilla
' Se crea nueva subrutina opcional llamada CambiarColorNUmCart.

' Versión 3.14:
' Se agrega opcion en formulario Compatibilidad Iconos.
' Se agregan instrucciones de Optimizacion a los botones más importantes del FORM para aumentar velocidad de Procesamiento
' Se agrega condicional Compatibilidad Iconos en GenerarBloquesBase
' Se agrega propiedad cdrMilimiters a AgregarTramadoA3
' Se actualiza SampleINIRespaldoObligatorias para agregar campo NivelNegro y NivelNegroQR
' Se actualiza CrearTextoNumSerieEnA3, CrearCodificacionEnA3 y GeneraRetiraQRenA3 para que asigne color negro segun Formulario
' Se mueven AgregarIcono y BorrarContenidoIconos al Modulo mdmRutinasOpcionales
' Se agrega un condiciona sr.count a AgregarFondoBox
'
' Version 3.13:
' Se edita CrearTextoNumSerieEnA3 para que asigne colo negro al %75
' Version 3.12:
' Se eleva en 0.5mm la ubicación del codigo de barras.
' Se reduce el fondoBox los 7mm que se agregaron en versión anterior.
' Se ajusta CambiarNivelColorNegro para evitar llamar muchas veces a l.active.x
' Se agrega validacion de previa que cierra programa si NumPagFinal > NumPagTotal
' Se agrega al formulario nuevos campos para modificar solo el 5negro en capas Obligatorias
' Se agrega nueva subrutina CambiarNivelColorNegroSoloCodificacion en Rutinas Obligatorias
' Version 3.11:
' Se corrige bug en AgregarFondoBox. Ahora el color tambien se agrega a Layer4A.
' Se extiende cuadrado de fondoBox en 7mm para que se aprecie mejor el NCart, Codificacion y NumSerie cuando se usa FondoA3
' Version 3.10:
' Se borran algunos cajas de texto del formulario/Opcionales/Rango de paginas para Iconos y se reorganiza botones
' En ReiniciarPaginasyCapas la condicion solo consulta si es una capa especial. Segun eso borra todas las capas incluidas las master.
' Se reestructura AgregarFondoA3 y se elimina AgregarFondoTiraA3
' Se reestructura AgregarTramadoA3
' Se reestructura completamente CambiarNivelColorNegro. Ahora cuenta con optimización.
' Version 3.9:
' Esta es una version BETA/ Actualización: ya fue validado por el cliente.
' Se agregan nuevas capas y se fusionan varias subrutinas en una sola RutinaBase1
' Se agrega nueva capa master para el manejo exlusivo de iconos
' Version 3.8:
' Se corrige subrutina ObtenerNumeroFilasCSV. El numero de filas se estaba disminuyendo en 1 innecesariamente.
' Se corrige subrtuina ReiniciarPaginasyCapas para que la condicion sea l.IsSpecialLayer = False
' Se elimina subrutina ComprimirLetras y se incorpora su contenido en CrearPaginasConCartillas
' Se crea funcion BorrarContenidoIconos y se agrega boton al formulario
' Compatibilidad para VBA7 y VBA6 (Corel X3 y X7): Se agrega Condicional para usar o no la funcion PtrSafe.
' Version 3.7:
' En la subrutina CrearPaginasConCartillas se QUITA la opcion de cargar imagen y se mueve a la Subrutina AgregarIcono
' Se adecua el Formulario para que incluya boton "Agregar Icono" y se mueve campo Ubicación Icon
' En CrearCodificacionEnA3 se escala el texto de la Codificación EAN13 de 0.5 a 0.4 y se baja 1.5mm
' Version3.6:
' Se agregan atributos opcionales en el formulario / Bloques Base relacionadas al Bingo de Letras
' Se agregan atributos al archivo .INI
' Se traslada Subrutina ComprimirLetras al Modulo mdmRutinaBase
' Correccion en la subrtuina CrearFondoTiraA3 para que las lineas verticales no sobrepasen la hoja A3
' Version3.5:
' Se crea la capa Layer4A exclusivamente para los iconos
' Se recomienda usar el formato .svg para el icono y exportarlo con 9.3mm de ancho
' Version3.4:
' En la subrutina CrearPaginasConCartillas se agrega la opcion de cargar imagen en las casillas que están en blanco
' Se reordena el formulario sección Bloques Base
' La subrutina ReiniciarPaginas ahora incluye un borrado de los campos de texto opcional FondoA3, FondoBox y TramadoA3
' Se actualiza la subrutina IniciarVariables para que incluya el campo UbicacionIcono
' Version3.3:
' Se crea la subrutina GeneraCSVconArrayDeLetras usanddo como fuente las Base de Datos de 18000 cartillas.
' Se crea la subrutina ComprimirLetras() que comprime el tamaño de las letras a 10mm de ancho
' Version3.2:
' Se corrige funcion CambiarNivelColorNegro() para que actualice la capa Layer13 (imagenes QR) con un color y el resto con otro color
' Version3.1:
' Se corrige el grosor de las lineas del Box a 0.25mm
' Se sube 2mm la posición Y de las imagenes QR
' Se agregar capa Layer0A para las lineas verticales del tramado A3
' Se agrega control de color del tramado A3 al Formulario
' Version 3:
' Se simplifica la SubRutina CambiarTextoContactanos. En lugar de crear nuevo texto se reemplaza contenido del texto existente.
' Se corrige la Subrutina CambiarNivelColorNegro para que aplique nuevo color tambien para Layer1.
' Se simplifica la Subrutina CrearBoxesEnA3 para que en lugar de casillas existan lineas divisorias
' Se cambia la Subrutina AgregarFondoBox para que solo aplique color al marco del Box
' Se cambian las subrutinas obligatorias relacionadas al QR (CrearTextoNumSerieMasCodificacionEnA3) para que se genere una hoja A3 Retira con la info de los QRs
' Se crea nuevo modulo mdmSampleINI con la funcion para leer y guardar campos desde un archivo INI. De esta forma se podran recordar los valores.
' Se cambian las subrutinas del Formulario y Modulos Base|Obligatorios|Opcionales para guarden/lean el archivo INI
' La capa Layer3 que contenia los textos de NumSerie ahora es MasterPage
' Se agrega un fondo A3 de lineas verticales tipo tramado
' Se independizan en el Formulario el color negro para textos/formas del color negro para imagenes QR.
' Version 2.3:
' Se aumenta en 1mm la distancia entre el Texto Contactanos y el Box
' Se agrega borrado de color de fondo en MasterPage
' Version 2.2:
' Se agrega un marco a cada Box de grosor 0.5mm
' Version 2.1:
' Se agrega grosor a los cuadros del Box en 0.25mm
' Se reduce en 1mm la distancia entre el Texto Contactanos y el Box
' Se reduce el tamaño de letra del texto TIRA de 16 a 14 puntos.
' Version 2:
' Nueva interfaz gráfica
' Opcion Colapsar/Expandir ventana /OK
' Generacion de todas las 750 paginas en Bloques de 150pag
' Texto editable en el espacio del codebar
' QR en retira
' Opcion de Invertir Paginas
' Opcion de Color y diseño editable en la Tira
' Opcion de Color y Diseño editable en la Retira
' Ubicacion BD
' Opci{on de color en textos


