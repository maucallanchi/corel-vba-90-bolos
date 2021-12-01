Attribute VB_Name = "mdmProgramQR"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function URLDownloadToCacheFile Lib "urlmon" Alias _
 "URLDownloadToCacheFileA" (ByVal lpUnkcaller As Long, ByVal szURL As String, _
 ByVal szFileName As String, ByVal dwBufLength As Long, ByVal dwReserved As Long, _
 ByVal IBindStatusCallback As Long) As Long
#Else
Private Declare Function URLDownloadToCacheFile Lib "urlmon" Alias _
 "URLDownloadToCacheFileA" (ByVal lpUnkcaller As Long, ByVal szURL As String, _
 ByVal szFileName As String, ByVal dwBufLength As Long, ByVal dwReserved As Long, _
 ByVal IBindStatusCallback As Long) As Long
#End If

Public Function DownloadFile(URL As String) As String
    Dim szFileName As String
    szFileName = Space$(300)
    If URLDownloadToCacheFile(0, URL, szFileName, Len(szFileName), 0, 0) = 0 Then
        DownloadFile = Trim(szFileName)
    End If
End Function

Private Sub QR_Insert()         'Ejemplo
    Dim strFile_Path As String
    Dim strQRText As String
    
    strQRText = "000010003240"
        
    'DownloadFile(InternetPath /?data= QRText [&size=MaxResolutionxMaxResolution][&color=RGB][ecc=L,M,Q or H][margin=surround area][&format=OutputFormat])
    'Do not use jpg or jpeg as the result is poor. The default is png.
    'MaxResolution in pixel must be the same in both directions. The default is 200x200
    'Possible output formats and they are case sensitive are png, gif, jpeg, jpg, svg or eps.
    'color & bgcolor specify the color & background colors. They are in RGB either as Range-g-b or hex values of either 3 or 6 characters.
    'The default colors are black & white.
    'ecc Error Corection Code specifies the degree of redundancy if the QR code is damaged. The default is L.
    'L Low 7% destroyed data may be corrected.
    'M Middle 15% destroyed data may be corrected.
    'Q Quality 25% destroyed data may be corrected.
    'H High 30% destroyed data may be corrected.
    'margin sets the No of pixels (space) around the barcode. It varies from 0 to 50 The default is 1. It is the same color as bgcolor.
    'The margin is ignored for eps & svg barcodes.
    'qzone seems to be the same as margin but works with all formats eps & png barcodes. qzone = 0 - 100. The default qzone = 0
    'strFile_Path = DownloadFile("http://api.qrserver.com/v1/create-qr-code/?data=" & strQRText & "&size=100x100&format=png")
    'strFile_Path = DownloadFile("http://api.qrserver.com/v1/create-qr-code/?data=" & strQRText & "&format=eps")
    strFile_Path = DownloadFile("http://api.qrserver.com/v1/create-qr-code/?data=" & strQRText & "&ecc=H&size=100x100&format=png")
    
    'If the program does not obtain a QR barcode exit.
    If strFile_Path = "" Then
        MsgBox "There is a problem." & vbCr & "This computer may not be able to connect with the internet or" & _
         vbCr & "The server at http://api.qrserver.com may be down."
        Exit Sub
    End If
    
    'To insert into Word.
    'ActiveDocument.Shapes.AddPicture strFile_Path
    
    'To insert into Excel
    'Cells(3, 4).Select
    'Place picture such that its top left corner is cells(3,4)
    'ActiveWorkbook.ActiveSheet.Pictures.Insert strFile_Path

    'To insert into CorelDraw
    ActiveDocument.ActiveLayer.Import (strFile_Path)
End Sub

