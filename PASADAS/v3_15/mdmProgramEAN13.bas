Attribute VB_Name = "mdmProgramEAN13"
Option Explicit

Public Function EAN_13(BARCODE_No As String)
  'The function name and its augument have been changed to more relevant english names
  'plus the function now discards the 13th character in BARCODE_No.
  'Otherwise the function is as originally written.
  'You can download the latest source files from "barcode fonts and encoders" found at SourceForge,
  'http://sourceforge.net/project/showfiles.php?group_id=120100

'Copyright (C) 2006 (Grandzebu)
'These programs and the fonts which are supplied with it are free, you can redistribute it and/or
'modify it under the terms of the GNU General Public License as published by the
'Free Software Foundation either version 2 of the License, or (at your option) any later version.
'The barcode encoding functions, source code and library, are governed by the
'GNU Lesser General Public License (GNU LGPL)
'These programs are distributed in the hope that they will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General
'Public License for more details.
'Please download a license copy at : http://www.gnu.org/licenses/gpl.html
  
  
  'V 1.0
  'Paramètres : une chaine de 12 chiffres
  'Retour : * une chaine qui, affichée avec la police EAN13.TTF, donne le code barre
  '         * une chaine vide si paramètre fourni incorrect
  Dim i%, checksum%, first%, CodeBarre$, tableA As Boolean
  EAN_13 = ""
  
  'Only accept the first 12 digits so as to ignore the 13th digit,
  'which is the check digit, if it is supplied.
  BARCODE_No = Left(BARCODE_No, 12)
  
  'Verify BARCODE_No has 12 characters
  If Len(BARCODE_No) = 12 Then
    'Et que ce sont bien des chiffres
    For i% = 1 To 12
      If Asc(Mid$(BARCODE_No, i%, 1)) < 48 Or Asc(Mid$(BARCODE_No, i%, 1)) > 57 Then
        i% = 0
        Exit For
      End If
    Next
    If i% = 13 Then
      'Calculate de la clé de contrôle
      For i% = 2 To 12 Step 2
        checksum% = checksum% + Val(Mid$(BARCODE_No, i%, 1))
      Next
      checksum% = checksum% * 3
      For i% = 1 To 11 Step 2
        checksum% = checksum% + Val(Mid$(BARCODE_No, i%, 1))
      Next
      BARCODE_No = BARCODE_No & (10 - checksum% Mod 10) Mod 10
      'Le premier chiffre est pris tel quel, le deuxième vient de la table A
      CodeBarre$ = Left$(BARCODE_No, 1) & Chr$(65 + Val(Mid$(BARCODE_No, 2, 1)))
      first% = Val(Left$(BARCODE_No, 1))
      For i% = 3 To 7
        tableA = False
         Select Case i%
         Case 3
           Select Case first%
           Case 0 To 3
             tableA = True
           End Select
         Case 4
           Select Case first%
           Case 0, 4, 7, 8
             tableA = True
           End Select
         Case 5
           Select Case first%
           Case 0, 1, 4, 5, 9
             tableA = True
           End Select
         Case 6
           Select Case first%
           Case 0, 2, 5, 6, 7
             tableA = True
           End Select
         Case 7
           Select Case first%
           Case 0, 3, 6, 8, 9
             tableA = True
           End Select
         End Select
       If tableA Then
         CodeBarre$ = CodeBarre$ & Chr$(65 + Val(Mid$(BARCODE_No, i%, 1)))
       Else
         CodeBarre$ = CodeBarre$ & Chr$(75 + Val(Mid$(BARCODE_No, i%, 1)))
       End If
     Next
      CodeBarre$ = CodeBarre$ & "*"   'Ajout séparateur central
      For i% = 8 To 13
        CodeBarre$ = CodeBarre$ & Chr$(97 + Val(Mid$(BARCODE_No, i%, 1)))
      Next
      CodeBarre$ = CodeBarre$ & "+"   'Ajout de la marque de fin
      EAN_13 = CodeBarre$
    End If
  End If
End Function
