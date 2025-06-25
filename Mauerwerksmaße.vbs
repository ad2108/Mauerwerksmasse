' --------------------------------------------------
' Mauerwerksmaße 
' --------------------------------------------------

' Author: ad2108
' Version: 1.0
' Date: 2025-06-04
' License: MIT

' Description:
'     Berechnung der Mauerwerksmaße

' --------------------------------------------------
' Funktionen der Mauerwerksmaße

' Außenmaß
Public Function Aussenmass(x)
  Aussenmass = x * 12.5 - 1
End Function

' Öffnungsmaß
Public Function Oeffnungsmass(x)
  Oeffnungsmass = x * 12.5 + 1
End Function

' Vorsprungsmaß
Public Function Vorsprungsmass(x)
  Vorsprungsmass = x * 12.5
End Function

' --------------------------------------------------
' Variable mit Fensternamen
MsgBox_Name = "Mauerwerksmasse"

' --------------------------------------------------
' Eingabe der Option

Input_Option = InputBox("1. Aussenmass A" & vbCrLf & "2. Oeffnungsmass O" &  vbCrLf & "3. Vorsprungmass V", MsgBox_Name)

' --------------------------------------------------
' Festlegung des weiteren Fensternamens

' Außenmaß
If Input_Option = 1 Then
  MsgBox_Name = "Aussenmass"

' Öffnungsmaß
ElseIf Input_Option = 2 Then  
  MsgBox_Name = "Oeffnungsmass"

' Vorsprungsmaß
ElseIf Input_Option = 3 Then
  MsgBox_Name = "Vorsprungsmass"
End If

' --------------------------------------------------
' ##################################################
' Programmschleife

Do While True:

' --------------------------------------------------
' Eingabe der Anzahl der Achtelmaße (Steine) und
' Konvertierung der Zahl in eine Ganzzahl

Input_Anzahl_Achtelmass = InputBox("Anzahl der Achtelmasse (Steine)", MsgBox_Name)
Input_Anzahl_Achtelmass = Int(Input_Anzahl_Achtelmass)

' --------------------------------------------------
' Berechnung des Maßes

' Außenmaß
If Input_Option = 1 Then
  Result = Aussenmass(Input_Anzahl_Achtelmass)

' Öffnungsmaße
ElseIf Input_Option = 2 Then
  Result = Oeffnungsmass(Input_Anzahl_Achtelmass)

' Vorsprungsmaß
ElseIf Input_Option = 3 Then
  Result = Vorsprungsmass(Input_Anzahl_Achtelmass)
End If

' --------------------------------------------------
' Umrechnung zu Kilometer, Meter, Zentimeter

' Falls kleiner als 1 m wird das Maß in cm angegeben
If Result < 100 Then
  Result = Result & " cm"

' Falls größer als 1 m dann wird das Maß in m angegeben
ElseIf Result < 100*1000 Then
  Result = Result / 100 & " m"

' Falls größer als 1000 m dann wird das Maß in km angegeben
ElseIf Result >= 100*1000 Then
  Result = Result / 100 / 1000 & " km"
End If

' --------------------------------------------------
' Ausgabe des Maßes und Anfrage auf Wiederholung der Berechnung

Loop_it = MsgBox(Result, vbRetryCancel , MsgBox_Name)

' --------------------------------------------------
' Falls nicht Retry angegeben wird, wird die Schleife verlassen

If Loop_it <> vbRetry Then
  Exit Do
End If

' --------------------------------------------------
' ##################################################
' End der Programmschleife

Loop

' --------------------------------------------------
' End of file
' --------------------------------------------------

