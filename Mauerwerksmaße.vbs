' --------------------------------------------------
' Mauerwerksmaße 
' --------------------------------------------------

' Author: ad2108
' Version: 1.0
' Date: 2025-05-08
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
' Anfang der Ausführung (Abbildung öffnen)

' ShellObjekt zuweisen
Set objShell = CreateObject("Wscript.Shell")

' Abfrage ob Abbildung angezeigt werden soll
Bild_Anzeigen = MsgBox("Abbildung oeffnen?", vbYesNo, MsgBox_Name)

' Falls Ja, wird über die Shell die Abbildung geöffnet
If Bild_Anzeigen = vbYes Then
  objShell.Run("https://i.pinimg.com/originals/1f/16/79/1f1679c01186000db88932e787d177d3.png")
End If

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
' Programschleife

Do While True:

' --------------------------------------------------
' Eingabe der Anzahl der Steine und
' Konvertierung der Zahl in eine Ganzzahl

Input_Anzahl_Steine = InputBox("Anzahl der Steine", MsgBox_Name)
Input_Anzahl_Steine = Int(Input_Anzahl_Steine)

' --------------------------------------------------
' Berechnung des Maßes

' Außenmaß
If Input_Option = 1 Then
  Result = Aussenmass(Input_Anzahl_Steine)

' Öffnungsmaße
ElseIf Input_Option = 2 Then
  Result = Oeffnungsmass(Input_Anzahl_Steine)

' Vorsprungsmaß
ElseIf Input_Option = 3 Then
  Result = Vorsprungsmass(Input_Anzahl_Steine)
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
ElseIf Result > 100*1000 Then
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
' End der Programschleife

Loop

' --------------------------------------------------
' End of file
' --------------------------------------------------

