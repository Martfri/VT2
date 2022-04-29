# VT2: Automatische Generierung von Web-Applikationen aus Excel-Dokumenten

Das Ziel dieser VT ist es, eine Web-Applikation zu erstellen, die mit einem Excel-Dokument interagieren kann, welches als Micro-Service zur Verfügung steht. Dem Nutzer der Web-Applikation soll es dabei möglich sein, durch seine Eingabe der Inputs in eine Eingabemaske den Micro-Service abzurufen. Beim Aufruf des Services soll dieser die Parameter in vordefinierte Felder der Excel-Datei schreiben sowie die Formeln in der Datei berechnen und vordefinierte Zellwerte als Ergebnis zurückgeben und im UI der Web-Applikation abbilden. 
Die Interaktion mit dem Excel-Dokument soll auf dem «model at runtime» Prinzip beruhen. Für die Interaktion mit dem Modell (also dem Excel-Dokument) soll Microsoft Graph [1] genutzt werden. Microsoft Graph bietet eine REST-API für Excel Dokumente, die in der Cloud gespeichert sind, um deren Integration in Web-Applikationen zu ermöglichen. Die Performance der Technologie soll gemäss Best-Practice-Methoden evaluiert werden.
Ausserdem soll es ermöglicht werden, das UI der Web-Applikation automatisch aus dem Excel-Dokument zu generieren (z. B. aus vordefinierten Input und Output Feldern des Excel-Dokuments). Alternativ dazu soll ein Ansatz erarbeitet werden, der es erlaubt, das Excel-Dokument in eine bestehend Web-Applikation zu integrieren. Dabei sollen Inputs für Berechnungen nicht aus einer Eingabemaske stammen, sondern aus einer integrierten Datenbank bezogen werden, welche über die benötigten Daten verfügt.
Die Validierung der verschiedenen Ansätze erfolgt anhand eines Beispiels aus der Industrie, das ein Excel-Dokument zur Berechnung der benötigten Ersatzteile für elektrische Geräte verwendet. Die im Excel-Dokument definierten Berechnungen sollen als Micro-Service verfügbar gemacht werden. 
Die in der Excel-Datei gelisteten Ersatzteile generieren Kosten für ihre Lagerung und Konservierung. Es ist daher ein natürliches Ziel, so wenig wie möglich davon zu besitzen. Andererseits stellt ein möglichst grosser Bestand sichergestellt werden, sodass im Falle von Ausfällen diese umgehend repariert werden können, wodurch wirtschaftliche Auswirkungen verringert werden. Gemäss verschiedener Input-Daten berechnet das Excel-Dokument den optimalen Bestand an Ersatzteilen, sodass die Gesamtkosten minimiert werden.
Die VT lässt sich in zwei Phasen unterteilen. Zu den verpflichtenden Zielen sind einige zusätzliche, optionale Ziele definiert (Zeitangaben sind ungefähr):
 

# 1.	Ende März - bis Ende April: 
-    Untersuchung von Microsoft Graph: Wie werden Excel-Dokumente als Micro-Service zur Verfügung gestellt? Wie kann dieser Service genutzt werden (z.B. Authentifizierung, Authorisierung).
-    Erstellung eines Proof-of-Concepts, zu einem auf Microsoft Graph basierenden Micro-Service der auf ein Excel-Dokument zugreift, Eingabedaten verändert, Berechnungen von Formeln triggert, und die neuen Werte der Formeln ausliest
-    Messung der Performance des Micro-Services gemäss Best-Practice-Methoden
-    Verfassen eines Berichts über bisherige Ergebnisse
-   (optional) Verfassen eines technischen Papers für die Konferenz MODELS 2022 (Einreichungstermin 18. Mai). Im Paper sollen die beiden Tools Microsoft Graph und formulas (Python) miteinander verglichen werden. Hauptaugenmerk gilt dabei auf Performance. 

# 2.	Ab Mai: 
-	Integration des Excel-Dokuments in eine bestehende Web-Applikation. Beispiel: aktuelle Daten aus einer Datenbank lesen und diese Daten als Input für den Excel-Service verwenden
-	Generieren eines UI aus dem Excel Dokument. Insbesondere sollen Input Felder generiert sowie entsprechende Outputs ausgegeben werden
-	Untersuchen von möglichen IT-Security-Problemen und Erarbeitung verschiedener Massnahmen (z.B. Inputvalidierung, Authentifizierung, Authorisierung)
-	Fertigstellung des Berichts
-	Erreichen von mindestens einem der optionalen Ziele:
o	(optional) Bestehende Visualisierungen, die im Excel-Dokument vorhanden sind (z. B. ein Diagramm), sollen automatisch extrahiert und im generierten UI verwendet werden 
o	(optional) Excel als Test Oracle: Testen des Proof-of-Concepts gegen Excel. Vergleich der Outputs und Prüfung auf Korrektheit
o	(optional) Rollenbasierte Anpassung des UI (z.B. Student sieht andere Felder als Dozent und kann weniger verändern)
