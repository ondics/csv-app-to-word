Messwerte nach Microsoft Word mit VBA übertragen
================================================

In Microsoft Word ist VBA (Visual Basic for Applications) eingebaut. Damit können Werte aus der Lufft I-BOX mit folgenden Schritten übertragen werden.

Voraussetzung 
-------------

Die OUT-App CSV (http://www.lufft-i-box.de/app/view/?appname=opus20) muss auf der Lufft I-BOX installiert sein.

Schritt 1: Start von VBA in Microsoft Word

Schritt 2: Schreiben Sie das VBA-Script, mit dem die Daten von der lufft I-BOX geholt und verarbeitet werden.

Als erklärendes Beispiel hierzu soll die Prozedur "GetChannels" dienen. Sie lädt die Metadaten aller in der lufft I-BOX gemessenen Kanäle und fügt sie an der Cursorposition im Dokument ein. Das Snippet kann per Copy&Paste direkt in den VBA Editor übertragen werden.

	Sub GetChannels()
	  Const URL$ = "http://<hostname>/ab/index.php/csv/1/api/getchannels"
	  Dim txt As String, i As Long, ret As String
	  With CreateObject("MSXML2.XMLHTTP")
	    .Open "GET", URL, False
	    .send
	    txt = .responseText
	  End With
	  Selection.TypeText (txt)
	End Sub

Ersetzen Sie <hostname> mit der IP-Adresse Ihrer Lufft I-BOX.

Schritt 3: Sie können diese Prozedur nun mit als Makro ausführen.

Um Messwerte zu erhalten, verwenden Sie folgende URL:

	http://<hostname>/ab/index.php/csv/1/api/getvalues?valueids=73-100

Die <valueids> entnehmen sie dem Befehl GetChannels. Mehrere <valueids> können kommagetrennt hintereinander gehängt werden.
