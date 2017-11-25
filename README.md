Hier mal ein paar allgemeine Hinweise.

Um Dateien auf den externen Speicher zu schreiben und zu lesen, benötigst du einige Berechtigungen. Guck dazu einfach in die manifest.xml

Berechtigungen sind so eine sache für sich, ab Version 6 reicht es nicht mehr aus, die Berechtigung nur in der manifest.xml zu schreiben, sondern der Benutzer muss noch mal extra danach gefragt werden. Dazu nutze ich die bibliothek EffortlessPermissions https://github.com/DreaminginCodeZH/EffortlessPermissions um das ganze zu vereinfachen.

Hier enigte Informationen zu Berechtigungen: https://developer.android.com/guide/topics/permissions/index.html

Nun aber zum eigentlichen, Excel. Hier gezeigt wird, wie man eine bestehende Excel datei bearbeiten kann und als neue Datei abspeichern. So kannst du eine Vorlage für deine Protokolle machen und dann „nur“ die werte eintragen und neu Abspeichern.

Wenn die App das erste mal gestartet wird, wird ein Ordner angelegt in dem die Protokolle gespeichert werden.

Jetzt schließt du das Tablet per USB, auf dem Tablet solltest du jetzt den erstellen Ordner sehen. Hier legt du eine Vorlage.xls datei rein.

Jetzt kannst du die App benutzen. Trage etwas in die EditTExt ein und drücke auf speichern. Schließe das Tablet wieder per USB an und jetzt sollte neben der Vorlage.cls eine Protokoll.cls datei vorhanden sein.

Der ganze Code ist in der MainActivity Klasse zu finden. Ich habe alles beschrieben was wo gemacht wird. Ich hoffe es hilft dir und du kannst damit Arbeiten.

Ich muss noch sagen, dass der Code natürlich nicht wirklich sauber ist, habe das ja nur zusammengeschustert um ein Beispiel darzustellen.