// v0.1.0

// vorsichtshalber die Dialoge einschalten
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALERTS;

// Prüfe, ob das Dok bereits geöffnet ist
if (app.documents.length == 0) {
    alert("Es ist kein Dokument geöffnet.");
}

else {
    // Variablen deklarieren
    var doc = app.activeDocument;

    // Checken, ob Absatzformat überhaupt existiert
    if (app.activeDocument.paragraphStyles.itemByName("speakers").isValid) {
        // Starte Loop durch alle Seiten des Doks
        for (var i = 0; i < doc.pages.length; i++) {
            var page = doc.pages.item(i);  // Aktuelle Seite
            // Durch alle Textrahmen der Seite iterieren
            for (var j = 0; j < page.textFrames.length; j++) {
                var textFrame = page.textFrames[j];  // Aktueller Textrahmen
                // Durch alle Absätze im Textrahmen iterieren
                for (var k = 0; k < textFrame.paragraphs.length; k++) {
                    var paragraph = textFrame.paragraphs[k];  // Aktueller Absatz
                    // Prüfen, ob der Absatz das Absatzformat "SPEAKERS" hat
                    if (paragraph.appliedParagraphStyle.name == "speakers") {
                        var speakerName = paragraph.contents
                        alert(speakerName);
                        // Kürze den Text auf den Namen (alles nach der öffnenden Klammer fliegt weg)
                        var speakerName_trimmed = speakerName.split("(")[0];
                        alert(speakerName_trimmed);
                    }
                }
            }
        }

    }
    else {
        alert("Das benötigte Absatzformat 'speakers' existiert nicht. Bitte erstellen und Skript neu starten.");
    }

}
