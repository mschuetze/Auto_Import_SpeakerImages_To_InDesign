// v0.5.1

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
    if ((doc.paragraphStyles.itemByName("title").isValid) && (doc.paragraphStyles.itemByName("speakers").isValid)) {
        // Checken, ob Objektformat überhaupt existiert
        if ((doc.objectStyles.itemByName("speakerBild_frei").isValid) && (doc.objectStyles.itemByName("speakerBild_wp").isValid)) {
            // Starte Loop durch alle Seiten des Doks
            // alert("Seitenanzahl: " + doc.pages.length);
            for (var i = 0; i < doc.pages.length; i++) {
                var page = doc.pages[i];  // Aktuelle Seite
                // alert("Seite: " + i);
                // Durch alle Textrahmen der Seite iterieren
                for (var j = 0; j < page.textFrames.length; j++) {
                    var textFrame = page.textFrames[j];  // Aktueller Textrahmen
                    // Durch alle Absätze im Textrahmen iterieren
                    for (var k = 0; k < textFrame.paragraphs.length; k++) {
                        var paragraph = textFrame.paragraphs[k];  // Aktueller Absatz
                        // Prüfen, ob der Absatz das Absatzformat "TITLE" hat
                        if (paragraph.appliedParagraphStyle.name == "title") {
                            // var titleContent = paragraph.contents
                            // alert(titleContent);

                            // // Kürze den Titel (alles ab der öffnenden eckigen Klammer fliegt weg)
                            // var titleContent = titleContent.replace(/\[.*?\]/g, '');
                            // // alert(titleContent);

                            // // angepassten Titel im Absatz einfügen
                            // paragraph.contents = titleContent
                        }
                        // Prüfen, ob der Absatz das Absatzformat "SPEAKERS" hat
                        else if (paragraph.appliedParagraphStyle.name == "speakers") {
                            var speakerName = paragraph.contents
                            // alert(speakerName);
                            
                            // ### SPEAKERNAMEN IM DOK ANPASSEN: START ###

                            // Ersetze die öffnende Klammer durch eine Pipe
                            var speakerName_doc = speakerName.replace(/\(/g, '| ');
                            // alert(speakerName_doc);

                            // Entferne die schließende Klammer
                            var speakerName_doc = speakerName_doc.replace(/\)/g, '');
                            // alert(speakerName_doc);
                            
                            // angepassten Namen im Absatz einfügen
                            paragraph.contents = speakerName_doc

                            // ### SPEAKERNAMEN IM DOK ANPASSEN: ENDE ###



                            // ### SPEAKERNAMEN FÜR ABBILDUNG: START ###

                            // Kürze den Text auf den Namen (alles ab der öffnenden Klammer fliegt weg)
                            var speakerName = speakerName.split("(")[0];
                            // alert(speakerName);

                            // Alles in Kleinbuchstaben umwandeln
                            var speakerName = speakerName.toLowerCase();
                            // alert(speakerName);
                            
                            // Umlaute ersetzen
                            // Funktion, um die deutschen Umlaute zu ersetzen
                            var textContent = speakerName;
                            var replacedText = replaceGermanUmlauts(textContent);
                            function replaceGermanUmlauts(text) {
                                // Ersetzen der deutschen Umlaute durch ihre entsprechenden Zeichen
                                var umlautReplacedText = text
                                    .replace(/ä/g, 'ae')    // 'ä' wird zu 'ae'
                                    .replace(/ö/g, 'oe')    // 'ö' wird zu 'oe'
                                    .replace(/ü/g, 'ue')    // 'ü' wird zu 'ue'
                                    .replace(/ß/g, 'ss')    // 'ß' wird zu 'ss'
                                    .replace(/Ä/g, 'Ae')    // 'Ä' wird zu 'Ae'
                                    .replace(/Ö/g, 'Oe')    // 'Ö' wird zu 'Oe'
                                    .replace(/Ü/g, 'Ue');   // 'Ü' wird zu 'Ue'
                                
                                return umlautReplacedText;
                            }
                            speakerName = replacedText;
                            // Vor- und Nachname vertauschen
                            // Vorname
                            var FirstName = speakerName.split(/\s+/)[0];
                            // alert(FirstName);
                            // Nachname
                            var LastName = speakerName.split(/\s+/)[1];
                            // alert(LastName);
                            // Dateiname für Freisteller erzeugen
                            var fileName_frei = LastName + "_" + FirstName + "_frei.png";
                            // alert(fileName_frei);
                            // Dateiname für Wordpress erzeugen
                            var fileName_wp = LastName + "_" + FirstName + "_wp_1024x1024.jpg";
                            // alert(fileName_wp);
                            // Pfad zum Bild angeben – dieser Teil ist immer gleich
                            var folderPath_static = "/Volumes/GRAFIK/Grafik1/Speaker- und Autorenbilder/";
                            // Pfad zum Bild angeben – dieser Teil variiert je nach Nachname
                            // Ersten Buchstaben erfassen – dieser bestimmt den Pfad zum Speakerbild
                            var firstLetter = fileName_frei.charAt(0);
                            // alert(firstLetter);
                            if (firstLetter === "a" || firstLetter === "b" || firstLetter === "c") {
                                var folderPath_variable = "abc/";
                            }
                            else if (firstLetter === "d" || firstLetter === "e" || firstLetter === "f") {
                                var folderPath_variable = "def/";
                            }
                            else if (firstLetter === "g" || firstLetter === "h" || firstLetter === "i") {
                                var folderPath_variable = "ghi/";
                            }
                            else if (firstLetter === "j" || firstLetter === "k" || firstLetter === "l") {
                                var folderPath_variable = "jkl/";
                            }
                            else if (firstLetter === "m" || firstLetter === "n" || firstLetter === "o") {
                                var folderPath_variable = "mno/";
                            }
                            else if (firstLetter === "p" || firstLetter === "q" || firstLetter === "r") {
                                var folderPath_variable = "pqr/";
                            }
                            else if (firstLetter === "s" || firstLetter === "t") {
                                var folderPath_variable = "st/";
                            }
                            else if (firstLetter === "u" || firstLetter === "v" || firstLetter === "w") {
                                var folderPath_variable = "uvw/";
                            }
                            else if (firstLetter === "x" || firstLetter === "y" || firstLetter === "z") {
                                var folderPath_variable = "xyz/";
                            }
                            // Statischen + variablen Teil des Pfades zusammenfügen
                            var folderPath = folderPath_static + folderPath_variable
                            // alert(folderPath);
                            // Dateipfad für Freisteller erzeugen
                            var filePath_frei = new File(folderPath + fileName_frei);
                            // Dateipfad für Wordpress erzeugen
                            var filePath_wp = new File(folderPath + fileName_wp);
                        }
                    }
                }
                // ##### Speakerbild platzieren #####
                // Prüfen, ob Speakerbild existiert
                if (filePath_frei.exists) {
                    // Rahmen für das Bild erstellen (Rechteck)
                    var frame = page.rectangles.add({
                        geometricBounds: ["24px", "24px", "124px", "124px"] // Position und Größe des Rahmens (oben, links, unten, rechts)
                    });
                    // Bild in den Textrahmen platzieren
                    var placedImage = frame.place(filePath_frei)[0];
                    var objectStyle = doc.objectStyles.itemByName("speakerBild_frei");
                    frame.appliedObjectStyle = objectStyle;
                    // Bild auf die Größe des Rahmens skalieren (optional)
                    placedImage.fit(FitOptions.FILL_PROPORTIONALLY);
                }
                else {
                    if (filePath_wp.exists) {
                        // Erstellen eines kreisrunden Rahmens (Ellipse)
                        var radius = 100; // Radius des Kreises (kann angepasst werden)
                        var xPosition = 150; // X-Position (kann angepasst werden)
                        var yPosition = 150; // Y-Position (kann angepasst werden)
                        // Ellipse (kreisrunder Objektrahmen) erstellen
                        var circleFrame = page.ovals.add({
                            geometricBounds: [yPosition - radius, xPosition - radius, yPosition + radius, xPosition + radius]
                        });
                        // Bild in den Textrahmen platzieren
                        var placedImage = circleFrame.place(filePath_wp)[0];
                        var objectStyle = doc.objectStyles.itemByName("speakerBild_wp");
                        circleFrame.appliedObjectStyle = objectStyle;
                        // Bild auf die Größe des Rahmens skalieren (optional)
                        placedImage.fit(FitOptions.FILL_PROPORTIONALLY);
                    }
                    else {
                        // alert("Die Bilddatei existiert nicht unter dem angegebenen Pfad: " + filePath_frei + filePath_wp);
                    }
                }
            }
        }
        else {
            alert("Die benötigten Objektformate 'speakerBild_frei' und 'speakerBild_wp' existieren nicht. Bitte erstellen, zuweisen und Skript neu starten.");
        }
    }
    else {
        alert("Die benötigten Absatzformate 'title' und 'speakers' existiert nicht. Bitte erstellen, zuweisen und Skript neu starten.");
    }

}
