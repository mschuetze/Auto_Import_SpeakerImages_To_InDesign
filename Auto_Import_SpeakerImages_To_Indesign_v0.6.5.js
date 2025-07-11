// v0.6.5

// vorsichtshalber die Dialoge einschalten
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALERTS;

// Utility-Funktion zur Bereinigung des Strings
function cleanString(text) {
    text = text.replace(/[\u200B-\u200D\uFEFF]/g, '');
    text = text.toLowerCase();
    text = text.replace(/^\s+|\s+$/g, ''); // statt .trim() – kompatibel mit ExtendScript
    return text;
}

// Funktion, um die deutschen Umlaute zu ersetzen
function replaceGermanUmlauts(text) {
    return text
        .replace(/Ä/g, 'Ae')
        .replace(/ä/g, 'ae')
        .replace(/Á/g, 'A')
        .replace(/á/g, 'a')
        .replace(/À/g, 'A')
        .replace(/à/g, 'a')
        .replace(/Â/g, 'A')
        .replace(/â/g, 'a')

        .replace(/Č/g, 'C')
        .replace(/č/g, 'c')
        .replace(/Ç/g, 'C')
        .replace(/ç/g, 'c')

        .replace(/É/g, 'E')
        .replace(/é/g, 'e')
        .replace(/È/g, 'E')
        .replace(/è/g, 'e')
        .replace(/ë/g, 'e')

        .replace(/Ğ/g, 'G')
        .replace(/ğ/g, 'g')

        .replace(/í/g, 'i')
        .replace(/î/g, 'i')
        .replace(/ï/g, 'i')

        .replace(/Ñ/g, 'N')
        .replace(/ñ/g, 'n')

        .replace(/Ö/g, 'Oe')
        .replace(/ö/g, 'oe')
        .replace(/ó/g, 'o')
        .replace(/ô/g, 'o')

        .replace(/Ř/g, 'R')
        .replace(/ř/g, 'r')

        .replace(/Š/g, 'S')
        .replace(/š/g, 's')
        .replace(/Ş/g, 'S')
        .replace(/ş/g, 's')
        .replace(/ß/g, 'ss')

        .replace(/Ü/g, 'Ue')
        .replace(/ü/g, 'ue')
        .replace(/ú/g, 'u')
        .replace(/û/g, 'u')

        .replace(/Ž/g, 'Z')
        .replace(/ž/g, 'z');
}

// Prüfe, ob das Dok bereits geöffnet ist
if (app.documents.length == 0) {
    alert("Es ist kein Dokument geöffnet.");
} else {
    var doc = app.activeDocument;

    if ((doc.paragraphStyles.itemByName("title").isValid) && (doc.paragraphStyles.itemByName("speakers").isValid)) {
        if ((doc.objectStyles.itemByName("speakerBild_frei").isValid) && (doc.objectStyles.itemByName("speakerBild_wp").isValid)) {
            for (var i = 0; i < doc.pages.length; i++) {
                var page = doc.pages[i];
                for (var j = 0; j < page.textFrames.length; j++) {
                    var textFrame = page.textFrames[j];
                    for (var k = 0; k < textFrame.paragraphs.length; k++) {
                        var paragraph = textFrame.paragraphs[k];
                        
                        if (paragraph.appliedParagraphStyle.name == "title") {
                            // optionaler Code für "title"
                        } else if (paragraph.appliedParagraphStyle.name == "speakers") {
                            var speakerName = cleanString(paragraph.contents);

                            // Kürze den Text auf den Namen (alles ab der Pipe fliegt weg)
                            speakerName = speakerName.split("|")[0];

                            // Umlaute ersetzen
                            speakerName = replaceGermanUmlauts(speakerName);

                            // Vor- und Nachname extrahieren
                            var FirstName = speakerName.split(/\s+/)[0];
                            var LastName = speakerName.split(/\s+/)[1];

                            var fileName_frei = LastName + "_" + FirstName + "_frei.png";
                            // alert(fileName_frei);
                            var fileName_frei_dr = LastName + "_" + FirstName + "_dr_frei.png";
                            var fileName_wp = LastName + "_" + FirstName + "_wp_1024x1024.jpg";
                            // alert(fileName_wp);
                            var fileName_wp_dr = LastName + "_" + FirstName + "_dr_wp_1024x1024.jpg";


                            var folderPath_static = "/Volumes/GRAFIK/Grafik1/Speaker- und Autorenbilder/";
                            var firstLetter = fileName_frei.charAt(0);

                            var folderPath_variable = "";
                            if ("abc".indexOf(firstLetter) !== -1) folderPath_variable = "abc/";
                            else if ("def".indexOf(firstLetter) !== -1) folderPath_variable = "def/";
                            else if ("ghi".indexOf(firstLetter) !== -1) folderPath_variable = "ghi/";
                            else if ("jkl".indexOf(firstLetter) !== -1) folderPath_variable = "jkl/";
                            else if ("mno".indexOf(firstLetter) !== -1) folderPath_variable = "mno/";
                            else if ("pqr".indexOf(firstLetter) !== -1) folderPath_variable = "pqr/";
                            else if ("st".indexOf(firstLetter) !== -1) folderPath_variable = "st/";
                            else if ("uvw".indexOf(firstLetter) !== -1) folderPath_variable = "uvw/";
                            else if ("xyz".indexOf(firstLetter) !== -1) folderPath_variable = "xyz/";

                            var folderPath = folderPath_static + folderPath_variable;

                            var filePath_frei = new File(folderPath + fileName_frei);
                            // alert(filePath_frei);
                            var filePath_frei_dr = new File(folderPath + fileName_frei_dr);
                            var filePath_wp = new File(folderPath + fileName_wp);
                            // alert(filePath_wp);
                            var filePath_wp_dr = new File(folderPath + fileName_wp_dr);

                        }
                    }
                }

                // Bildplatzierung
                if ((filePath_frei.exists) || (filePath_frei_dr.exists)) {
                    var frame = page.rectangles.add({
                        geometricBounds: ["24px", "24px", "124px", "124px"]
                    });
                    var placedImage = (filePath_frei.exists) ?
                        frame.place(filePath_frei)[0] :
                        frame.place(filePath_frei_dr)[0];
                    
                    frame.appliedObjectStyle = doc.objectStyles.itemByName("speakerBild_frei");
                    placedImage.fit(FitOptions.PROPORTIONALLY);
                } else if ((filePath_wp.exists) || (filePath_wp_dr.exists)) {
                    var radius = 100;
                    var xPosition = 150;
                    var yPosition = 150;

                    var circleFrame = page.ovals.add({
                        geometricBounds: [yPosition - radius, xPosition - radius, yPosition + radius, xPosition + radius]
                    });
                    var placedImage = (filePath_wp.exists) ?
                        circleFrame.place(filePath_wp)[0] :
                        circleFrame.place(filePath_wp_dr)[0];

                    circleFrame.appliedObjectStyle = doc.objectStyles.itemByName("speakerBild_wp");
                    placedImage.fit(FitOptions.FILL_PROPORTIONALLY);
                } else {
                    alert("Die Bilddatei existiert nicht unter dem angegebenen Pfad:\n" + filePath_frei + "\n" + filePath_wp);
                }
            }
        } else {
            alert("Die Objektformate 'speakerBild_frei' und/oder 'speakerBild_wp' fehlen.");
        }
    } else {
        alert("Absatzformate 'title' und/oder 'speakers' fehlen.");
    }
}
