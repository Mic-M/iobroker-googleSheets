/*******************************************************************************
 * ---------------------------
 * Script: Google Sheets (Google Tabellen) auslesen und in ioBroker in Datenpunkte setzen
 * ---------------------------
 * Autor: Mic
 * Aktuelle Version:    https://github.com/Mic-M/Mic-M-iobroker.googleSheets
 * Support:             https://forum.iobroker.net/topic/28193/
 * 
 * Change log:
 * 0.2 - Added JSON state for JSON Table in VIS (Widget: "basic - Table")
 * 0.1 - initial version
 ******************************************************************************/

/*******************************************************************************
 * WICHTIG
 *******************************************************************************
 Für die Verwendung dieses Scripts musst du für deine Google-Tabelle die Linkfreigabe aktivieren.
 Das benötigen wir,  damit das Script auf deine Google-Tabelle zugreifen kann.
 
 WICHTIG: Damit kann dann außer dir kann dann jeder lesend auf die Datei zugreifen, 
 der den Link (die URL) einsehen kann.  Daher bitte keine sensiblen Daten in der Google-Tabelle vorhalten.
 
Vorgehensweise:
1) In Google Tabellen oben auf den grünen "Freigabe"-Button klicken
2) Die Linkfreigabe entsprechend aktivieren.

Zudem musst du den API-Zugriff aktivieren, Details siehe unten unter "Konfiguration".
 */


/*******************************************************************************
 * Konfiguration: Allgemein
 ******************************************************************************/


// Datenpunkte: Hauptpfad - bitte ohne abschließendem Punkt "."
const STATE_PATH = 'javascript.'+ instance + '.' + 'GoogleSheets.Wertpapiere';

// Wie oft ausführen? 
const GOOGLE_UPDATE_SCHEDULE = '30 2 * * *'; // Um 02:30 jeden Tag

// Sollen Datenpunkte für jede einzelne Zeile angelegt werden?
// Falls nicht, so wird nur ein State ".jsonTable" für die JSON-Tabelle (in VIS: Widget "basic - Table") angelegt (JSON-Datenpunkt wird immer angelegt)
// Beispiel-Datenpunkte: "javascript.0.GoogleSheets.Wertpapiere.Microsoft.AktuellerKurs", "javascript.0.GoogleSheets.Wertpapiere.Microsoft.ISIN", usw.
const CREATE_TABLE_STATES = true;

// Falls CREATE_TABLE_STATES = true:
// In welcher Spalte befindet sich der Name (z.B. Name des Wertpapiers) für die States? 1. Spalte = 1, 2. Spalte 2, usw.
// Der jeweilige Wert dieser Spalte wird verwendet, um darunter Datenpunkte anzulegen.
// d.h. aus "Microsoft" wird dann z.B. "javascript.0.GoogleSheets.Wertpapiere.Microsoft.AktuellerKurs"
const GOOGLE_COL_NO_NAME = 1;



/*******************************************************************************
 * Konfiguration: Google API
 ******************************************************************************/

// Google API-Schlüssel ist erforderlich:
//   1) https://console.developers.google.com/?hl=de aufrufen
//   2.) Im Dashboard "APIS UND DIENSTE AKTIVIEREN" auswählen
//   3.) "Google Sheets API" aktivieren
//   4.) API-Schlüssel erstellen, dabei unter "API-Einschränkungen" auf "Google Sheets API" limitieren (zur Sicherheit)
//   5.) Ggf. ca. 5 Minuten warten, bis Änderungen in Google greifen

const GOOGLE_API_KEY = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
 

// Die ID der Google-Tabelle aus der URL von Google Tabellen.
// Die URL ist beispielsweise:https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx/edit
// Dabei entspricht "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" der ID.
// Diese ID hier eintragen:
const GOOGLE_SPREADSHEET_ID = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';


/*******************************************************************************
 * Konfiguration: Google Tabelle
 ******************************************************************************/

// Name des Tabellenblattes in Google-Tabelle
const GOOGLE_SHEETNAME = 'Test'

// Start-Spalte: erste Spalte, die wir auslesen: 1. Spalte = A, 2. Spalte = B, usw.
const GOOGLE_COLUMN_FIRST = 'A';

// Letzte auszulesende Spalte:  1. Spalte = A, 2. Spalte = B, usw.
const GOOGLE_COLUMN_LAST = 'F';

// Ab welcher Zeile starten wir das auslesen?
const GOOGLE_ROW_FIRST = 1;

// Letzte auszulesende Zeile. 
// Man kann auch z.B. 9999 eingeben, damit werden alle Zeilen berücksichtigt, aber es werden von der Google API nur Zeilen mit Werten übertragen
const GOOGLE_ROW_LAST = 9999;

// JSON-Tabelle: Welche Spalten sollen für die JSON-Tabelle verwendet werden? 1. Spalte = 1, 2. Spalte 2, usw.
// Hiermit kann ebenso die Spalten-Reihenfolge bestimmt werden.
const GOOGLE_JSON_COLUMNS = [1, 6, 4, 5]

/*******************************************************************************
 * Konfiguration: Logging
 ******************************************************************************/
// Debug-Ausgabe im Log?
let LOG_DEBUG = false;


/**********************************************************************************************************
 ++++++++++++++++++++++++++++ Ab hier nichts mehr ändern / Stop editing here! ++++++++++++++++++++++++++++
 *********************************************************************************************************/


/*******************************************************************************
 * Executed on every script start.
 *******************************************************************************/
let googleSchedule;
init();
function init() {

    // Execute initially.
    fetchGoogleSheetsData();

    // Set schedule to update regularly.
    clearSchedule(googleSchedule);
    googleSchedule = schedule(GOOGLE_UPDATE_SCHEDULE, fetchGoogleSheetsData);

}


/*******************************************************************************
 * Main Script
 *******************************************************************************/
function fetchGoogleSheetsData() {

    let cellRange = GOOGLE_COLUMN_FIRST + GOOGLE_ROW_FIRST + ':' + GOOGLE_COLUMN_LAST + GOOGLE_ROW_LAST;
    let googleApiURL = 'https://sheets.googleapis.com/v4/spreadsheets/' + GOOGLE_SPREADSHEET_ID + '/values/' + GOOGLE_SHEETNAME + '!' + cellRange + '?key=' + GOOGLE_API_KEY;

    let thisRequest = require('request');

    let thisOptions = {
      uri: googleApiURL,
      method: 'GET',
      timeout: 5000,
      followRedirect: true,
      maxRedirects: 5
    };

    thisRequest(thisOptions, function(error, response, body) {
        if (!error && response.statusCode == 200) {
            let returnObject = JSON.parse(body)['values'];
            log('Successfuly retrieved data from Google Sheets via Google API.');

            // für JSON
            let jsonResult = '[';

            // Enthält erste Zeile (Überschrift) als Array
            let arrTableHeader = returnObject[0];
            if (LOG_DEBUG) log('DEBUG: Das Array Table-Header hat ' + arrTableHeader.length + ' Elemente.')
            
            // Prüfung der User-Eingaben anhand Anzahl Spalten lt. TableHeader
            if ( (GOOGLE_COL_NO_NAME -1) > arrTableHeader.length ) {
                log('Script-Abbruch: in GOOGLE_COL_NO_NAME wurde Spaltennummer (' + GOOGLE_COL_NO_NAME + ') eingegeben, aber Google Sheets API hat nur (' + arrTableHeader.length + ') Spalten ausgelesen.', 'error');
                return; // raus hier
            }
            

            // Hier holen wir uns die bereits existierenden State-IDs
            let allExistingStates = [];
            $(STATE_PATH + '.*').each(function (id, i) {
                let obj = getObject(id); // Alle Inhalte in Objekt. "obj._id" gibt z.B. State-ID, "obj.common.name" den Namen, etc.
                if (obj){
                    allExistingStates.push(obj._id);
                } 
            });

            // Wir durchlaufen alle weiteren Zeilen (ohne Überschriften, daher i = 1)
            let duplicateChecker = []; // Hier nehmen wir alle Zellwerte von GOOGLE_COL_NO_NAME auf, um zu Prüfen, ob Duplikate vorliegen.
            for (let i = 1; i < returnObject.length; i++) {

                jsonResult += (jsonResult.slice(-1) === '}') ? ', {' : '{'; // Komma dazu, falls es bereits Eintrag gibt

                let lpReturnObj = returnObject[i];
                let name = lpReturnObj[GOOGLE_COL_NO_NAME - 1]; // -1, da wir in der "Konfiguration" mit 1 als Zähler anfangen, nicht 0.
                let targetStatePart1 = STATE_PATH + '.' + (cleanStringForState(name))

                // Wir durchlaufen jede Zelle des aktuellen Zeilen-Arrays, um die States anzulegen.
                if (CREATE_TABLE_STATES) {

                    for (let k = 0; k < lpReturnObj.length; k++) {

                        let lpEntry = lpReturnObj[k]; // Zellinhalt

                        if (k === (GOOGLE_COL_NO_NAME - 1) ) {

                            if (duplicateChecker.indexOf(lpEntry) > -1) {
                                log('Script-Abbruch: Duplikat in Spalte, die für Datenpunkt gedacht ist: Der Wert [' + lpEntry + '] ist mehrfach in Spalte (' + GOOGLE_COL_NO_NAME + ') vorhanden, daher können wir keine Datenpunkte damit erstellen.', 'error');
                                log('Mögliche Lösungen: (A) In Spalte Nr. ' + GOOGLE_COL_NO_NAME + ' der Google-Tabelle alle Duplikate umbenennen / bereinigen, (B) in GOOGLE_COL_NO_NAME andere Spalte wählen., (C) CREATE_TABLE_STATES auf false setzen und damit die Erstellung der Datenpunkte deaktivieren.', 'error');
                                return; // We exit the entire function!
                            } else {
                                // Kein Duplikat, wir nehmen den Wert in den Duplikate-Checker auf
                                if (LOG_DEBUG) log('DEBUG: Kein Duplikat gefunden. Length DuplicateChecker: ' + duplicateChecker.length + ', Element: ' + lpEntry);
                                duplicateChecker.push(lpEntry);
                            }

                        }

                        let lpFinalState = targetStatePart1 + '.' + cleanStringForState(arrTableHeader[k]);
                        if ( lpFinalState.length > (STATE_PATH.length + 2) ) { // Prüfung auf gültigen State.

                            // State aus der Liste der alten States entfernen.
                            arrayRemoveElementsByValue(allExistingStates, lpFinalState);

                            // State erstellen, falls nicht vorhanden, und befüllen
                            createState(lpFinalState, {'name':name + ': ' + cleanString2(arrTableHeader[k]), 'type':'string', 'read':true, 'write':false, 'role':'info', 'def':'' }, function() {
                                if (!isLikeEmpty(lpEntry && (typeof lpEntry === 'string') ) ) {
                                    setState(lpFinalState, lpEntry);
                                }
                            });
                        }
                    }
                }

                // Für JSON
                for (let lpEntry of GOOGLE_JSON_COLUMNS) {
                    // since we start counting with 1 in GOOGLE_JSON_COLUMNS
                    let elem = lpEntry - 1 
                    // Prüfung des Wertes aus GOOGLE_JSON_COLUMNS           
                    if ( (elem >= 0) && (elem <= arrTableHeader.length) ) {
                        jsonResult += (jsonResult.slice(-1) === '"') ? ', ' : ''; // Komma dazu, falls es bereits Eintrag gibt
                        jsonResult += '"' + cleanStringJson(arrTableHeader[elem]) + '": "' + cleanStringJson(lpReturnObj[elem]) + '"';
                    } else {
                        log('Ungültiger Wert in GOOGLE_JSON_COLUMNS: [' + lpEntry + ']', 'warn');
                    }
                }

                jsonResult += '}'

            } // for

            // JSON finalisieren und State setzen
            jsonResult += ']';
            createState(STATE_PATH + '.jsonTable', {'name':'JSON Table', 'type':'string', 'read':true, 'write':true, 'role':'state', 'def':'' }, function() {
                setState(STATE_PATH + '.jsonTable', jsonResult);
            });


            // Jetzt alte States löschen.
            if (CREATE_TABLE_STATES) {
                if (allExistingStates.length > 1)  {
                    // Alle alten, nicht mehr benötigten States löschen, allerdings nicht .jsonTable
                    for (let p = 0; p < allExistingStates.length; p++) {
                        if (allExistingStates[p].indexOf('.jsonTable') === -1) {
                            deleteState(allExistingStates[p]);
                        }
                    }
                    log('We happily deleted ' + allExistingStates.length + ' states no longer needed.')

                } else {
                    log('There were no old states to be deleted.')
                }
            }
        }
        else {
            log('Es ist ein Fehler beim Auslesen der Google-Tabelle aufgetreten; die Daten konnten erst gar nicht geholt werden von Google.', 'error')
            log('Prüfe neben der ausgebenden Meldung deine Script-Konfiguration, etwa auf einen falschen Namen der Google-Tabelle, einzustellen unter GOOGLE_SHEETNAM. Derzeit ist dort gesetzt: ' +  GOOGLE_SHEETNAME, 'error');
            


            //und wende dich an das ioBroker-Forum, wenn du tatsächlich nicht weiter kommst.', 'error')
            log('Antwort von Google: ' + response.body, 'error');
            return; // Raus hier!
        }

    });

 


}



/**
 * Clean a given string for using in ioBroker as part of a atate
 * Will just keep letters, incl. Umlauts, numbers, "-" and "_"
 * @param  {string}  strInput   Input String
 * @return {string}   the processed string 
 */
function cleanStringForState(strInput) {
    let strResult = strInput.replace(/([^a-zA-ZäöüÄÖÜß0-9\-_]+)/gi, '');
    return strResult;
}


/**
 * Clean a given string
 * Will just keep letters, inkl. Umlauts, numbers, "-", "_", "(", ")", "/" and white space " "
 * @param  {string}  strInput   Input String
 * @return {string}   the processed string 
 */
function cleanString2(strInput) {
    let strResult = strInput.replace(/([^a-zA-ZäöüÄÖÜß0-9.\-_\s\/\(\)]+)/gi, '');
    return strResult;
}

/**
 * Clean a given string for JSON
 * @param  {string}  strInput   Input String
 * @return {string}   the processed string 
 */
function cleanStringJson(strInput) {
    let strResult = strInput.replace(/([^a-zA-ZäöüÄÖÜß0-9.,':€@$%&\-_\s\/\(\)]+)/gi, '');
    return strResult;
}




/**
 * Checks if Array or String is not undefined, null or empty.
 * 08-Sep-2019: added check for [ and ] to also catch arrays with empty strings.
 * @param inputVar - Input Array or String, Number, etc.
 * @return true if it is undefined/null/empty, false if it contains value(s)
 * Array or String containing just whitespaces or >'< or >"< or >[< or >]< is considered empty
 */
function isLikeEmpty(inputVar) {
    if (typeof inputVar !== 'undefined' && inputVar !== null) {
        let strTemp = JSON.stringify(inputVar);
        strTemp = strTemp.replace(/\s+/g, ''); // remove all whitespaces
        strTemp = strTemp.replace(/\"+/g, "");  // remove all >"<
        strTemp = strTemp.replace(/\'+/g, "");  // remove all >'<
        strTemp = strTemp.replace(/\[+/g, "");  // remove all >[<
        strTemp = strTemp.replace(/\]+/g, "");  // remove all >]<
        if (strTemp !== '') {
            return false;
        } else {
            return true;
        }
    } else {
        return true;
    }
}



/**
 * Removing Array element(s) by input value. 
 * @param {array}   arr             the input array
 * @param {string}  valRemove       the value to be removed
 * @param {boolean} [exact=true]    OPTIONAL: default is true. if true, it must fully match. if false, it matches also if valRemove is part of element string
 * @return {array}  the array without the element(s)
 */
function arrayRemoveElementsByValue(arr, valRemove, exact) {
 
    if (exact === undefined) exact = true;
 
    for ( let i = 0; i < arr.length; i++){ 
        if (exact) {
            if ( arr[i] === valRemove) {
                arr.splice(i, 1);
                i--; // required, see https://love2dev.com/blog/javascript-remove-from-array/
            }
        } else {
            if (arr[i].indexOf(valRemove) != -1) {
                arr.splice(i, 1);
                i--; // see above
            }
        }
    }
    return arr;
}

