/*******************************************************************************
 * ---------------------------
 * Script: Google Sheets (Google Tabellen) auslesen und in ioBroker in Datenpunkte setzen
 * ---------------------------
 * Autor: Mic
 * Change log:
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
{1}
Zudem musst du den API-Zugriff aktivieren, Details siehe unten unter "Konfiguration".
 */
 
 
/*******************************************************************************
 * Konfiguration
 ******************************************************************************/
 
 
// Datenpunkte: Hauptpfad - bitte ohne abschließendem Punkt "."
const STATE_PATH = 'javascript.'+ instance + '.' + 'GoogleSheets.Wertpapiere';
 
// Wie oft ausführen? 
const GOOGLE_UPDATE_SCHEDULE = '30 2 * * *'; // Um 02:30 jeden Tag
 
 
// Google API-Schlüssel ist erforderlich:
//   1) https://console.developers.google.com/?hl=de aufrufen
//   2.) Im Dashboard "APIS UND DIENSTE AKTIVIEREN" auswählen
//   3.) "Google Sheets API" aktivieren
//   4.) API-Schlüssel erstellen, dabei unter "API-Einschränkungen" auf "Google Sheets API" limitieren (zur Sicherheit)
//   5.) Ggf. ca. 5 Minuten warten, bis Änderungen in Google greifen
 
const GOOGLE_API_KEY = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
 
 
// Die ID der Google-Tabelle aus der URL von Google Tabellen.
// Die URL ist beispielsweise:https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx/edit
// Dabei entspricht "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" der ID.
const GOOGLE_SPREADSHEET_ID = 'yyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy';
 
// Name des Tabellenblattes in Google-Tabelle
const GOOGLE_SHEETNAME = 'Test'
 
// Start-Spalte: erste Spalte, die wir auslesen: 1. Spalte = A, 2. Spalte = B, usw.
const GOOGLE_COLUMN_FIRST = 'A';
 
// Letzte auszulesende Spalte
const GOOGLE_COLUMN_LAST = 'F';
 
// Ab welcher Zeile starten wir das auslesen?
const GOOGLE_ROW_FIRST = 1;
 
// Letzte auszulesende Zeile. Man kann auch z.B. 9999 eingeben, damit werden alle Zeilen berücksichtigt, aber es werden von der Google API nur Zeilen mit Werten übertragen
const GOOGLE_ROW_LAST = 9999;
 
// In welcher Spalte befindet sich der Name der Aktie? 1. Spalte = 1, 2. Spalte 2, usw.
const GOOGLE_COL_NO_NAME = 1;
 
// States immer bereinigen? Sobald sich Spaltentitel ändern, Zeilen gelöscht werden, usw., werden hiermit automatisch alte States gelöscht
// Wird absolut empfohlen, auf true zu lassen.
const GOOGLE_REMOVE_STATES = true;
 
 
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
 
            // Enthält erste Zeile (Überschrift) als Array
            let arrTableHeader = returnObject[0];
 
            // Hier holen wir uns die bereits existierenden State-IDs
            let allExistingStates = [];
            $(STATE_PATH + '.*').each(function (id, i) {
                let obj = getObject(id); // Alle Inhalte in Objekt. "obj._id" gibt z.B. State-ID, "obj.common.name" den Namen, etc.
                if (obj){
                    allExistingStates.push(obj._id);
                } 
            });
 
            // Wir durchlaufen alle weiteren Zeilen (ohne Überschriften, daher i = 1)
            for (let i = 1; i < returnObject.length; i++) {
                let lpReturnObj = returnObject[i];
                let name = lpReturnObj[GOOGLE_COL_NO_NAME - 1]; // -1, da wir in der "Konfiguration" mit 1 als Zähler anfangen, nicht 0.
                let targetStatePart1 = STATE_PATH + '.' + (cleanStringForState(name))
 
                // Wir durchlaufen jede Zelle des aktuellen Zeilen-Arrays
                for (let k = 0; k < lpReturnObj.length; k++) {
                    let lpEntry = returnObject[i][k];
 
                    let lpFinalState = targetStatePart1 + '.' + cleanStringForState(arrTableHeader[k]);
                    if ( lpFinalState.length > (STATE_PATH.length + 2) ) { // Prüfung auf gültigen State.
 
                        // State aus der Liste der alten States entfernen.
                        arrayRemoveElementsByValue(allExistingStates, lpFinalState);
 
                        // State erstellen, falls nicht vorhanden, und befüllen
                        createState(lpFinalState, {'name':name + ': ' + cleanString2(arrTableHeader[k]), 'type':'string', 'read':true, 'write':true, 'role':'info', 'def':'' }, function() {
                            if (!isLikeEmpty(lpEntry && (typeof lpEntry === 'string') ) ) {
                                setState(lpFinalState, lpEntry);
                            }
                        });
                    }
                }                
            } // for
 
            // Jetzt alte States löschen.
            if (allExistingStates.length > 0)  {
                // Alle alten, nicht mehr benötigten States löschen
                for (let p = 0; p < allExistingStates.length; p++) {
                    deleteState(allExistingStates[p]);
                }
                log('We happily deleted ' + allExistingStates.length + ' states no longer needed.')
 
            } else {
                log('There were no old states to be deleted.')
            }
 
        }
        else {
            log('Response: ' + response.body, 'warn');
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
        strTemp = strTemp.replace(/[+/g, "");  // remove all >[<
        strTemp = strTemp.replace(/]+/g, "");  // remove all >]<
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
 
 
