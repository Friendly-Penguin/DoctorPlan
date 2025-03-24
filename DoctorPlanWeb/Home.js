
    let cellToHighlight;
    let messageBanner;

    // Initialization when Office JS and JQuery are ready.
    Office.onReady(() => {
        $(() => {
            // Initialize he Office Fabric UI notification mechanism and hide it.
            const element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            //// If not using Excel 2016 or later, use fallback logic.
            //if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
            //    $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
            //    $('#button-text').text("Display!");
            //    $('#button-desc').text("Display the selection");

            //    $('#highlight-button').on('click',displaySelectedCells);
            //    return;
            //}


            // Imposta il testo per il pulsante Hello World
            //$('#button-text1').text("Hello World!");
            //$('#button-desc1').text("Writes Hello World in cell A1");


            initClingo();

            $('#bottone-func-1').on('click', helloWorld);

            // Gestore per il pulsante di conferma nel form
            $('#confirm-table-button').on('click', function () {
                // Ottieni il numero di righe dal campo di input
                const rows = parseInt($('#table-rows').val()) || 3; // Default a 3 se non valido

                // Nascondi il form
                $('#table-config').hide();

                // Crea la tabella con il numero di righe specificato
                createSurgeonShiftTable(rows);
            });

            //// Gestore per il pulsante di annullamento
            //$('#cancel-table-button').on('click', function () {
            //    // Nascondi il form senza fare nulla
            //    $('#table-config').hide();
            //});

            $('#btnRisolvi').on('click', async function () {

                const programmaAsp = $('#aspProgram').val();
                const risultatoDiv = $('#risultato');

                risultatoDiv.text("Risolvendo il programma ASP...");

                try {
                    const risultato = await risolviASP(programmaAsp);
                    risultatoDiv.text(`<pre>${JSON.stringify(risultato, null, 2)}</pre>`);
                } catch (error) {
                    risultatoDiv.text(`Errore: ${error}`);
                }


            });

        });
    });

    

    async function highlightHighestValue() {
        try {
            await Excel.run(async (context) => {
                const sourceRange = context.workbook.getSelectedRange().load("values, rowCount, columnCount");

                await context.sync();
                let highestRow = 0;
                let highestCol = 0;
                let highestValue = sourceRange.values[0][0];

                // Find the cell to highlight
                for (let i = 0; i < sourceRange.rowCount; i++) {
                    for (let j = 0; j < sourceRange.columnCount; j++) {
                        if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                            highestRow = i;
                            highestCol = j;
                            highestValue = sourceRange.values[i][j];
                        }
                    }
                }

                cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                sourceRange.worksheet.getUsedRange().format.fill.clear();
                sourceRange.worksheet.getUsedRange().format.font.bold = false;

                // Highlight the cell
                cellToHighlight.format.fill.color = "orange";
                cellToHighlight.format.font.bold = true;
                await context.sync;
            });
        } catch (error) {
            errorHandler(error);
        }
    }

    

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

// Definisci la funzione writeHelloWorld fuori dal blocco onReady
function helloWorld() {
    Excel.run(function (context) {
        // Ottieni la cella A1
        var range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");

        // Imposta il valore "Hello World"
        range.values = [["Hello World"]];

        // Applica le modifiche
        return context.sync();
    }).catch(function (error) {
        console.log("Error: " + error);
        showNotification("Error", error);
    });
}

// Modifichiamo la funzione createSurgeonShiftTable per accettare il numero di righe
function createSurgeonShiftTable(numRows) {
    Excel.run(function (context) {
        // Ottieni il foglio di lavoro attivo
        var sheet = context.workbook.worksheets.getActiveWorksheet();

        // Definiamo il range per la tabella (iniziamo dalla cella A1)
        var headerRange = sheet.getRange("A1:B1");

        // Impostiamo le intestazioni della tabella
        headerRange.values = [["Chirurghi", "Turni"]];

        // Formattazione delle intestazioni
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = "#4472C4";  // Colore blu
        headerRange.format.font.color = "white";

        // Calcoliamo il range della tabella completa
        var fullRangeAddress = "A1:B" + (numRows + 1); // +1 perché la prima riga è l'intestazione
        var fullRange = sheet.getRange(fullRangeAddress);

        // Crea una tabella con le intestazioni
        var table = sheet.tables.add(fullRange, true);
        table.name = "TabellaChirurghiTurni";

        // Aggiustiamo la larghezza delle colonne
        sheet.getRange("A:A").format.columnWidth = 150;
        sheet.getRange("B:B").format.columnWidth = 150;

        // Creiamo un array di righe vuote per la tabella
        var data = [];
        for (var i = 0; i < numRows; i++) {
            data.push(["", ""]);
        }

        // Se vogliamo pre-compilare con alcuni esempi (opzionale)
        if (data.length >= 1) data[0] = ["Dr. Rossi", "Lunedì 8:00-16:00"];
        if (data.length >= 2) data[1] = ["Dr. Bianchi", "Martedì 8:00-16:00"];
        if (data.length >= 3) data[2] = ["Dr. Verdi", "Mercoledì 8:00-16:00"];

        // Aggiungiamo le righe alla tabella
        if (numRows > 0) {
            var dataRange = sheet.getRange("A2:B" + (numRows + 1));
            dataRange.values = data;
        }

        // Mostriamo un messaggio all'utente
        sheet.getRange("D1").values = [["Tabella creata con successo!"]];
        sheet.getRange("D2").values = [["Hai creato una tabella con " + numRows + " righe."]];

        return context.sync();
    }).catch(function (error) {
        console.log("Error: " + error);
        // Se hai una funzione showNotification definita nel tuo codice
        if (typeof showNotification === "function") {
            showNotification("Error", error);
        }
    });
}

async function initClingo() {
    try {
        // Carica il modulo WASM di Clingo
        const clingoModule = await import('./wasm/clingo.web.js');

        // Esplora in dettaglio l'oggetto clingoModule
        console.log("Tipo di clingoModule:", typeof clingoModule);
        console.log("Contenuto di clingoModule:");
        console.dir(clingoModule); // Mostra tutte le proprietà dell'oggetto

        // Verifica se Clingo è disponibile globalmente
        if (window.clingo) {
            console.log("Clingo trovato globalmente:", window.clingo);

            // Esplora la struttura di window.clingo per essere sicuri
            console.dir(window.clingo);  // Esplora la struttura di window.clingo

            // Inizializza Clingo con il metodo 'init', ma senza la funzione locateFile
            await window.clingo.init("./wasm/");

            console.log("Clingo inizializzato con successo.");
        } else {
            throw new Error("Clingo non è disponibile globalmente.");
        }

    } catch (error) {
        console.error("Errore nell'inizializzazione di Clingo:", error);
    }
}









// Funzione di esempio per risolvere un programma ASP
async function risolviASP(programma) {
    if (!window.clingo) {
        await initClingo();
    }

    try {
        // Esegui il programma ASP
        const risultato = await window.clingo.run(programma);
        return risultato;
    } catch (error) {
        console.error("Errore nell'esecuzione del programma ASP:", error);
        return null;
    }
}