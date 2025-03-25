
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


            

            $('#bottone-func-1').on('click', helloWorld);
            $('#cancel-table-button').on('click', deleteTable);

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

            $('#btnRisolvi').on('click', risolviClingo);

        });
    });


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
        if (data.length >= 1) data[0] = ["luigi", "mattina"];
        if (data.length >= 2) data[1] = ["antonio", "pomeriggio"];
        

        // Aggiungiamo le righe alla tabella
        if (numRows > 0) {
            var dataRange = sheet.getRange("A2:B" + (numRows + 1));
            dataRange.values = data;
        }



        return context.sync();
    }).catch(function (error) {
        console.log("Error: " + error);
    });
}

function deleteTable() {
    Excel.run(async function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var tables = sheet.tables; // Ottiene tutte le tabelle nel foglio
        tables.load("items/name"); // Carica i nomi delle tabelle

        await context.sync(); // Sincronizza per ottenere i dati

        for (let table of tables.items) {
            table.columns.load("items/name"); // Carica i nomi delle colonne
            await context.sync(); // Sincronizza prima di accedere ai dati

            let columnNames = table.columns.items.map(col => col.name);

            // Controlla se l'intestazione è esattamente ["Chirurghi", "Turni"]
            if (JSON.stringify(columnNames) === JSON.stringify(["Chirurghi", "Turni"])) {
                let tableRange = table.getRange(); // Ottiene l'intervallo della tabella
                tableRange.format.autofitColumns();
                tableRange.format.columnWidth = 48;
                await context.sync(); // Assicuriamoci che la formattazione sia applicata

                // Dopo aver effettuato l'autofit, possiamo eliminare la tabella
                table.delete(); // Elimina la tabella
                console.log("Tabella eliminata con intestazione Chirurghi - Turni");
            }
        }

        return context.sync(); // Ultima sincronizzazione per garantire che tutto sia stato applicato correttamente
    }).catch(function (error) {
        console.error("Error: " + error);
    });
}

/*
async function runClingo() {
    try {
        // Verifica che 'clingo' sia definito
        console.log("Verifica se 'clingo' è definito:", typeof clingo);

        // Verifica se clingo è definito
        if (typeof clingo === 'undefined') {
            throw new Error("clingo is not defined. Please ensure that clingo-wasm is loaded.");
        }

        // Inizializza Clingo con il file WASM (opzionale, se necessario)
        await clingo.init("https://cdn.jsdelivr.net/npm/clingo-wasm@0.2.1/dist/clingo.wasm");

        // Esegui un programma di esempio su Clingo
        const result1 = await clingo.run("a. b :- a.");
        console.log("Risultato 1:", result1);

        const result2 = await clingo.run("{a; b; c}.", 0);
        console.log("Risultato 2:", result2);

        // Mostra i risultati nel div con id "results"
        $('#results').html(`
            <p><strong>Risultato 1:</strong> ${JSON.stringify(result1)}</p>
            <p><strong>Risultato 2:</strong> ${JSON.stringify(result2)}</p>
        `);

    } catch (error) {
        console.error("Errore durante l'esecuzione di Clingo:", error);
        $('#results').html('<p>Si è verificato un errore durante l\'esecuzione di Clingo.</p>');
    }
}
*/

async function risolviClingo() {
    try {
        // 1. Lettura dati da Excel
        const datiTurni = await leggiDatiTurni();

        // Verifica se sono stati letti dei dati
        if (!datiTurni || datiTurni.length === 0) {
            console.error("Errore: Nessuna tabella trovata o nessun dato letto.");
            return; // Esci dalla funzione se non ci sono dati
        }

        console.log("Dati letti da Excel");

        // 2. Creazione File con dati da passare a clingo
        await scriviDatiToNuovoFoglio(datiTurni);

        // 3. Leggo i dati formattati
        let dati = await leggiDatiDalFoglio();

        // Verifica se i dati letti sono validi
        if (!dati || dati.length === 0) {
            console.error("Errore: Nessun dato formattato trovato nel foglio.");
            return; // Esci dalla funzione se non ci sono dati
        }

        console.log("Dati letti dal foglio:", dati);

        //4. Eseguire Clingo
        let risposta = await eseguiClingoWasm(dati); // Qui puoi eseguire Clingo con i dati

        //5. Mostri i risultati
        console.log(risposta);
        mostraRisultati(risposta);

    } catch (error) {
        console.error("Errore durante l'esecuzione: " + error);
    }
}

async function leggiDatiTurni() {
    return Excel.run(async (context) => {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var tables = sheet.tables; // Ottiene tutte le tabelle nel foglio
        tables.load("items/name"); // Carica i nomi delle tabelle

        await context.sync(); // Sincronizza per ottenere i dati

        let datiTurni = []; // Array per raccogliere i dati

        // Scorri tutte le tabelle
        for (let table of tables.items) {
            table.columns.load("items/name"); // Carica i nomi delle colonne
            await context.sync(); // Sincronizza prima di accedere ai dati

            let columnNames = table.columns.items.map(col => col.name);

            // Controlla se l'intestazione è esattamente ["Chirurghi", "Turni"]
            if (JSON.stringify(columnNames) === JSON.stringify(["Chirurghi", "Turni"])) {
                const range = table.getDataBodyRange(); // Ottieni i dati
                range.load("values");
                await context.sync();

                // Aggiungi i valori letti nella variabile datiTurni
                datiTurni = range.values;
                break; // Interrompi il ciclo dopo aver trovato la tabella corretta
            }
        }

        // Restituisci i dati letti (se trovati)
        return datiTurni;
    });
}

async function scriviDatiToNuovoFoglio(dati) {
    return Excel.run(async (context) => {
        // Aggiungi un nuovo foglio di lavoro
        const nuovoFoglio = context.workbook.worksheets.add("Risultati Formattati");

        // Crea l'intervallo dinamicamente in base al numero di righe
        const intervalloRisultati = nuovoFoglio.getRange("A1").getResizedRange(dati.length - 1, 0); // Numero di righe = dati.length, una colonna (colonna A)

        // Formatta i dati nel formato "(nome chirurgo, turno)"
        const datiFormattati = dati.map(row => {
            return [`(${row[0]}, ${row[1]}).`]; // Formatta ogni riga come "(nome chirurgo, turno)"
        });

        // Scrivi i dati formattati nel nuovo foglio
        intervalloRisultati.values = datiFormattati;

        await context.sync(); // Sincronizza per applicare le modifiche
        console.log("Dati formattati e scritti nel nuovo foglio.");
    });
}

async function eseguiClingoWasm(datiFormattati) {
    try {
        // Inizializza Clingo con il file WASM (opzionale, se necessario)
        await clingo.init("https://cdn.jsdelivr.net/npm/clingo-wasm@0.2.1/dist/clingo.wasm");

        // Prepara il programma di Clingo di base
        const scriptClingo = `
giorno(lun).
giorno(mar).
giorno(mer).
giorno(giov).

1={orario(Chirurgo,G):chirurgo(Chirurgo,_)} :- giorno(G).
        `;

        // Modifica i dati letti in regole valide per Clingo
        const datiClingo = datiFormattati.map((entry) => {
            // Parsa la stringa come (nome, turno) e formatta correttamente
            const match = entry.match(/^\(([^,]+),\s*([^,]+)\)\.$/);
            if (match) {
                const chirurgo = match[1].trim().replace(/['"]+/g, ''); // Rimuovi eventuali virgolette
                const turno = match[2].trim().replace(/['"]+/g, ''); // Rimuovi eventuali virgolette

                // Verifica se ci sono dati validi (evita vuoti)
                if (chirurgo && turno) {
                    return `
chirurgo('${chirurgo}', ${turno}).
                `;
                }
            }
            return ''; // Se il formato non è valido o vuoto, ritorna una stringa vuota
        }).join("\n");

        // Combina il programma di Clingo con i dati formattati
        const programmaCompleto = scriptClingo + "\n" + datiClingo;

        console.log("Programma completo di Clingo:", programmaCompleto);

        // Esegui il programma di Clingo
        const risultato = await clingo.run(programmaCompleto);
        console.log("Risultato di Clingo:", risultato);

        return risultato;
    } catch (error) {
        console.error("Errore durante l'esecuzione di Clingo:", error);
    }
}

async function leggiDatiDalFoglio() {
    return Excel.run(async (context) => {
        // Ottieni il foglio "Risultati Formattati"
        const sheet = context.workbook.worksheets.getItem("Risultati Formattati");

        // Partiamo dalla cella A1
        let riga = 1; // Comincia dalla riga 1
        let datiLetti = [];

        while (true) {
            // Ottieni la cella corrente nella colonna A (ad esempio, A1, A2, A3, ...)
            const range = sheet.getRange(`A${riga}`);
            range.load("values");

            await context.sync(); // Sincronizza per caricare i dati dalla cella

            // Se la cella è vuota, esci dal ciclo
            if (range.values[0][0] === "" || range.values[0][0] === null) {
                break; // Esci dal ciclo se la cella è vuota
            }

            // Aggiungi il valore della cella all'array dei dati
            datiLetti.push(range.values[0][0]);

            // Incrementa la riga per leggere la successiva
            riga++;
        }

        // Restituisci i dati letti
        return datiLetti;
    });
}

async function mostraRisultati(risultatoClingo) {
    try {
        await Excel.run(async (context) => {
            const fogli = context.workbook.worksheets;
            const foglioEsistente = fogli.getItemOrNullObject("Risultati Formattati");
            await context.sync();

            if (!foglioEsistente.isNullObject) {
                foglioEsistente.delete();
            }

            const nuovoFoglio = context.workbook.worksheets.add("Orario");

            // Estrai i witnesses
            const witnesses = risultatoClingo.Call[0].Witnesses;

            // Prepara i dati
            let datiFormattati = [];

            if (witnesses && witnesses.length > 0) {
                // Estrai i valori dal primo witness
                datiFormattati = witnesses[0].Value.map(valore => [valore]);
            }

            // Aggiungi intestazione
            const intestazioni = [["Risultati"]];
            datiFormattati = intestazioni.concat(datiFormattati);

            // Ottieni l'intervallo
            const intervalloRisultati = nuovoFoglio.getRange(`A1:A${datiFormattati.length}`);

            // Imposta i valori
            intervalloRisultati.values = datiFormattati;

            await context.sync();
            console.log("Nuovo foglio con i risultati di Clingo creato.");
        });
    } catch (error) {
        console.error("Errore durante la sostituzione del foglio:", error);
        console.error("Dettagli errore:", JSON.stringify(error));
    }
}