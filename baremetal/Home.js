
    let cellToHighlight;
    let messageBanner;

    // Initialisierung, wenn Office JS und JQuery gelesen werden.
    Office.onReady(() => {
        $(() => {
            // Initialisieren Sie den Benachrichtigungsmechanismus der Office Fabric-Benutzeroberfläche, und blenden Sie ihn aus.
            const element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            // Wenn nicht Excel 2016 oder höher verwendet wird, Fallbacklogik verwenden.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("Dieses Beispiel zeigt den Wert der Zellen an, die Sie in der Tabelle ausgewählt haben.");
                $('#button-text').text("Anzeigen");
                $('#button-desc').text("Zeigt die Auswahl an.");

                $('#highlight-button').on('click',displaySelectedCells);
                return;
            }

            $("#template-description").text("Dieses Beispiel hebt den größten Wert aus den Zellen hervor, die Sie in der Tabelle ausgewählt haben.");
            $('#button-text').text("Hervorheben");
            $('#button-desc').text("Hebt die größte Zahl hervor.");
                
            loadSampleData();

            // Fügt einen Klickereignishandler für die Hervorhebungsschaltfläche hinzu.
            $('#highlight-button').on('click',highlightHighestValue);
        });
    });

    async function loadSampleData() {
        const values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                // Schreiben von Beispielwerten in einen Bereich im aktiven Arbeitsblatt
                sheet.getRange("B3:D5").values = values;
                await context.sync();
            });
        } catch (error) {
            errorHandler(error);
        }
    }

    async function highlightHighestValue() {
        try {
            await Excel.run(async (context) => {
                const sourceRange = context.workbook.getSelectedRange().load("values, rowCount, columnCount");

                await context.sync();
                let highestRow = 0;
                let highestCol = 0;
                let highestValue = sourceRange.values[0][0];

                // Sucht nach der hervorzuhebenden Zelle.
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

                // Hebt die Zelle hervor.
                cellToHighlight.format.fill.color = "orange";
                cellToHighlight.format.font.bold = true;
                await context.sync;
            });
        } catch (error) {
            errorHandler(error);
        }
    }

    async function displaySelectedCells() {
        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.load("text");
                await context.sync();
                const textValue = range.text.toString();
                showNotification('Der ausgewählte Text lautet:', '"' + textValue + '"');
            });
        } catch (error) {
            errorHandler(error);
        }
    }

    // Eine Hilfsfunktion zur Behandlung von Fehlern.
    function errorHandler(error) {
        // Stellen Sie immer sicher, dass kumulierte Fehler abgefangen werden, die bei der Ausführung von "Excel.run" auftreten.
        showNotification("Fehler", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Eine Hilfsfunktion zum Anzeigen von Benachrichtigungen.
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
