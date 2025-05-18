// Die Initialisierungsfunktion muss bei jedem Laden einer neuen Seite ausgeführt werden.
Office.onReady(() => {
        // Wenn eine Initialisierung erfolgen muss, kann dies hier geschehen.
});

async function sampleFunction(event) { 
const values = [
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        try {
        await Excel.run(async (context) => {
                // Write sample values to a range in the active worksheet.
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.getRange("B3:D5").values = values;
                await context.sync();
        });
        } catch (error) {
        console.log(error.message);
        }
        // Das Aufrufen von event.completed ist erforderlich. event.completed teilt der Plattform mit, dass die Verarbeitung abgeschlossen wurde.
        event.completed();
}
