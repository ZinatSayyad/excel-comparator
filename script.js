function compareAndReplace() {
    const fileSheet1 = document.getElementById('fileSheet1').files[0];
    const fileSheet2 = document.getElementById('fileSheet2').files[0];

    if (!fileSheet1 || !fileSheet2) {
        alert('Please select both Excel sheets.');
        return;
    }

    const reader = new FileReader();

    reader.onload = function(e) {
        const dataSheet1 = new Uint8Array(e.target.result);
        const workbookSheet1 = XLSX.read(dataSheet1, { type: 'array' });
        const sheet1 = workbookSheet1.Sheets[workbookSheet1.SheetNames[0]];

        const dataSheet2 = new Uint8Array(reader.result);
        const workbookSheet2 = XLSX.read(dataSheet2, { type: 'array' });
        const sheet2 = workbookSheet2.Sheets[workbookSheet2.SheetNames[0]];

        const range = XLSX.utils.decode_range(sheet1['!ref']);

        for (let R = range.s.r; R <= range.e.r; R++) {
            const cellSheet1 = sheet1[XLSX.utils.encode_cell({ r: R, c: 0 })];
            const cellSheet2 = sheet2[XLSX.utils.encode_cell({ r: R, c: 0 })];

            if (cellSheet1 && cellSheet2 && cellSheet1.v === cellSheet2.v) {
                const cellBSheet2 = sheet2[XLSX.utils.encode_cell({ r: R, c: 1 })];
                const cellCSheet2 = sheet2[XLSX.utils.encode_cell({ r: R, c: 2 })];
                sheet1[XLSX.utils.encode_cell({ r: R, c: 1 })] = cellBSheet2;
                sheet1[XLSX.utils.encode_cell({ r: R, c: 2 })] = cellCSheet2;
            }
        }

        const workbookOutput = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbookOutput, sheet1, 'Sheet1');
        XLSX.writeFile(workbookOutput, 'output.xlsx');
        alert('Comparison and replacement completed! The output is saved as "output.xlsx".');
    };

    reader.readAsArrayBuffer(fileSheet1);
}
