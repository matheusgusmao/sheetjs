const createSpreadsheet = () => {

    var ws_data = [
        ['Campo1', 'Campo2', 'Campo3', 'Campo4', 'Campo5', 'Campo6', 'Campo7', 'Campo8', 'Campo9', 'Campo10', 'Campo11', 'Campo12', 'Campo13', 'Campo14', 'Campo15']
    ];

    var fields = {};

    document.querySelectorAll('.data-spreadsheet').forEach(el => {
        fields[el.id] = el.value;
    })

    var actualFields = [];

    // dinâmico - se atentar à ordenação dos campos:
    Object.entries(fields).forEach(([key, value]) => {
        actualFields.push(value);
    });

    ws_data.push(actualFields);

    const worksheet = XLSX.utils.json_to_sheet(ws_data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "mySheet");

    /* fix headers */
    XLSX.utils.sheet_add_aoa(worksheet, [["mySheet"]]);


    var merge = { s: { r: 0, c: 0 }, e: { r: 0, c: 14 } };
    if (!worksheet['!merges']) worksheet['!merges'] = [];
    worksheet['!merges'].push(merge);


    for (i in worksheet) {
        if (typeof(worksheet[i]) != "object") continue;
        let cell = XLSX.utils.decode_cell(i);
    
        worksheet[i].s = { // styling for all cells
            font: {
                name: "arial"
            },
            alignment: {
                vertical: "center",
                horizontal: "center",
                wrapText: '1', // any truthy value here
            },
            border: {
                right: {
                    style: "thin",
                    color: "000000"
                },
                left: {
                    style: "thin",
                    color: "000000"
                },
            }
        };
    
        if (cell.r == 0 ) { // first row
            worksheet[i].s.border.bottom = { // bottom border
                style: "thin",
                color: "000000"
            };
        }
    
        if (cell.r % 2) { // every other row
            worksheet[i].s.fill = { // background color
                patternType: "solid",
                fgColor: { rgb: "b2b2b2" },
                bgColor: { rgb: "b2b2b2" } 
            };
        }

        if (cell.r == 1) {
            worksheet[i].s.fill = {
                patternType: "solid",
                fgColor: { rgb: "002f65" },
                bgColor: { rgb: "002f65" }
            }
        }
    }

    XLSX.writeFile(workbook, "mySheet.xlsx");
}