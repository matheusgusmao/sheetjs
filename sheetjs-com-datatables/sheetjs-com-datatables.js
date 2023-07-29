var workbookSheet;

function dataTable(dados) {

    const rows = [];

    for (const row of dados) {

        rows.push({
            campo1:     row['campo1'],
            campo2:     row['campo2'],
            campo3:     row['campo3'],
            campo4:     row['campo4'],
            campo5:     row['campo5'],
            campo6:     row['campo6'],
            campo7:     row['campo7'],
            campo8:     row['campo8'],
            campo9:     row['campo9'],
            campo10:    row['campo10'],
            campo11:    row['campo11'],
            campo12:    row['campo12'],
            campo13:    row['campo13'],
            campo14:    row['campo14'],
            campo15:    row['campo15']
        });

    };

    $('#table').DataTable({

        data: rows,
        responsive: true,
        scrollX: true,
        columns: [
            {
                data: 'campo1',
                title: 'campo1',
            },
            {
                data: 'campo2',
                title: 'campo2',
            },
            {
                data: 'campo3',
                title: 'campo3',
            },
            {
                data: 'campo4',
                title: 'campo4',
            },
            {
                data: 'campo5',
                title: 'campo5',
            },
            {
                data: 'campo6',
                title: 'campo6',
            },
            {
                data: 'campo7',
                title: 'campo7',
            },
            {
                data: 'campo8',
                title: 'campo8',
            },
            {
                data: 'campo9',
                title: 'campo9',
            },
            {
                data: 'campo10',
                title: 'campo10',
            },
            {
                data: 'campo11',
                title: 'campo11',
            },
            {
                data: 'campo12',
                title: 'campo12',
            },
            {
                data: 'campo13',
                title: 'campo13',
            },
            {
                data: 'campo14',
                title: 'campo14',
            },
            {
                data: 'campo15',
                title: 'campo15',
            }
        ],
        "drawCallback": function (settings) {
            createSpreadsheet(settings.aiDisplay, rows);
        }
    });

}

function createSpreadsheet(rowsExcel, rows) {

    var ws_data = [
        [
            'campo1',
            'campo2',
            'campo3',
            'campo4',
            'campo5',
            'campo6',
            'campo7',
            'campo8',
            'campo9',
            'campo10',
            'campo11',
            'campo12',
            'campo13',
            'campo14',
            'campo15'
        ]
    ];

    rowsExcel.forEach(idx => {

        if (rows.includes(rows[idx])) {

            ws_data.push([
                rows[idx].campo1,
                rows[idx].campo2,
                rows[idx].campo3,
                rows[idx].campo4,
                rows[idx].campo5,
                rows[idx].campo6,
                rows[idx].campo7,
                rows[idx].campo8,
                rows[idx].campo9,
                rows[idx].campo10,
                rows[idx].campo11,
                rows[idx].campo12,
                rows[idx].campo13,
                rows[idx].campo14,
                rows[idx].campo15
            ]);

        }

    });

    const worksheet = XLSX.utils.json_to_sheet(ws_data);
    workbookSheet = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbookSheet, worksheet, 'mySheet');
    XLSX.utils.sheet_add_aoa(worksheet, [['mySheet']]);

    var merge = { s: { r: 0, c: 0 }, e: { r: 0, c: 14 } };
    if (!worksheet['!merges']) worksheet['!merges'] = [];
    worksheet['!merges'].push(merge);

    for (i in worksheet) {

        if (typeof (worksheet[i]) != 'object') continue;

        let cell = XLSX.utils.decode_cell(i);

        worksheet[i].s = {
            font: {
                name: 'Arial'
            },
            alignment: {
                vertical: 'center',
                horizontal: 'center',
                wrapText: '1',
            },
            border: {
                right: {
                    style: 'thin',
                    color: '000000'
                },
                left: {
                    style: 'thin',
                    color: '000000'
                },
            }
        };

        if (cell.r % 2) {
            worksheet[i].s.fill = {
                patternType: 'solid',
                fgColor: { rgb: 'D1D3E5' },
                bgColor: { rgb: 'D1D3E5' }
            };
        }

        if (cell.r == 1) {
            worksheet[i].s = {
                font: {
                    name: 'Arial',
                    color: { rgb: "FFFFFF" }
                },
                alignment: {
                    vertical: 'center',
                    horizontal: 'center',
                    wrapText: '1',
                },
                border: {
                    right: {
                        style: 'thin',
                        color: '000000'
                    },
                    left: {
                        style: 'thin',
                        color: '000000'
                    },
                }
            };

            worksheet[i].s.fill = {
                patternType: 'solid',
                fgColor: { rgb: '002F65' },
                bgColor: { rgb: '002F65' }
            };
        }
    }

}