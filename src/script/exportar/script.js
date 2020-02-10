var clients = [
    {
        "Name": "Otto Clay",
        "Age": 25,
        "Country": "United States",
        "Address": "Ap #897-1459 Quam Avenue",
    },
    {
        "Name": "Connor Johnston",
        "Age": 45,
        "Country": "Canada",
        "Address": "Ap #370-4647 Dis Av.",
    }
]


$("#btn-exportar").click(function () {
    criarExcel(clients)
})

function criarExcel(toSheet) {
    var nameSS = "Excel Down Teste";

    var wb = XLSX.utils.book_new();

    wb.Props = {
        Title: nameSS,
        Subject: "Desc Excel",
        Author: "Avanade",
        CreatedDate: new Date()
    };
    wb.SheetNames.push(nameSS);
    var ws_data = toSheet;

    if (ws_data.length > 0) {
        var ws = XLSX.utils.json_to_sheet(ws_data); //setar tamanho das colunas

        var numberParams = Object.keys(toSheet[0]).length;
        var wscols = [];

        for (var i = 0; i < numberParams; i++) {
            var ii = {
                wch: 20
            };
            wscols.push(ii);
        }

        ws['!cols'] = wscols; //criar filtros

        ws['!autofilter'] = {
            ref: "A1:AAA5000"
        };
        wb.Sheets[nameSS] = ws;
        var wbout = XLSX.write(wb, {
            bookType: 'xlsx',
            type: 'binary',
            cellStyles: true
        });
        saveAs(new Blob([_s2ab_Excel(wbout)], {
            type: "application/octet-stream"
        }), nameSS + '.xlsx');
    }
};

function _s2ab_Excel(s) {
    var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer

    var view = new Uint8Array(buf); //create uint8array as viewer

    for (var i = 0; i < s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xFF;
    } //convert to octet


    return buf;
}
