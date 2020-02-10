"use strict"
$(document).ready(function () {
    $(".table").hide();
})

$("#btn-importar").bind("click", function () {
    var reader = new FileReader();
    var file = $("#input-excel")[0].files[0];

    if (file) {
        reader.readAsArrayBuffer(file);
        reader.onload = function (e) {
            var data = new Uint8Array(reader.result);
            var wb = XLSX.read(data, { type: 'array' });

            //em sheetName é o campo armazena o nome da aba do excel, 
            // aqui está configurado para pegar a primeira aba, ou seja, 
            // posição 0;
            var sheetName = wb.SheetNames[0];
            var sheet = wb.Sheets[sheetName];
            sheet = XLSX.utils.sheet_to_json(sheet);

            createTable(sheet);
        };
    } else {
        alert("Não existe arquivo");
    }
});

function createTable(sheet) {
    $(".table").show();

    $.each(sheet, function (index, item) {
        $("#tbody").append(htmlTable(index, item));
    });
}


function htmlTable(index, item) {
    return `<tr>
        <th scope="row">${index}</th>
        <td>${item.tecnologia}</td>
        <td>${item.nivel}</td>
    </tr>`
}