# Exportação e importação de Excel
Projeto utilizando a biblioteca SheetJS, para a realização de importação e exportação de excel.

# Scripts
Para realizar a importação e/ou exportação é necessário de dois script's, exatamente na ordem a seguir:
```html
<script src="../src/script/import-excel/cdn/FileSaver.min.js"></script>
<script src="../src/script/import-excel/cdn/xlsx.full.min.js"></script>
```

# Exportação
Na exportação a bilioteca espera receber um array de objetos conforme exemplo abaixo:
```js
var clients = [
    { "Name": "Otto Clay", "Age": 25, "Country": "United States", "Address": "Ap #897-1459 Quam Avenue", "Married": false },
    { "Name": "Connor Johnston", "Age": 45, "Country": "Canada", "Address": "Ap #370-4647 Dis Av.", "Married": true }
];
```
