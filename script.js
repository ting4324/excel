let table = document.querySelector(".table");
let table2 = document.querySelector(".table2");

(async () => {

    let workbook = XLSX.read(await (await fetch("./Table_Input.xlsx")).arrayBuffer());
    let sheet = workbook.Sheets[workbook.SheetNames[0]];

    let html = XLSX.utils.sheet_to_html(sheet, { header: ""});
    table.innerHTML = `<h3>Table 1</h3>${html}`;

    function getCellValue(sheet, cell) {
        return sheet[cell] ? parseFloat(sheet[cell].v) : 0;
    }

    let A5 = getCellValue(sheet, 'B6');    
    let A20 = getCellValue(sheet, 'B21');  
    let A15 = getCellValue(sheet, 'B16');  
    let A7 = getCellValue(sheet, 'B8');   
    let A13 = getCellValue(sheet, 'B14');  
    let A12 = getCellValue(sheet, 'B13');  

    let alpha = A5 + A20;               
    let beta = A15 / A7;                
    let charlie = A13 * A12;             

    table2.innerHTML = `
        <h3>Table 2</h3>
        <table>
            <tr>
                <th>Category</th>
                <th>Value</th>
            </tr>
            <tr>
                <td>Alpha</td>
                <td>${alpha}</td>
            </tr>
            <tr>
                <td>Beta</td>
                <td>${beta}</td> 
            </tr>
            <tr>
                <td>Charlie</td>
                <td>${charlie}</td>
            </tr>
        </table>
    `;
})()
