var result = {};//almacenamos el JSON con el stock completo
var index = document.getElementById('index');
var form = document.getElementById('form');

var preview = document.getElementById('start-control');

//si el objeto result no está vacío
if (!Object.keys(result).length) {
    var btnContinue = document.getElementById('btnContinue');
    btnContinue.style.display = "block";
} else {
    alertDanger.style.display = "flex";
    h2.innerHTML = "No hay datos";
    p.innerHTML = "Comenzá un nuevo control de stock."
    }
function continueBtn() {
    index.style.display = 'none';
    preview.style.display = "flex";
    
}

function newBtn() {
    index.style.display = 'none';
    form.style.display = 'flex'
    localStorage.removeItem('resultSaved');
}

//Toma el nombre del archivo y imprime en pantalla
function cambiar() {
    var fileName = document.getElementById('file-upload').files[0].name;
    document.getElementById('info').innerHTML = fileName;
}

//Validaciones input file
function upload() {
    var files = document.getElementById('file-upload').files;
    if (files.length == 0) {
        alertDanger.style.display = "flex";
        h2.innerHTML = "¡Falta lo más importante!";
        p.innerHTML = "Tenés que cargar un archivo para poder comenzar"
        return;
    }
    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON(files[0]);
    } else {
        alertDanger.style.display = "flex";
        h2.innerHTML = "¡Oh no!";
        p.innerHTML = "No cargaste un archivo válido. Recordá que tiene que ser .XLS o .XLSX"
    }
}

function closePreview() {//cerar la tabla del stock completo
    tablePreview.style.display = 'none';
}
function viewTable() {//Mostramos la tabla del stock completo
    tablePreview.style.display = 'block';
}
var cant = 0;//nro del row
var table = document.getElementById('table-stock');//obtenemos la tabla para ver el stock completo
var tablePreview = document.getElementById('previewTable')//obtenemos la tabla donde se imprimen los resultados del scan


//bucle para imprimir el JSON en la tabla del stock completo
function tableStock(result) {
    result.Hoja1.forEach((a) => {
        let newRow = '<tr>' +
            '<td>' + cant + '</td>' +
            '<td>' + a.Código + '</td>' +
            '<td class="description">' + a.Descripción + '</td>' +
            '<td>' + a.ec + '</td>' +
            '<td>' + a.stock + '</td>' +
            '</tr>';
        cant = cant + 1;
        table.innerHTML += newRow;
    });
}

//convertimos el excel en JSON
function excelFileToJSON(file) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function (e) {

            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });

            workbook.SheetNames.forEach(function (sheetName) {
                var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                if (roa.length > 0) {
                    result[sheetName] = roa;
                }
            });

            preview.style.display = "flex";
            form.style.display = "none";

            tableStock(result);
            localStorage.setItem("resultSaved", JSON.stringify(result));
        }
    } catch (e) {
        console.error(e);
    }
}

var searching = [];//almacenamos los valores del input que se encontraron en el stock completo
let resultStorage = JSON.parse(localStorage.getItem('resultSaved'));





var inp = document.getElementById('search');//seleccionamos el input
var ubicacion = document.getElementById('ubication');
var almacen = document.getElementById('almacen');
var alertDanger = document.getElementById('alert');
var h2 = document.getElementById("h2");
var p = document.getElementById('p');
var btnAlert = document.getElementById('btn-alert');
var historyBg = document.getElementById('historyBg');
var back = document.getElementById('back');

function backPage() {
    preview.style.display = 'none';
    index.style.display = 'flex';
    form.style.display = "none";
}

function closeAlert() {
    alertDanger.style.display = 'none';
    inp.value = "";
    inp.focus();
}

function closeHistory() {
    historyBg.style.display = 'none';

}
var allInp = [];
const BreakError = {};
var jsonToExcel = JSON.stringify(allInp);
inp.addEventListener("input", () => {
    var inpValor = inp.value.toUpperCase();
    var ifIncludes;

    var tableInp = document.getElementById('tableInp');

    if (inpValor.length > 5) {
        if (ubicacion.value != "select" && almacen.value != "select") {
            try {
                resultStorage.Hoja1.forEach(j => {
                    ifIncludes = j.Código.includes(inpValor);
                    if (ifIncludes == true) {
                        throw BreakError;
                    }
                });
            } catch (error) {
                if (error !== BreakError) throw error;
            }
            if (ifIncludes == false && inpValor.length === 6) {
                alertDanger.style.display = "flex";
                h2.innerHTML = "Código inválido";
                p.innerHTML = "Por favor cargá un código válido."
                ifIncludes = "";
            }
            if (ifIncludes == true) {
                searching.push(inpValor);
                var count = 1;
                var unique = [];
                var repeat = [];
                var orderSum = [];

                var date = new Date();
                var hr = date.toLocaleTimeString();
                allInp.push({ code: inpValor, qty: 1, ubi: ubicacion.value, alm: almacen.value, time: hr })
                searching = searching.sort();

                for (var i = 0; i < searching.length; i++) {
                    let a = i + 1;
                    if (searching[a] == searching[i]) {
                        count++;
                    } else {
                        unique.push(searching[i]);
                        repeat.push({ cant: count, cantLog: cantLOG(searching[i]) });
                        count = 1;
                    }
                }
                for (let a = 0; a < unique.length; a++) {
                    orderSum.push({ code: unique[a], cantLog: repeat[a].cantLog, cant: repeat[a].cant });
                }

                var cant2 = 0;
                function tableStock2(orderSum) {
                    orderSum.forEach((k) => {
                        let newRow2 = '<tr id="row' + cant2 + '">' +
                            '<td>' + cant2 + '</td>' +
                            '<td id="code' + cant2 + '"onclick="history(code' + cant2 + ')">' + k.code + '</td>' +
                            '<td class="description">' + k.cantLog + '</td>' +
                            '<td>' + k.cant + '</td>' +
                            '<td id="dif' + cant2 + '">' + dif(k.cantLog, k.cant) + '</td>' +
                            '</tr>';
                        cant2 = cant2 + 1;
                        tableInp.innerHTML += newRow2;

                        var tableToExport = document.getElementById("2");
                        tableToExport.innerHTML = "";
                        allInp.forEach((b) => {
                            let newRowExport = '<tr>' +
                                '<td>' + b.code + '</td>' +
                                '<td>' + b.qty + '</td>' +
                                '<td>' + b.alm + '</td>' +
                                '<td>' + b.ubi + '</td>' +
                                '<td>' + b.time + '</td>' +
                                '</tr>';
                            tableToExport.innerHTML += newRowExport;

                        });
                    });
                }

                tableInp.innerText = "";
                tableStock2(orderSum);
                idColor();



                function dif(a, b) {
                    var diferencia = a - b;

                    return diferencia;
                }
                function idColor() {
                    for (let m = 0; m < cant2; m++) {
                        var idDif = document.getElementById('dif' + m);
                        var idRow = document.getElementById('row' + m);
                        if (idDif.innerHTML < 0) {
                            idDif.setAttribute("class", "rojo");
                        }
                        if (idDif.innerHTML == 0) {
                            idRow.setAttribute("class", "rowComplete");
                        }
                        if (idDif.innerHTML > 0) {
                            idDif.setAttribute("class", "naranja");
                        }

                    }

                }
                inp.value = "";

            }
        } else {
            if (inpValor.length > 5) {
                alertDanger.style.display = "flex";
                h2.innerHTML = "Campos vacios";
                p.innerHTML = "Por favor completá todos los campos."
            }

        }
    }
});



function cantLOG(index) {
    var cantLOGICA
    resultStorage.Hoja1.forEach(r => {
        if (index === r.Código) {
            cantLOGICA = parseInt(r.stock);
        }
    })
    return cantLOGICA;
}


function history(id) {
    historyBg.style.display = "flex";
    var cant3 = 0;
    var h2Histo = document.getElementById('h2History');
    h2Histo.innerHTML = id.innerHTML;
    var tableHistory = document.getElementById('tableHistory');
    tableHistory.innerText = "";
    allInp.forEach((m) => {
        if (m.code == id.innerHTML) {
            console.log(m.code)
            let newRowHistory = '<tr id="row' + cant3 + '">' +
                '<td>' + cant3 + '</td>' +
                '<td>' + m.code + '</td>' +
                '<td>' + m.qty + '</td>' +
                '<td>' + m.alm + '</td>' +
                '<td>' + m.ubi + '</td>' +
                '<td>' + m.time + '</td>' +
                '</tr>';
            tableHistory.innerHTML += newRowHistory;
            cant3++;
        }
    });
}

var tablesToExcel = (function () {

    var uri = 'data:application/vnd.ms-excel;base64,'
        , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>'
        , templateend = '</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head>'
        , body = '<body>'
        , tablevar = '<table>{table'
        , tablevarend = '}</table>'
        , bodyend = '</body></html>'
        , worksheet = '<x:ExcelWorksheet><x:Name>'
        , worksheetend = '</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>'
        , worksheetvar = '{worksheet'
        , worksheetvarend = '}'
        , base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))) }
        , format = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
        , wstemplate = ''
        , tabletemplate = '';

    return function (table, name, filename) {
        var tables = table;

        for (var i = 0; i < tables.length; ++i) {
            wstemplate += worksheet + worksheetvar + i + worksheetvarend + worksheetend;
            tabletemplate += tablevar + i + tablevarend;
        }

        var allTemplate = template + wstemplate + templateend;
        var allWorksheet = body + tabletemplate + bodyend;
        var allOfIt = allTemplate + allWorksheet;

        var ctx = {};
    
        if (allInp.length > 0) {
            for (var j = 0; j < tables.length; ++j) {
                ctx['worksheet' + j] = name[j];
            }
            
            for (var k = 0; k < tables.length; ++k) {
                var exceltable;
                if (!tables[k].nodeType) exceltable = document.getElementById(tables[k]);
                ctx['table' + k] = exceltable.innerHTML;
            }
            
            window.location.href = uri + base64(format(allOfIt, ctx));
        } else {
            alertDanger.style.display = "flex";
            h2.innerHTML = "No hay datos";
            p.innerHTML = "Comenzá con el control de stock antes de descargar."
        }

    }
})();