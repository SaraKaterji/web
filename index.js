let values = [];
let form = document.getElementById('form');
form.addEventListener("submit", function (e) {
    e.preventDefault();
    values = [];
    const formData = new FormData(form);
    for (const entry of formData.entries()) {
        values.push(entry[1]);
    }
    var searchstring = values[0] + values[1];
    readExcel(searchstring);
});

function readExcel(searchstring) {
    var xhr = new XMLHttpRequest();
    xhr.open('GET', 'https://sarakaterji.github.io/web/t.xlsx', true);
    xhr.responseType = 'blob';
    xhr.onload = function (e) {
        if (this.status == 200) {
            var blob = this.response;
            excelFileToJSON(blob, searchstring);
        }
    };
    xhr.onerror = function (e) {
        alert("Error " + e.target.status + " occurred while receiving the document.");
    };
    xhr.send();
}

//Method to read excel file and convert it into JSON 
function excelFileToJSON(file, searchstring) {
    try {
        var reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function (e) {
            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });
            var sheet_3 = workbook.SheetNames[2];
            var lookuptable = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_3]);
            var sheet_2 = workbook.SheetNames[1];
            var messages = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_2]);
            lookupJson(lookuptable, messages, searchstring);
        }
    } catch (e) {
        console.error(e);
    }
}

function lookupJson(lookuptable, messages, searchstring) {
    searchstring = searchstring.replaceAll(' ','').toUpperCase();
    var energielabel = document.getElementById("energielabel");
    var message = document.getElementById("message");
    var found = false;
    if (lookuptable.length > 0) {
        for (var i = 0; i < lookuptable.length; i++) {
            var row = lookuptable[i];
            if (row["searchstring"] == searchstring) {
                //energielabel.value = row["label"];
                for (var j = 0; j < messages.length; j++) {
                    if (messages[j]["result"] == row["label"]) {
                        message.innerHTML =  row["label"] + "<br/><br/>" + messages[j]["message"];
                        found = true;
                        break;
                    }
                }
                break;
            }
        }
        if (found == false) {
            //energielabel.value = "Not found";
            //message.value = "Not found";
            message.innerHTML = "Niet gevonden"
        }
    }
}
