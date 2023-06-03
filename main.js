
function read_file(file){
  return new Promise((resolve, reject) => {
    var reader = new FileReader();  
    reader.onload = () => {
      resolve(reader.result )
    };
    reader.onerror = reject;
    reader.readAsText(file);
  });
}


function save_xlsx(table) {
    var sheet = XLSX.utils.aoa_to_sheet(table);
    var book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, "result");
    XLSX.writeFile(book, "result.xlsx");
}


function convert_html(html_str) {
    var table = [];

    var dom_parser = new DOMParser();
    var doc = dom_parser.parseFromString(html_str, "text/html");

    var table_html = doc.getElementsByClassName("stg-flex-table")[0];
    for (var row_idx = 1; row_idx < table_html.children.length; row_idx++) {
        var row_html = table_html.children[row_idx];
        var row = [];
        for (var col_idx = 0; col_idx < row_html.children.length; col_idx++) {
            text = row_html.children[col_idx].innerText;
            if (col_idx == 2) {
                text = text.replace("\xa0₽", "");
                text = text.replace(" ", "");
            }
            row.push(text);
        }
        table.push(row);
    }

    return table;
}


function on_drop(event) {
    event.preventDefault();
    var files = event.dataTransfer.files;
    var promises = [];
    for(var i = 0; i < files.length; i++) {
        var file = files[i];
        promises.push(read_file(file));
    }
    Promise.all(promises).then((results) => {
        table = [];
        for(var i = 0; i < results.length; i++) {
            table = table.concat(convert_html(results[i]));
        }
        save_xlsx(table);

        var dropzone = document.getElementById("drop_zone");
        dropzone.classList.remove("dragover");
        dropzone.classList.add("not_dragover");
    }).catch((error) => {
        console.log(error);
    });
}

function on_drag_over(event) {
    event.preventDefault();
    var dropzone = document.getElementById("drop_zone");
    dropzone.classList.remove("not_dragover");
    dropzone.classList.add("dragover");
}

function on_drag_leave(event) {
    event.preventDefault();
    var dropzone = document.getElementById("drop_zone");
    dropzone.classList.remove("dragover");
    dropzone.classList.add("not_dragover");
}

window.onload = function() {
    var dropzone = document.getElementById("drop_zone");
    dropzone.ondrop = on_drop;
    dropzone.ondragover = on_drag_over;
    dropzone.ondragleave = on_drag_leave;
}
