var _target = document.getElementById('drop');
var _file = document.getElementById('file');
var _grid = document.getElementById('grid');
var currentSheetIdx;
var _tableWrapper = document.getElementById('table-wrapper');
var _tableBody = document.getElementById('table-body');
var _btnDownload = document.getElementById('btn-download');

var _workstart = function() {}
var _workend = function() {
    currentSheetIdx = 0;
    _btnDownload.classList.remove('d-none')
    _tableWrapper.classList.remove('d-none')
}

/** Alerts **/
var _badfile = function() {
    alertify.alert('This file does not appear to be a valid Excel file.  If we made a mistake, please send this file to <a href="mailto:dev@sheetjs.com?subject=I+broke+your+stuff">dev@sheetjs.com</a> so we can take a look.', function() {});
};

var _pending = function() {
    alertify.alert('Please wait until the current file is processed.', function() {});
};

var _large = function(len, cb) {
    alertify.confirm("This file is " + len + " bytes and may take a few moments.  Your browser may lock up during this process.  Shall we play?", cb);
};

var _failed = function(e) {
    console.log(e, e.stack);
    alertify.alert('We unfortunately dropped the ball here.  Please test the file using the <a href="/js-xlsx/">raw parser</a>.  If there are issues with the file processor, please send this file to <a href="mailto:dev@sheetjs.com?subject=I+broke+your+stuff">dev@sheetjs.com</a> so we can make things right.', function() {});
};


var _onsheet = function(json, sheetnames, select_sheet_cb) {
    var name = sheetnames[currentSheetIdx]
    var map = headerMapping[name]
    var arr
    var idxMap = []
    var rowFragment = '';
    if (map) {
        var L = 0;
        json.forEach(function(r) { if (L < r.length) L = r.length; });
        for (var i = json[0].length; i < L; ++i) {
            json[0][i] = "";
        }
        json.forEach((row, idx) => {
            if (idx === 0) {
                row.forEach((col, idx) => {
                    idxMap.push(baseMapping[map[col]])
                });
                return false;
            }
            if(row.length===0) return false;
            arr = [name, ...new Array(40)]
            row.forEach((col, idx) => {
                arr[idxMap[idx]] = col
            });
            rowFragment += `
                <tr>
                    ${arr.map((val, idx) => {
                        return `<td>${val?val : '--'}</td>`
                    }).join('')}
                </tr>`
        })
    } else {
        alert(`對應表內沒有『${name}』`);
    }
    //console.log(idxMap);
    _tableBody.innerHTML += (rowFragment)
    currentSheetIdx++
    if (currentSheetIdx < sheetnames.length) {
        select_sheet_cb(currentSheetIdx)
    }
};

function download(type, fn, dl) {
    var elt = document.getElementById('table');
    var wb = XLSX.utils.table_to_book(elt, { sheet: "Sheet JS" });
    return dl ?
        XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }) :
        XLSX.writeFile(wb, fn || ((+ new Date()) + '.' + (type || 'xlsx')));
}
/** Drop it like it's hot **/
DropSheet({
    file: _file,
    drop: _target,
    on: {
        workstart: _workstart,
        workend: _workend,
        sheet: _onsheet
    },
    errors: {
        badfile: _badfile,
        pending: _pending,
        failed: _failed,
        large: _large
    }
})