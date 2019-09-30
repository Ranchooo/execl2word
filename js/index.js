function selectFile() {
    document.getElementById('selectFile').click();
}

// 读取本地excel文件
function readWorkbookFromLocalFile(file, callback) {
    var reader = new FileReader();
    reader.onload = function (e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, { type: 'binary' });
        if (callback) callback(workbook);
    };
    reader.readAsBinaryString(file);
}

// 读取 excel文件
function outputWorkbook(workbook) {
    var sheetNames = workbook.SheetNames; // 工作表名称集合
    sheetNames.forEach(name => {
        var worksheet = workbook.Sheets[name]; // 只能通过工作表名称来获取指定工作表
        for (var key in worksheet) {
            // v是读取单元格的原始值
            console.log(key, key[0] === '!' ? worksheet[key] : worksheet[key].v);
        }
    });
}

function readWorkbook(workbook) {
    var sheetNames = workbook.SheetNames; // 工作表名称集合
    var worksheet = workbook.Sheets[sheetNames[0]]; // 这里我们只读取第一张sheet
    var csv = XLSX.utils.sheet_to_csv(worksheet);
    document.getElementById('result').innerHTML = csv2table(csv);
}

// 将csv转换成表格
function csv2table(csv) {
    var html = '<table id="table">';
    var rows = csv.split('\n');
    rows.pop(); // 最后一行没用的
    rows.forEach(function (row, idx) {
        var columns = row.split(',');
        columns.unshift(idx + 1); // 添加行索引
        if (idx == 0) { // 添加列索引
            html += '<tr>';
            for (var i = 0; i < columns.length; i++) {
                html += '<th>' + (i == 0 ? '' : String.fromCharCode(65 + i - 1)) + '</th>';
            }
            html += '</tr>';
        }
        html += '<tr>';
        columns.forEach(function (column) {
            html += '<td>' + column + '</td>';
        });
        html += '</tr>';
    });
    html += '</table>';
    return html;
}

// 监听选择文件按钮
$(function () {
    document.getElementById('selectFile').addEventListener('change', function (e) {
        var files = e.target.files;
        if (files.length == 0) return;
        var f = files[0];
        if (!/\.xlsx$/g.test(f.name)) {
            alert('仅支持读取xlsx格式！');
            return;
        }
        // 显示当前操作文件名
        let html = f.name;
        document.getElementById('showFileName').innerHTML = html;
        readWorkbookFromLocalFile(f, function (workbook) {
            readWorkbook(workbook);
        });
    });
});

function loadRemoteFile(url) {
    readWorkbookFromRemoteFile(url, function (workbook) {
        readWorkbook(workbook);
    });
}
// 数字转中文读数
function changeNum(num) {
    num = parseInt(num);
    let numChar = ['〇', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];
    let newNum = "";
    if (num <= 10) {
        return numChar[num];
    } else if (num > 10 && num <= 19) {
        newNum = '十' + numChar[parseInt(num % 10)];
        return newNum;
    } else if (num >= 20 && num <= 99) {
        if (parseInt(num % 10) == 0) {
            newNum = numChar[parseInt((num % 100) / 10)] + '十';
            return newNum;
        } else {
            newNum = numChar[parseInt((num % 100) / 10)] + '十' + numChar[parseInt(num % 10)];
            return newNum;
        }
    } else if (num >= 100 && num <= 999) {
        newNum = numChar[parseInt((num % 1000) / 100)] + numChar[parseInt((num % 100) / 10)] + numChar[parseInt(num % 10)];
        return newNum;
    } else if (num >= 1000 && num <= 9999) {
        newNum = numChar[parseInt((num % 10000) / 1000)] + numChar[parseInt((num % 1000) / 100)] + numChar[parseInt((num % 100) / 10)] + numChar[parseInt(num % 10)];
        return newNum;
    } else {
        // ToDo
        return;
    }
}
// 预览word
function showWord() {
    document.getElementById('result').classList.add('hidden');
    var tableList = document.getElementById("table");
    var str = "";
    // 获取table中的某一列的值
    // alert(tableList.rows[728].cells[1].innerHTML);
    var html = '<div>';
    var ke = 1;
    var shu = 1;
    for (var i = 1; i < tableList.rows.length; i++) {
        if (i == 1) {
            continue;
        } else if (i == 2) {
            html += '<p style="text-align:center;font-weight:bold">' + changeNum(ke) + '、' + tableList.rows[i].cells[1].innerHTML + '</p>';
            html += '<p style="font-weight:bold">（' + changeNum(shu) + '）' + tableList.rows[i].cells[2].innerHTML + ' ' + tableList.rows[i].cells[3].innerHTML + '</p>';
            html += '<p>' + (i - 1) + '. ' + tableList.rows[i].cells[4].innerHTML + ' ' + tableList.rows[i].cells[5].innerHTML + ' ' + tableList.rows[i].cells[6].innerHTML + '</p>';
            ke++;
            shu++;
        } else if (tableList.rows[i - 1].cells[1].innerHTML == tableList.rows[i].cells[1].innerHTML) {
            if (tableList.rows[i - 1].cells[2].innerHTML == tableList.rows[i].cells[2].innerHTML) {
                html += '<p>' + (i - 1) + '. ' + tableList.rows[i].cells[4].innerHTML + ' ' + tableList.rows[i].cells[5].innerHTML + ' ' + tableList.rows[i].cells[6].innerHTML + '</p>';
            } else {
                html += '<p style="font-weight:bold">（' + changeNum(shu) + '）' + tableList.rows[i].cells[2].innerHTML + ' ' + tableList.rows[i].cells[3].innerHTML + '</p>';
                html += '<p>' + (i - 1) + '. ' + tableList.rows[i].cells[4].innerHTML + ' ' + tableList.rows[i].cells[5].innerHTML + ' ' + tableList.rows[i].cells[6].innerHTML + '</p>';
                shu++;
            }
        } else {
            html += '<p style="text-align:center;font-weight:bold">' + changeNum(ke) + '、' + tableList.rows[i].cells[1].innerHTML + '</p>';
            html += '<p style="font-weight:bold">（' + changeNum(shu) + '）' + tableList.rows[i].cells[2].innerHTML + ' ' + tableList.rows[i].cells[3].innerHTML + '</p>';
            html += '<p>' + (i - 1) + '. ' + tableList.rows[i].cells[4].innerHTML + ' ' + tableList.rows[i].cells[5].innerHTML + ' ' + tableList.rows[i].cells[6].innerHTML + '</p>';
            ke++;
            shu++;
        }
    }
    html += '</div>';

    document.getElementById('show').innerHTML = html;
}

// 导出word文件
function exportWord() {
    let fielname = prompt("请输入导出保存的文件名：", "ExeclToWord");
    if (fielname == null || fielname == "") {
        return;
    } else {
        $("#show").wordExport(fielname);
    }
}

// 重置页面
function resetPage() {
    if (window.confirm('重置会丢失当前文档编辑区的信息，是否重置？')) {
        location.reload();
    }
}