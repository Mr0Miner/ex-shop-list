const filePath = document.getElementById('file-path').value;

document.addEventListener('DOMContentLoaded', function () {
    fetch(filePath)
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            let totalProfit = 0;
            let totalProducts = 0;
            let totalMult = 0;
            let totalDiff = 0;
            const rows = [];

            workbook.SheetNames.forEach(sheetName => {
                const sheet = workbook.Sheets[sheetName];
                // اضافه کردن لاگ برای بررسی مقادیر
                console.log(sheet['E363'], sheet['F363'], sheet['G363'], sheet['H363']);
                
                totalProfit += sheet['E363'] ? sheet['E363'].v : 0;
                totalProducts += sheet['F363'] ? sheet['F363'].v : 0;
                totalMult += sheet['G363'] ? sheet['G363'].v : 0;
                totalDiff += sheet['H363'] ? sheet['H363'].v : 0;

                const range = XLSX.utils.decode_range(sheet['!ref']);
                for (let R = range.s.r; R <= range.e.r; ++R) {
                    const row = [];
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const cell_address = { c: C, r: R };
                        const cell_ref = XLSX.utils.encode_cell(cell_address);
                        const cell = sheet[cell_ref];
                        row.push(cell ? cell.v : "");
                    }
                    rows.push(row);
                }
            });

            displayTable(rows);

            document.getElementById('total-profit').textContent = totalProfit.toLocaleString();
            document.getElementById('total-products').textContent = totalProducts.toLocaleString();
            document.getElementById('total-mult').textContent = totalMult.toLocaleString();
            document.getElementById('total-diff').textContent = totalDiff.toLocaleString();

            const totalProfitWords = numberToWords(totalProfit);
            const totalProductsWords = numberToWords(totalProducts);
            const totalMultWords = numberToWords(totalMult);
            const totalDiffWords = numberToWords(totalDiff);

            document.getElementById('total-profit-words').textContent = totalProfitWords;
            document.getElementById('total-products-words').textContent = totalProductsWords;
            document.getElementById('total-mult-words').textContent = totalMultWords;
            document.getElementById('total-diff-words').textContent = totalDiffWords;
        })
        .catch(error => console.error(error));
});

document.getElementById('file-path').addEventListener('change', function() {
    const filePath = this.files[0]; // دریافت فایل انتخاب شده
    if (filePath) {
        const reader = new FileReader();
        reader.onload = function(event) {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            // ادامه کد برای پردازش داده‌ها
            let totalProfit = 0;
            let totalProducts = 0;
            let totalMult = 0;
            let totalDiff = 0;
            const rows = [];

            workbook.SheetNames.forEach(sheetName => {
                const sheet = workbook.Sheets[sheetName];
                // اضافه کردن لاگ برای بررسی مقادیر
                console.log(sheet['E1'], sheet['F1'], sheet['G1'], sheet['H1']);
                
                totalProfit += sheet['E1'] ? sheet['E1'].v : 0; // تغییر به E1
                totalProducts += sheet['F1'] ? sheet['F1'].v : 0; // تغییر به F1
                totalMult += sheet['G1'] ? sheet['G1'].v : 0; // تغییر به G1
                totalDiff += sheet['H1'] ? sheet['H1'].v : 0; // تغییر به H1

                const range = XLSX.utils.decode_range(sheet['!ref']);
                for (let R = range.s.r; R <= range.e.r; ++R) {
                    const row = [];
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                        const cell_address = { c: C, r: R };
                        const cell_ref = XLSX.utils.encode_cell(cell_address);
                        const cell = sheet[cell_ref];
                        row.push(cell ? cell.v : "");
                    }
                    rows.push(row);
                }
            });

            displayTable(rows);
            // ادامه کد برای نمایش مقادیر کل
        };
        reader.readAsArrayBuffer(filePath); // خواندن فایل به عنوان ArrayBuffer
    }
});

function displayTable(rows) {
    const tbody = document.querySelector('#excel-table tbody');
    tbody.innerHTML = '';

    rows.slice(1).forEach(row => {
        if (row[0]) {
            const tr = document.createElement('tr');
            if (row.length >= 10) {
                [row[0], row[1], row[3], row[2], row[4], row[5], row[9]].forEach((cell, index) => {
                    const td = document.createElement('td');
                    td.textContent = cell || "";
                    if (index === 5) {
                        if (cell === 0 || cell === "") {
                            td.classList.add('cell-out-of-stock');
                            td.textContent = "ناموجود";
                        } else if (cell === 1) {
                            td.classList.add('cell-low-stock');
                            td.textContent = "1";
                        } else if (cell > 1) {
                            td.classList.add('cell-stock');
                            td.textContent = cell;
                        }
                    } else if (index === 6) {
                        td.textContent = cell || "";

                        if (cell > 50) {
                            td.classList.add('cell-stock');
                        } else if (cell <= 50) {
                            td.classList.add('cell-low-stock');
                        }
                    } else if (index === 3) {
                        td.textContent = cell || "";

                        if (cell >= 50) {
                            td.classList.add('cell-stock');
                        } else if (cell <= 50) {
                            td.classList.add('cell-low-stock');
                        }
                    } else if (index === 4) {
                        td.textContent = cell || "";

                        if (cell >= 50) {
                            td.classList.add('cell-stock');
                        } else if (cell <= 50) {
                            td.classList.add('cell-low-stock');
                        }
                    }
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            }
        }
    });
}

function searchTable() {
    const searchInput = document.getElementById('search-input').value.toLowerCase();
    const searchColumn = document.getElementById('search-column').value;
    const rows = document.querySelectorAll('#excel-table tbody tr');

    rows.forEach(row => {
        const cell = row.cells[searchColumn];
        const text = cell ? cell.innerText.toLowerCase() : "";
        row.style.display = text.includes(searchInput) ? '' : 'none';
    });
}

function numberToWords(num) {
    return num.toString().num2persian();
}

function openPage(pageName, elmnt) {
    var i, tabcontent, tablinks;
    tabcontent = document.getElementsByClassName("tabcontent");
    for (i = 0; i < tabcontent.length; i++) {
        tabcontent[i].style.display = "none";
    }
    tablinks = document.getElementsByClassName("tablink");
    for (i = 0; i < tablinks.length; i++) {
        tablinks[i].style.backgroundColor = "";
    }
    document.getElementById(pageName).style.display = "block";
    elmnt.style.backgroundColor = "#555";
}

document.getElementById("defaultOpen").click();

function toggleVisibility(id) {
    const element = document.getElementById(id);
    if (element) {
        element.style.display = (element.style.display === 'none' || element.style.display === '') ? 'inline' : 'none'; // تغییر وضعیت نمایش
    }
}