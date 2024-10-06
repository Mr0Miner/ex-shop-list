let filePath = document.getElementById('file-path').value;

document.addEventListener('DOMContentLoaded', function () {
    // تنظیم حالت پیشفرض به روز
    document.body.classList.add('day'); // اضافه کردن حالت روز
    document.querySelector('.container').classList.add('day'); // اضافه کردن حالت روز به container

    // اضافه کردن دکمه تغییر حالت
    const toggleButton = document.createElement('button');
    toggleButton.className = 'toggle-button'; // اضافه کردن کلاس برای استایل

    const svgIcon = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><path d="M47.0052 40.22993a6.78243 6.78243 0 00-6.77468 6.77517c.31793 8.967 13.23278 8.96474 13.54937-.00012A6.78241 6.78241 0 0047.0052 40.22993zm0 11.55113c-6.31094-.20109-6.30947-9.3518.00012-9.55192C53.31614 42.43018 53.31468 51.58094 47.0052 51.78106zM46.00559 35.00981v2.399a.99961.99961 0 101.99922 0v-2.399A.9998.9998 0 0046.00559 35.00981zM37.81643 37.81632a.99923.99923 0 000 1.41351l1.6966 1.6966a.9995.9995 0 101.4135-1.41351l-1.6966-1.6966A.99924.99924 0 0037.81643 37.81632zM38.409 47.0051a.99935.99935 0 00-.99961-.99961H35.00991a.99961.99961 0 000 1.99922h2.39945A.99935.99935 0 0038.409 47.0051zM39.513 53.08377l-1.6966 1.6966a.9995.9995 0 101.4135 1.41351l1.6966-1.6966A.99969.99969 0 0039.513 53.08377z"/><path d="M47.0052 55.60182a.99935.99935 0 00-.99961.99961v2.399a.99961.99961 0 001.99922 0v-2.399A.99935.99935 0 0047.0052 55.60182zM54.49738 53.08377a.99967.99967 0 00-1.41345 1.41351l1.69654 1.6966A.9995.9995 0 0056.194 54.78037zM59.00049 46.00549H56.60105a.9998.9998 0 00.00006 1.99922h2.39938A.9998.9998 0 0059.00049 46.00549zM53.79062 41.21929a.99638.99638 0 00.70676-.29286l1.6966-1.6966a.9995.9995 0 00-1.41351-1.41351l-1.6966 1.6966A1.00626 1.00626 0 0053.79062 41.21929zM35.7166 14.31084l-1.2924-1.2924h8.58257a4.0027 4.0027 0 013.99843 3.99843V28.01256a.99961.99961 0 001.99922 0V17.01687a6.0042 6.0042 0 00-5.99765-5.99764H34.4242l1.29246-1.29247A.99969.99969 0 0034.3031 8.31332c-4.38055 4.53606-4.38824 2.87094.00006 7.41109A.99968.99968 0 0035.7166 14.31084zM29.719 48.29756a.99966.99966 0 00-1.41345 1.41351l1.2924 1.29246H21.0154A4.0027 4.0027 0 0117.017 47.0051V36.00942a.99961.99961 0 10-1.99921 0V47.0051a6.0042 6.0042 0 005.99764 5.99765H29.598l-1.29246 1.29246a.9997.9997 0 001.41357 1.41344C34.09907 51.173 34.10774 52.83783 29.719 48.29756zM27.99606 23.98094a1.0056 1.0056 0 00-1.19685-1.4716 7.85738 7.85738 0 01-2.785.50518A7.99779 7.99779 0 0122.36448 7.19114a1.00566 1.00566 0 00.25473-1.905C14.302 1.10667 3.81073 7.66428 4.02226 17.01705 4.07412 29.95369 21.04255 34.87876 27.99606 23.98094zM6.02129 17.01687a11.04641 11.04641 0 0113.21265-10.776 9.99482 9.99482 0 005.35818 18.75625C17.74539 31.6113 5.95137 26.58923 6.02129 17.01687z"/></svg>'; // مخفف SVG

    toggleButton.innerHTML = svgIcon; // تغییر به ایموجی روز
    toggleButton.onclick = function() {
        if (document.body.classList.contains('day')) {
            document.body.classList.remove('day');
            document.body.classList.add('night');
            document.querySelector('.container').classList.remove('day');
            document.querySelector('.container').classList.add('night'); // تغییر حالت container به شب
            
            // تغییر حالت برای سایر کلاس‌ها
            document.querySelectorAll('.tablink').forEach(link => {
                link.classList.remove('day');
                link.classList.add('night');
            });
            // اضافه کردن تغییر حالت برای سایر المان‌ها
            document.querySelectorAll('.info-table td').forEach(td => {
                td.classList.remove('day');
                td.classList.add('night');
            });
            document.querySelectorAll('.search-container select, .search-container input').forEach(input => {
                input.classList.remove('day');
                input.classList.add('night');
            });
            // ... سایر کلاس‌ها را به همین صورت تغییر دهید ...

            toggleButton.innerHTML = svgIcon; // تغییر به ایموجی شب
        } else {
            document.body.classList.remove('night');
            document.body.classList.add('day');
            document.querySelector('.container').classList.remove('night');
            document.querySelector('.container').classList.add('day'); // تغییر حالت container به روز
            
            // تغییر حالت برای سایر کلاس‌ها
            document.querySelectorAll('.tablink').forEach(link => {
                link.classList.remove('night');
                link.classList.add('day');
            });
            // اضافه کردن تغییر حالت برای سایر المان‌ها
            document.querySelectorAll('.info-table td').forEach(td => {
                td.classList.remove('night');
                td.classList.add('day');
            });
            document.querySelectorAll('.search-container select, .search-container input').forEach(input => {
                input.classList.remove('night');
                input.classList.add('day');
            });
            // ... سایر کلاس‌ها را به همین صورت تغییر دهید ...
            
            toggleButton.innerHTML = svgIcon; // تغییر به ایموجی روز
        }
    };
    document.querySelector('.container').appendChild(toggleButton); // اضافه کردن دکمه به container

    // بررسی و خواندن مسیر فایل
    if (filePath) {
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
    }

    // دکمه برای بالا آمدن
    const scrollToTopButton = document.getElementById('scrollToTop');

    // نمایش یا پنهان کردن دکمه بر اساس اسکرول
    window.onscroll = function() {
        if (document.body.scrollTop > 100 || document.documentElement.scrollTop > 100) {
            scrollToTopButton.style.display = "block"; // نمایش دکمه
        } else {
            scrollToTopButton.style.display = "none"; // پنهان کردن دکمه
        }
    };

    // عملکرد دکمه برای بالا آمدن
    scrollToTopButton.onclick = function() {
        window.scrollTo({top: 0, behavior: 'smooth'}); // اسکرول به بالا با انیمیشن
    };
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

    rows.slice(1).forEach((row, index) => { // اضافه کردن index به عنوان شماره ردیف
        if (row[0]) {
            const tr = document.createElement('tr');
            const rowIndex = index + 1; // شماره ردیف
            const tdIndex = document.createElement('td');
            tdIndex.textContent = rowIndex; // نمایش شماره ردیف
            tr.appendChild(tdIndex); // اضافه کردن شماره ردیف به ردیف جدول

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

    rows.forEach((row, index) => {
        const cell = row.cells[searchColumn];
        const text = cell ? cell.innerText.toLowerCase() : "";
        const rowIndex = (index + 1).toString(); // دریافت شماره ردیف و تبدیل به رشته
        // جستجو در شماره ردیف و ستون انتخاب شده
        row.style.display = text.includes(searchInput) || rowIndex.includes(searchInput) ? '' : 'none';
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
    elmnt.style.backgroundColor = "#bbb";
}

document.getElementById("defaultOpen").click();

function toggleVisibility(id) {
    const element = document.getElementById(id);
    if (element) {
        element.style.display = (element.style.display === 'none' || element.style.display === '') ? 'inline' : 'none'; // تغییر وضعیت نمایش
    }
}
