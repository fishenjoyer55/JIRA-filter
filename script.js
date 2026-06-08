upload = document.getElementById("upload");
upload.addEventListener("change", guy => {
    let file = guy.target.files[0];
    let reader = new FileReader();

    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, {type: "array"});

        const sheetCount = workbook.SheetNames.length;
        const sheetSelector = document.getElementById("sheet-selector");
        const dropdownSelector = document.getElementById("dropdown-selector");
        const dropdownFilter = document.getElementById("dropdown-filter");
        let sheetSelectorHTML = "<select><option value='--'>--</option>";
        workbook.SheetNames.forEach(sheetName => {
            sheetSelectorHTML += "<option value='" + sheetName + "'>" + sheetName + "</option>";
        });
        sheetSelectorHTML += "</select>";
        sheetSelector.innerHTML = sheetSelectorHTML;

        let selectedSheet = workbook.SheetNames[1];
        let keyword = "";

        let dropdownSelectorHTML = "<select><option value='--'>--</option>";
        let rows = XLSX.utils.sheet_to_json(workbook.Sheets[selectedSheet], {header: 1});
        for (i of [0, 5, 7]) {
            dropdownSelectorHTML += "<option value='" + i + "'>" + rows[0][i] + "</option>";
        }
        dropdownSelectorHTML += "</select>";
        dropdownSelector.innerHTML = dropdownSelectorHTML;

        const giveFilterOptions = (col) => {
            let dropdownFilterHTML = "<select><option value='--'>--</option>";
            keywordList = [];
            rows.slice(1).forEach(row => {
                if (!keywordList.includes(row[col])) {
                    keywordList.push(row[col]);
                }
            });
            keywordList.forEach(entry => { 
                dropdownFilterHTML += "<option value='" + entry + "'>" + entry + "</option>";
            });
            dropdownFilterHTML += "</select>"; 
            dropdownFilter.innerHTML = dropdownFilterHTML;   
        }
        
        const checkCell = (cell) => {
            cellContents = cell !== undefined ? String(cell) : "something's wrong with your cell lmao";
            return cellContents.toLowerCase().includes(keyword);
        }
        const render = (specificColumn) => {
            const sheet = workbook.Sheets[selectedSheet];
            rows = XLSX.utils.sheet_to_json(sheet, {header: 1});
            let outputHTML = "<table border='1'>";
            writeRow = row => {
                    outputHTML += "<tr>";
                    row.forEach(cell => {
                        cellContents = cell !== undefined ? cell : "something's wrong with your cell lmao";
                        outputHTML += "<td>" + cellContents + "</td>";
                    });
                    outputHTML += "</tr>";  
            }
            const sliceStarter = selectedSheet == "Summary" ? 2 : 1;
            writeRow(rows[sliceStarter - 1]);
            rows.slice(sliceStarter).forEach(row => {
                let containsKeyword = false;
                if (specificColumn == -1) {
                    row.forEach(cell => {
                        if (checkCell(cell)) {
                            containsKeyword = true;
                        }
                    })
                } else {
                    containsKeyword = checkCell(row[specificColumn]);
                }
                if (containsKeyword) {
                    writeRow(row);
                }
            });
            outputHTML += "</table>" + "<p> end of table </p>";

            document.getElementById("output").innerHTML = outputHTML;
        }
        
        sheetSelector.addEventListener("change", event => {
            selectedSheet = event.target.value;
            render(-1);
        });

        document.getElementById("input-filter").addEventListener("keydown", event => {
            if (event.key === "Enter") {
                keyword = event.target.value.toLowerCase();
                render(-1);
            }
        });

        dropdownFilter.addEventListener("change", event => {
            keyword = event.target.value.toLowerCase();
            render(specificColumn);
        });

        dropdownSelector.addEventListener("change", event => {
            giveFilterOptions(parseInt(event.target.value));
            specificColumn = parseInt(event.target.value);
        });

        render(-1);
    }

    reader.readAsArrayBuffer(file);
});