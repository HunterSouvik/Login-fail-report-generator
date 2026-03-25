document.addEventListener('DOMContentLoaded', () => {
    // Current State
    const state = {
        failData: null,
        failHeaders: [],
        successData: null,
        successHeaders: [],
        templateBuffer: null, // Store raw original file buffer for ExcelJS
        templateFailRowIdx: -1,
        templateSuccRowIdx: -1,
        templateFailHeaders: [],
        templateSuccHeaders: []
    };

    // DOM Elements - Stepper
    const steps = [
        document.getElementById('step1Indicator'),
        document.getElementById('step2Indicator'),
        document.getElementById('step3Indicator'),
        document.getElementById('step4Indicator'),
        document.getElementById('step5Indicator')
    ];
    const panels = [
        document.getElementById('step1'),
        document.getElementById('step2'),
        document.getElementById('step3'),
        document.getElementById('step4'),
        document.getElementById('step5')
    ];

    // -- Live Clock Logic --
    function updateClock() {
        const dateEl = document.getElementById('clockDate');
        const timeEl = document.getElementById('clockTime');
        if (!dateEl || !timeEl) return;

        const now = new Date();
        
        // Date: DD MMM YYYY
        const day = String(now.getDate()).padStart(2, '0');
        const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
        const month = months[now.getMonth()];
        const year = now.getFullYear();
        dateEl.textContent = `${day} ${month} ${year}`;

        // Time: HH:mm:ss
        const hours = String(now.getHours()).padStart(2, '0');
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const seconds = String(now.getSeconds()).padStart(2, '0');
        timeEl.textContent = `${hours}:${minutes}:${seconds}`;
    }

    updateClock();
    setInterval(updateClock, 1000);

    // -- Theme Toggle Logic --
    const themeToggle = document.getElementById('themeToggle');
    const sunIcon = themeToggle.querySelector('.sun-icon');
    const moonIcon = themeToggle.querySelector('.moon-icon');

    function setTheme(isLight) {
        if (isLight) {
            document.body.classList.add('light-mode');
            sunIcon.style.display = 'none';
            moonIcon.style.display = 'block';
        } else {
            document.body.classList.remove('light-mode');
            sunIcon.style.display = 'block';
            moonIcon.style.display = 'none';
        }
        localStorage.setItem('wifi_report_theme', isLight ? 'light' : 'dark');
    }

    // Initialize theme
    const savedTheme = localStorage.getItem('wifi_report_theme');
    if (savedTheme === 'light') {
        setTheme(true);
    }

    themeToggle.addEventListener('click', () => {
        const isLight = document.body.classList.contains('light-mode');
        setTheme(!isLight);
    });

    // -- Keyboard Shortcuts --
    document.addEventListener('keydown', (e) => {
        // Alt + T: Toggle Theme
        if (e.altKey && e.key.toLowerCase() === 't') {
            const isLight = document.body.classList.contains('light-mode');
            setTheme(!isLight);
        }
        // Esc: Close Modal
        if (e.key === 'Escape') {
            const guideModal = document.getElementById('guideModal');
            if (guideModal && guideModal.style.display === 'flex') {
                toggleModal(false);
            }
        }
    });

    // Navigation Buttons
    const btnNext1 = document.getElementById('nextToStep2');
    const btnBack1 = document.getElementById('backToStep1');
    const btnNext2 = document.getElementById('nextToStep3');
    const btnBack2 = document.getElementById('backToStep2');
    const btnNext3 = document.getElementById('nextToStep4');
    const btnBack3 = document.getElementById('backToStep3');
    const btnGenerate = document.getElementById('generateExcelBtn');
    const successFile = document.getElementById('successFile');
    const templateFile = document.getElementById('templateFile');
    const btnDownloadTemplate = document.getElementById('downloadTemplate');

    // Navigation Functions
    function goToStep(index) {
        panels.forEach((p, i) => {
            p.style.display = i === index ? 'block' : 'none';
        });
        steps.forEach((s, i) => {
            s.classList.remove('active');
            if (i < index) s.classList.add('completed');
            if (i === index) s.classList.add('active');
            if (i > index) s.classList.remove('completed');
        });
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    btnNext1.addEventListener('click', () => goToStep(1));
    btnBack1.addEventListener('click', () => goToStep(0));
    btnNext2.addEventListener('click', () => goToStep(2));
    btnBack2.addEventListener('click', () => goToStep(1));
    btnNext3.addEventListener('click', () => {
        buildMappingUI();
        goToStep(3);
    });
    btnBack3.addEventListener('click', () => goToStep(2));

    // Drag and Drop Logic Helper
    function setupDropArea(dropArea, fileInput, onFileDrop) {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, preventDefaults, false);
        });
        function preventDefaults(e) { e.preventDefault(); e.stopPropagation(); }
        ['dragenter', 'dragover'].forEach(eventName => {
            dropArea.addEventListener(eventName, () => dropArea.classList.add('dragover'), false);
        });
        ['dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, () => dropArea.classList.remove('dragover'), false);
        });
        dropArea.addEventListener('drop', e => {
            const dt = e.dataTransfer;
            if (dt.files.length) { fileInput.files = dt.files; onFileDrop(dt.files[0]); }
        });
        dropArea.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', function() {
            if (this.files.length) onFileDrop(this.files[0]);
        });
    }

    // Helper: Read Data File using SheetJS for robust CSV/Excel raw data extraction
    function readDataFile(file, onSuccess) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                
                // Get raw rows to find the headers row
                const rawRows = XLSX.utils.sheet_to_json(firstSheet, {header: 1, defval: ""});
                if (rawRows.length === 0) throw new Error("File is empty.");
                
                // Heuristic: Find first row with at least 3 unique text values (skips titles/empty lines)
                let headerIdx = 0;
                for (let i = 0; i < rawRows.length; i++) {
                    const row = rawRows[i];
                    const uniqueVals = new Set(row.filter(v => v !== "" && isNaN(v)));
                    
                    const hasTitle = Array.from(uniqueVals).some(v => String(v).toUpperCase().includes("LOGIN-"));
                    if (uniqueVals.size >= 3 && !hasTitle) {
                        headerIdx = i;
                        break;
                    }
                }
                
                const headers = rawRows[headerIdx]; 
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, {range: headerIdx, defval: ""});
                
                onSuccess(jsonData, headers);
            } catch (err) {
                alert("Error reading Data file: " + err.message);
                console.error(err);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // 1. Process Fail Data File
    const failDropArea = document.getElementById('failDropArea');
    const failFileInfoBox = document.getElementById('failFileInfoBox');
    setupDropArea(failDropArea, document.getElementById('failFile'), file => {
        readDataFile(file, (jsonData, headers) => {
            state.failData = jsonData;
            state.failHeaders = headers;
            
            // Proactively build selectors
            populateSelector(document.getElementById('failMacSelect'), headers, "mac");
            
            failDropArea.style.display = 'none';
            failFileInfoBox.style.display = 'flex';
            document.getElementById('failFileName').textContent = file.name;
            document.getElementById('failFileMeta').textContent = `${jsonData.length} records`;
            btnNext1.disabled = false;
        });
    });

    document.getElementById('removeFailFile').addEventListener('click', () => {
        state.failData = null;
        state.failHeaders = [];
        document.getElementById('failFile').value = '';
        failDropArea.style.display = 'block';
        failFileInfoBox.style.display = 'none';
        btnNext1.disabled = true;
    });

    // 2. Process Success Data File
    const successDropArea = document.getElementById('successDropArea');
    const successFileInfoBox = document.getElementById('successFileInfoBox');
    setupDropArea(successDropArea, document.getElementById('successFile'), file => {
        readDataFile(file, (jsonData, headers) => {
            state.successData = jsonData;
            state.successHeaders = headers;

            // Proactively build selectors
            populateSelector(document.getElementById('successMacSelect'), headers, "mac");

            successDropArea.style.display = 'none';
            successFileInfoBox.style.display = 'flex';
            document.getElementById('successFileName').textContent = file.name;
            document.getElementById('successFileMeta').textContent = `${jsonData.length} records`;
            btnNext2.disabled = false;
        });
    });

    document.getElementById('removeSuccessFile').addEventListener('click', () => {
        state.successData = null;
        state.successHeaders = [];
        document.getElementById('successFile').value = '';
        successDropArea.style.display = 'block';
        successFileInfoBox.style.display = 'none';
        btnNext2.disabled = true;
    });

    // 3. Process Template File (using ExcelJS to preserve styles)
    const templateDropArea = document.getElementById('templateDropArea');
    const templateFileInfoBox = document.getElementById('templateFileInfoBox');
    setupDropArea(templateDropArea, document.getElementById('templateFile'), file => {
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                const buffer = e.target.result;
                state.templateBuffer = buffer; // save for writing later
                
                // Parse with ExcelJS to find headers
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(buffer);
                
                let headersDetected = [];

                // Scan all worksheets in case the template is not on the first one
                workbook.eachSheet(worksheet => {
                    if (headersDetected.length >= 2) return; // Found enough headers
                    
                    worksheet.eachRow(function(row, rowNumber) {
                        let textCols = [];
                        let uniqueVals = new Set();
                        
                        row.eachCell({ includeEmpty: false }, function(cell, colNumber) {
                            let val = "";
                            // ExcelJS cell value can be string, number, date, richText, formula, etc.
                            if (typeof cell.value === 'string') {
                                val = cell.value.trim();
                            } else if (cell.value && cell.value.richText) {
                                // Handle stylized/rich text headers
                                val = cell.value.richText.map(rt => rt.text).join('').trim();
                            } else if (cell.value && typeof cell.value === 'object' && cell.value.result) {
                                // Handle formula results
                                val = String(cell.value.result).trim();
                            }
                            
                            // Headers are typically non-empty strings and not just numbers
                            if (val !== "" && isNaN(val)) {
                                textCols.push({ col: colNumber, val: val });
                                uniqueVals.add(val);
                            }
                        });
                        
                        // Heuristic: A header row should have several unique text labels.
                        // EXCLUDE Title Rows: Titles like "LOGIN-Failure" usually have 1 unique value or contain the word "LOGIN-" with very few other columns.
                        const hasTitle = Array.from(uniqueVals).some(v => v.toUpperCase().includes("LOGIN-"));
                        if (uniqueVals.size >= 4 && !hasTitle) {
                            headersDetected.push({
                                worksheet: worksheet,
                                rowIdx: rowNumber,
                                headers: textCols
                            });
                        }
                    });
                });

                if (headersDetected.length < 2) {
                    throw new Error("Could not find the two distinct header sections (Failure and Success) in the template correctly. Please ensure your template has unique column names (ID, Date, etc.) for both tables.");
                }

                // The first substantive row is Fail Headers, the second is Success Headers
                state.templateFailRowIdx = headersDetected[0].rowIdx;
                state.templateFailHeaders = headersDetected[0].headers;
                state.activeWorksheetName = headersDetected[0].worksheet.name; // Remember which sheet has the data
                
                state.templateSuccRowIdx = headersDetected[1].rowIdx;
                state.templateSuccHeaders = headersDetected[1].headers;

                templateDropArea.style.display = 'none';
                templateFileInfoBox.style.display = 'flex';
                document.getElementById('templateFileName').textContent = file.name;
                document.getElementById('templateFileMeta').textContent = `Tables detected at rows ${state.templateFailRowIdx} & ${state.templateSuccRowIdx}`;
                btnNext3.disabled = false;

            } catch (err) {
                alert("Error reading Template: " + err.message);
                console.error(err);
            }
        };
        reader.readAsArrayBuffer(file);
    });

    document.getElementById('removeTemplateFile').addEventListener('click', () => {
        state.templateBuffer = null;
        state.templateFailHeaders = [];
        state.templateSuccHeaders = [];
        document.getElementById('templateFile').value = '';
        templateDropArea.style.display = 'block';
        templateFileInfoBox.style.display = 'none';
        btnNext3.disabled = true;
    });

    // 4. Build the MAC selectors and Column Mapping Interface
    const failMacSelect = document.getElementById('failMacSelect');
    const successMacSelect = document.getElementById('successMacSelect');
    const failMappingBody = document.getElementById('failMappingBody');
    const successMappingBody = document.getElementById('successMappingBody');

    function findBestMatch(targetColName, sourceHeadersArray) {
        const t = String(targetColName).toLowerCase().trim();
        if (t.includes("login-") || t === "") return ""; 

        // 1. Exact or strict include match (Highest Priority)
        for (const sOriginal of sourceHeadersArray) {
            if (!sOriginal) continue;
            const s = String(sOriginal).toLowerCase().trim();
            if (s === t) return sOriginal;
        }

        // 2. Contains match (Medium Priority)
        for (const sOriginal of sourceHeadersArray) {
            if (!sOriginal) continue;
            const s = String(sOriginal).toLowerCase().trim();
            if (s.includes(t) || t.includes(s)) return sOriginal;
        }

        // 3. Synonym match (Fallback Priority)
        const synonyms = {
            "mac": ["mac", "mac address", "macid", "device mac", "terminal mac", "mac add", "dev mac"],
            "date": ["date", "time", "datetime", "login time", "failure time", "event time", "timestamp", "failuredate"],
            "user": ["username", "account", "user", "user id", "mobile", "phone", "terminal id", "account id"],
            "reason": ["reason", "fail reason", "error", "message", "remark", "result", "failreason", "failmsg", "status", "fault", "outcome", "error description", "fail status"],
            "nas": ["nas", "nas name", "nas ip", "ap name", "ap mac"],
            "room": ["room", "room no", "room number", "apartment", "unit"],
            "lastname": ["last name", "lastname", "surname", "family name"]
        };

        for (const [key, list] of Object.entries(synonyms)) {
            const isTargetInStock = list.some(syn => t.includes(syn) || syn.includes(t));
            if (isTargetInStock) {
                for (const sOriginal of sourceHeadersArray) {
                    if (!sOriginal) continue;
                    const s = String(sOriginal).toLowerCase().trim();
                    if (list.some(syn => s.includes(syn) || syn.includes(s))) {
                        return sOriginal;
                    }
                }
            }
        }
        return "";
    }

    function populateSelector(selectEl, optionsArray, guessMatchName) {
        if (!selectEl) return;
        const currentVal = selectEl.value;
        
        selectEl.innerHTML = '<option value="">-- Choose Column --</option>';
        const bestMatch = findBestMatch(guessMatchName, optionsArray);

        optionsArray.forEach(optVal => {
            if (optVal === undefined || optVal === null || optVal === "") return;
            const opt = document.createElement('option');
            const valStr = String(optVal);
            opt.value = valStr;
            opt.textContent = valStr;
            
            // Auto-match if it's the best match found
            if (valStr === bestMatch) {
                opt.selected = true;
            }
            // If it matches what the user already picked, definitely re-select it
            if (valStr === currentVal) {
                opt.selected = true;
            }
            selectEl.appendChild(opt);
        });
    }

    function buildMappingList(container, templateHeadersObj, sourceHeadersArray, savedMappings) {
        container.innerHTML = '';
        templateHeadersObj.forEach((th) => {
            const targetColName = th.val;
            const targetColIndex = th.col;

            const row = document.createElement('div');
            row.className = 'mapping-row';

            const colNameDiv = document.createElement('div');
            colNameDiv.className = 'template-col-name';
            colNameDiv.textContent = targetColName;

            const selectDiv = document.createElement('div');
            const select = document.createElement('select');
            select.dataset.targetColIdx = targetColIndex; 
            select.dataset.targetColName = targetColName;

            const defaultOpt = document.createElement('option');
            defaultOpt.value = "";
            defaultOpt.textContent = "-- Skip / Leave Blank --";
            select.appendChild(defaultOpt);

            // Priority: 1. Saved Mapping, 2. Best Match
            const savedVal = savedMappings ? savedMappings[targetColName] : null;
            const bestMatch = findBestMatch(targetColName, sourceHeadersArray);

            sourceHeadersArray.forEach(sourceCol => {
                const opt = document.createElement('option');
                opt.value = sourceCol;
                opt.textContent = sourceCol;
                
                if (savedVal === sourceCol) {
                    opt.selected = true;
                } else if (!savedVal && sourceCol === bestMatch) {
                    opt.selected = true;
                }
                select.appendChild(opt);
            });

            selectDiv.appendChild(select);
            row.appendChild(colNameDiv);
            row.appendChild(selectDiv);
            container.appendChild(row);
        });
    }

    function updateMatchPreview() {
        const failMacKey = failMacSelect.value;
        const successMacKey = successMacSelect.value;
        const badge = document.getElementById('matchCounter');
        const countText = document.getElementById('matchCountText');

        if (!failMacKey || !successMacKey || !state.failData || !state.successData) {
            badge.style.display = 'none';
            return;
        }

        const normalizeMac = (macStr) => String(macStr).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
        const failMacSet = new Set();
        state.failData.forEach(row => {
            const mac = row[failMacKey];
            if (mac) failMacSet.add(normalizeMac(mac));
        });

        let matchCount = 0;
        state.successData.forEach(row => {
            const mac = row[successMacKey];
            if (mac) {
                if (failMacSet.has(normalizeMac(mac))) matchCount++;
            }
        });

        countText.textContent = `${matchCount} matches found`;
        badge.style.display = 'flex';
    }

    function saveMappings() {
        const mappings = {
            failMac: failMacSelect.value,
            successMac: successMacSelect.value,
            failFields: {},
            successFields: {}
        };

        failMappingBody.querySelectorAll('select').forEach(s => {
            if (s.value) mappings.failFields[s.dataset.targetColName] = s.value;
        });
        successMappingBody.querySelectorAll('select').forEach(s => {
            if (s.value) mappings.successFields[s.dataset.targetColName] = s.value;
        });

        localStorage.setItem('wifi_report_mappings', JSON.stringify(mappings));
    }

    function buildMappingUI() {
        const saved = JSON.parse(localStorage.getItem('wifi_report_mappings') || '{}');
        
        populateSelector(failMacSelect, state.failHeaders, "mac");
        if (saved.failMac && state.failHeaders.includes(saved.failMac)) {
            failMacSelect.value = saved.failMac;
        }

        populateSelector(successMacSelect, state.successHeaders, "mac");
        if (saved.successMac && state.successHeaders.includes(saved.successMac)) {
            successMacSelect.value = saved.successMac;
        }

        buildMappingList(failMappingBody, state.templateFailHeaders, state.failHeaders, saved.failFields);
        buildMappingList(successMappingBody, state.templateSuccHeaders, state.successHeaders, saved.successFields);
        
        updateMatchPreview();
    }

    failMacSelect.addEventListener('change', updateMatchPreview);
    successMacSelect.addEventListener('change', updateMatchPreview);

    // 5. Generate Output Excel
    btnGenerate.addEventListener('click', async () => {
        // Validation
        const failMacKey = failMacSelect.value;
        const successMacKey = successMacSelect.value;

        if (!failMacKey || !successMacKey) {
            alert("Please select the MAC Address column for both reports so we can match them.");
            return;
        }

        saveMappings();

        btnGenerate.disabled = true;
        const originalText = btnGenerate.innerHTML;
        btnGenerate.innerHTML = 'Intersecting Data...';

        try {
            // Helper to clean MAC addresses (strips all non-alphanumeric chars for robust matching)
            const normalizeMac = (macStr) => {
                return String(macStr).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
            };

            // -- Data Processing: Fail-First Priority --
            // 1. Collect all MACs from the Fail Report for matching
            const failMacSet = new Set();
            state.failData.forEach(row => {
                const mac = row[failMacKey];
                if (mac) failMacSet.add(normalizeMac(mac));
            });

            // 2. All Failures are included in the report
            const filteredFailData = state.failData;
            
            // 3. Successes are only included if the MAC exists in the Fail report
            const matchedSuccessDataFlat = [];
            const macToSuccMap = new Map(); // Store first success record for each MAC for cross-reference

            state.successData.forEach(row => {
                const mac = row[successMacKey];
                if (mac) {
                    const cleanMac = normalizeMac(mac);
                    if (failMacSet.has(cleanMac)) {
                        matchedSuccessDataFlat.push(row);
                        if (!macToSuccMap.has(cleanMac)) {
                            macToSuccMap.set(cleanMac, row);
                        }
                    }
                }
            });

            console.log(`Yielded ${filteredFailData.length} total Fail rows and ${matchedSuccessDataFlat.length} matching Success rows.`);

            if (filteredFailData.length === 0) {
                alert("No matching MAC addresses were found between the two reports.");
                btnGenerate.innerHTML = originalText;
                btnGenerate.disabled = false;
                return;
            }

            btnGenerate.innerHTML = 'Writing to Template...';

            // -- Excel Injecting --
            // Load template workbook with ExcelJS
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(state.templateBuffer);
            // Use the worksheet where we originally found the headers
            const worksheet = workbook.getWorksheet(state.activeWorksheetName) || workbook.worksheets[0];

            // Get mapping dicts { targetColIdx : sourceColName }
            function extractMapping(bodyEl) {
                const selects = bodyEl.querySelectorAll('select');
                const mapping = {};
                selects.forEach(sel => {
                    if (sel.value) mapping[sel.dataset.targetColIdx] = sel.value;
                });
                return mapping;
            }

            // Helper to clean up long error messages into concise reasons
            function smartReasonClean(val, srcRow, matchSuccRow, failMap, succMap) {
                // If we have a success record, prioritize comparing data
                if (matchSuccRow) {
                    // Find mapping names for Room and Last Name
                    // We look at the template headers to know which target column is which
                    const rHdr = state.templateFailHeaders.find(h => h.val.toLowerCase().includes("room"));
                    const nHdr = state.templateFailHeaders.find(h => h.val.toLowerCase().includes("last name") || h.val.toLowerCase().includes("surname"));
                    
                    const rSuccHdr = state.templateSuccHeaders.find(h => h.val.toLowerCase().includes("room"));
                    const nSuccHdr = state.templateSuccHeaders.find(h => h.val.toLowerCase().includes("last name") || h.val.toLowerCase().includes("surname"));

                    if (rHdr && rSuccHdr && failMap[rHdr.col] && succMap[rSuccHdr.col]) {
                        const failRoom = String(srcRow[failMap[rHdr.col]] || "").trim();
                        const succRoom = String(matchSuccRow[succMap[rSuccHdr.col]] || "").trim();
                        if (failRoom && succRoom && failRoom !== succRoom) {
                            return "Wrong Room Number";
                        }
                    }

                    if (nHdr && nSuccHdr && failMap[nHdr.col] && succMap[nSuccHdr.col]) {
                        const failName = String(srcRow[failMap[nHdr.col]] || "").trim().toLowerCase();
                        const succName = String(matchSuccRow[succMap[nSuccHdr.col]] || "").trim().toLowerCase();
                        if (failName && succName && failName !== succName) {
                            return "Wrong Last Name";
                        }
                    }
                }

                if (!val) return "";
                const s = String(val).toLowerCase();
                
                // Priority Check: Room Number
                if (s.includes("room") && (s.includes("not match") || s.includes("wrong") || s.includes("invalid") || s.includes("not found"))) {
                    return "Wrong Room Number";
                }
                // Last Name Check
                if (s.includes("last name") || s.includes("surname")) {
                    if (s.includes("not match") || s.includes("wrong") || s.includes("invalid") || s.includes("fail")) {
                        return "Wrong Last Name";
                    }
                }
                // General Account Issues
                if (s.includes("not found") || s.includes("no such user") || s.includes("doesn't exist")) {
                    return "Invalid Account/User";
                }
                if (s.includes("password") || s.includes("credential")) {
                    return "Incorrect Password";
                }
                if (s.includes("timeout") || s.includes("time out") || s.includes("timed out")) {
                    return "Connection Timeout";
                }
                if (s.includes("max") || s.includes("limit") || s.includes("concurrent")) {
                    return "Device Limit Reached";
                }
                
                return val; // Return original if no match
            }

            const failMap = extractMapping(failMappingBody);
            const succMap = extractMapping(successMappingBody);

            // Function to duplicate specific row style for newly inserted rows
            function applyStyleFromTo(sourceRow, targetRow) {
                sourceRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    const newCell = targetRow.getCell(colNumber);
                    newCell.style = Object.assign({}, cell.style);
                    newCell.value = null; // NEVER copy values when styling
                });
            }

            // We must inject Bottom-Up. If we inject Fail rows first, it pushes the Success Header down, invalidating our saved row index!
            
            // 1. Inject Success Data (Bottom Table)
            let succInsertIdx = state.templateSuccRowIdx + 1;
            const succStyleRow = worksheet.getRow(succInsertIdx); 
            
            for (let i = 0; i < matchedSuccessDataFlat.length; i++) {
                const srcRow = matchedSuccessDataFlat[i];
                let newRow;
                
                if (i === 0) {
                    newRow = worksheet.getRow(succInsertIdx);
                    // Clear cell values to prevent placeholder repetition
                    newRow.eachCell({ includeEmpty: true }, cell => cell.value = null);
                } else {
                    worksheet.spliceRows(succInsertIdx + i, 0, []);
                    newRow = worksheet.getRow(succInsertIdx + i);
                    applyStyleFromTo(succStyleRow, newRow);
                }
                
                for (const [colIdx, srcColName] of Object.entries(succMap)) {
                    if (srcRow[srcColName] !== undefined) {
                        const cell = newRow.getCell(Number(colIdx));
                        
                        let cellValue = srcRow[srcColName];
                        // If the column is Status/Reason/Result and it's a success, set it to "Login Success"
                        const targetColName = successMappingBody.querySelector(`select[data-target-col-idx="${colIdx}"]`).dataset.targetColName.toLowerCase();
                        if ((targetColName.includes("reason") || targetColName.includes("status") || targetColName.includes("result")) && (!cellValue || String(cellValue).trim() === "")) {
                            cellValue = "Login Success";
                        }

                        cell.value = cellValue;
                    }
                }
                newRow.commit();
            }

            // 2. Inject Fail Data (Top Table)
            let failInsertIdx = state.templateFailRowIdx + 1;
            const failStyleRow = worksheet.getRow(failInsertIdx); 
            
            for (let i = 0; i < filteredFailData.length; i++) {
                const srcRow = filteredFailData[i];
                let newRow;

                if (i === 0) {
                    newRow = worksheet.getRow(failInsertIdx);
                    // Clear cell values to prevent placeholder repetition
                    newRow.eachCell({ includeEmpty: true }, cell => cell.value = null);
                } else {
                    worksheet.spliceRows(failInsertIdx + i, 0, []);
                    newRow = worksheet.getRow(failInsertIdx + i);
                    applyStyleFromTo(failStyleRow, newRow);
                }

                for (const [colIdx, srcColName] of Object.entries(failMap)) {
                    if (srcRow[srcColName] !== undefined) {
                        const cell = newRow.getCell(Number(colIdx));
                        
                        let cellValue = srcRow[srcColName];
                        // Apply smart cleaning ONLY for columns that look like 'Reason'
                        const targetColName = failMappingBody.querySelector(`select[data-target-col-idx="${colIdx}"]`).dataset.targetColName.toLowerCase();
                        if (targetColName.includes("reason") || targetColName.includes("status")) {
                            const macInRow = srcRow[failMacKey];
                            const matchSuccRow = macInRow ? macToSuccMap.get(normalizeMac(macInRow)) : null;
                            cellValue = smartReasonClean(cellValue, srcRow, matchSuccRow, failMap, succMap);
                        }

                        cell.value = cellValue;
                    }
                }
                newRow.commit();
            }

            // Auto-fit Columns and Center Headers
            // Note: Column widths and alignments are now preserved from the template.

            // Clean up: Note - if the template had several empty styled rows, they are pushed down. 
            // In a production environment, we might find and delete them, but preserving them is safer to not break the user's custom footers or boxes.

            // Save File
            const outBuffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            
            const now = new Date();
            const dateStr = [
                String(now.getDate()).padStart(2, '0'),
                String(now.getMonth() + 1).padStart(2, '0'),
                now.getFullYear()
            ].join('.');
            const timeStr = [
                String(now.getHours()).padStart(2, '0'),
                String(now.getMinutes()).padStart(2, '0'),
                String(now.getSeconds()).padStart(2, '0')
            ].join('-');
            
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Login_Faild_Report_${dateStr}_${timeStr}.xlsx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            // Success feedback
            btnGenerate.innerHTML = '✓ Download Complete';
            
            // Show Statistics Dashboard
            document.getElementById('statTotalFail').textContent = filteredFailData.length;
            document.getElementById('statMatchedSucc').textContent = matchedSuccessDataFlat.length;
            
            const efficiency = filteredFailData.length > 0 
                ? Math.round((matchedSuccessDataFlat.length / filteredFailData.length) * 100) 
                : 0;
            document.getElementById('statEfficiency').textContent = `${efficiency}%`;

            setTimeout(() => {
                btnGenerate.innerHTML = originalText;
                btnGenerate.disabled = false;
                goToStep(4); // Move to Results Step
            }, 1000);

        } catch (err) {
            console.error(err);
            alert("An error occurred while generating the document: " + err.message);
            btnGenerate.innerHTML = originalText;
            btnGenerate.disabled = false;
        }
    });

    // -- Guide Modal Logic --
    const guideModal = document.getElementById('guideModal');
    const btnOpenGuide = document.getElementById('openGuide');
    const btnCloseGuide = document.getElementById('closeGuide');
    const btnCloseGuideBtn = document.getElementById('closeGuideBtn');

    const toggleModal = (show) => {
        guideModal.style.display = show ? 'flex' : 'none';
        document.body.style.overflow = show ? 'hidden' : 'auto';
    };

    btnOpenGuide.addEventListener('click', () => toggleModal(true));
    btnCloseGuide.addEventListener('click', () => toggleModal(false));
    btnCloseGuideBtn.addEventListener('click', () => toggleModal(false));

    // Close modal on outside click
    guideModal.addEventListener('click', (e) => {
        if (e.target === guideModal) toggleModal(false);
    });

    // -- Template Download Logic --
    if (btnDownloadTemplate) {
        btnDownloadTemplate.addEventListener('click', () => {
            const fileName = 'assets/LOGINN FAIL TEMPLATE.xlsx';
            const a = document.createElement('a');
            a.href = fileName;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        });
    }

    // -- Results Dashboard Actions --
    const btnProcessAnother = document.getElementById('processAnotherBtn');
    if (btnProcessAnother) {
        btnProcessAnother.addEventListener('click', () => {
            // Reset State
            state.failData = null;
            state.failHeaders = [];
            state.successData = null;
            state.successHeaders = [];
            state.templateBuffer = null;
            state.templateFailRowIdx = -1;
            state.templateSuccRowIdx = -1;
            state.templateFailHeaders = [];
            state.templateSuccHeaders = [];

            // Reset File Inputs
            document.getElementById('failFile').value = '';
            document.getElementById('successFile').value = '';
            document.getElementById('templateFile').value = '';

            // Reset File Info Boxes
            document.getElementById('failFileInfoBox').style.display = 'none';
            document.getElementById('successFileInfoBox').style.display = 'none';
            document.getElementById('templateFileInfoBox').style.display = 'none';

            // Reset File Details
            document.getElementById('failFileName').textContent = '';
            document.getElementById('failFileMeta').textContent = '0 rows detected';
            document.getElementById('successFileName').textContent = '';
            document.getElementById('successFileMeta').textContent = '0 rows detected';
            document.getElementById('templateFileName').textContent = '';
            document.getElementById('templateFileMeta').textContent = '0 columns detected';

            // Reset Drop Areas (remove success styling)
            document.getElementById('failDropArea').style.display = 'block';
            document.getElementById('successDropArea').style.display = 'block';
            document.getElementById('templateDropArea').style.display = 'block';

            // Reset Buttons
            document.getElementById('nextToStep2').disabled = true;
            document.getElementById('nextToStep3').disabled = true;
            document.getElementById('nextToStep4').disabled = true;

            // Reset Indicators
            steps.forEach(s => s.classList.remove('completed', 'active'));

            // Go back to Step 1
            goToStep(0);
        });
    }

});
