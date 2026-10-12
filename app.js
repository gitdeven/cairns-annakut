let selectedItemIds = [];

let allItems = [];

let selectedCategories = [];

let selectedTypes = [];

const DATA_FILE = "./cairns-items.csv";
const CENTRE_NAME = "Cairns";
const BACKEND_URL = "https://script.google.com/macros/s/AKfycbyjQcSrcWFu0HhLXJoWPxRMnWRDWMZ7f3aXEfFOoIk0zoMKiYlD6GJa1FhmpvXQ1ZVeZQ/exec";

function normalizeBoolean(value) {
    if (typeof value === "string") {
        return value.trim().toLowerCase() === "true";
    }

    return Boolean(value);
}

function buildExcelRows(name, phone, email, selectedItems, additional) {
    const rows = [];

    selectedItems.forEach(itemId => {
        const item = allItems.find(x => x.id === itemId);

        if (!item) {
            return;
        }

        rows.push({
            centre: CENTRE_NAME,
            full_name: name,
            phone: phone,
            email: email,
            item_name: item.item_name,
            category: item.category,
            item_type: item.item_type,
            qty_per_thaal: item.qty_per_thaal,
            submitted_at: new Date().toISOString(),
            item_status: "Allocated"
        });
    });

    if (additional) {
        rows.push({
            centre: CENTRE_NAME,
            full_name: name,
            phone: phone,
            email: email,
            item_name: additional,
            category: "Additional Item",
            item_type: "Additional Item",
            qty_per_thaal: "",
            submitted_at: new Date().toISOString(),
            item_status: "Additional"
        });
    }

    return rows;
}

function downloadFile(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

function exportToExcel(rows) {
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Allocation");
    XLSX.writeFile(workbook, `${CENTRE_NAME.toLowerCase()}-allocation-${Date.now()}.xlsx`);
}

function exportUpdatedCsv() {
    const csvRows = allItems.map(item => ({
        item_name: item.item_name,
        category: item.category,
        item_type: item.item_type,
        qty_per_thaal: item.qty_per_thaal,
        available: item.available ? "TRUE" : "FALSE"
    }));

    const header = ["item_name", "category", "item_type", "qty_per_thaal", "available"];
    const csvLines = [header.join(",")];

    csvRows.forEach(row => {
        const values = header.map(key => {
            const value = row[key] ?? "";
            const escaped = String(value).replace(/"/g, '""');
            return `"${escaped}"`;
        });

        csvLines.push(values.join(","));
    });

    const blob = new Blob([csvLines.join("\n")], { type: "text/csv;charset=utf-8;" });
    downloadFile(blob, `${CENTRE_NAME.toLowerCase()}-items-updated.csv`);
}

function markSelectedItemsUnavailable(selectedItems) {
    selectedItems.forEach(itemId => {
        const item = allItems.find(x => x.id === itemId);

        if (item) {
            item.available = false;
        }
    });
}

/* ==============================
   LOAD ITEMS
============================== */

async function loadItems() {
    try {
        let response;

        if (BACKEND_URL && BACKEND_URL !== "PASTE_YOUR_GOOGLE_APPS_SCRIPT_URL_HERE") {
            response = await fetch(`${BACKEND_URL}?action=getItems&centre=${encodeURIComponent(CENTRE_NAME)}`);
        } else {
            response = await fetch(DATA_FILE);
        }

        if (!response.ok) {
            throw new Error(`Unable to load the item data.`);
        }

        const responseText = await response.text();

        if (responseText.trim().startsWith("{")) {
            const payload = JSON.parse(responseText);
            const rows = payload.items || [];

            allItems = rows
                .map(row => ({
                    id: String(row.id || `${row.item_name}-${row.category}-${row.item_type}`),
                    item_name: String(row.item_name || "").trim(),
                    category: String(row.category || "").trim(),
                    item_type: String(row.item_type || "").trim(),
                    qty_per_thaal: String(row.qty_per_thaal || "").trim(),
                    available: typeof row.available === "undefined" ? true : normalizeBoolean(row.available)
                }))
                .filter(item => item.item_name && item.available !== false);
        } else {
            const workbook = XLSX.read(responseText, { type: "string" });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

            allItems = rows
                .map(row => ({
                    id: String(row.id || `${row.item_name}-${row.category}-${row.item_type}`),
                    item_name: String(row.item_name || "").trim(),
                    category: String(row.category || "").trim(),
                    item_type: String(row.item_type || "").trim(),
                    qty_per_thaal: String(row.qty_per_thaal || "").trim(),
                    available: typeof row.available === "undefined" ? true : normalizeBoolean(row.available)
                }))
                .filter(item => item.item_name && item.available !== false);
        }

        selectedItemIds = selectedItemIds.filter(itemId => allItems.some(item => item.id === itemId));
        populateCategories();
        populateItemTypes();
        renderItems(allItems);
        updateSelectedItems();
    } catch (error) {
        console.error(error);
        document.getElementById("items").innerHTML = "Unable to load items.";
    }
}

/* ==============================
   CATEGORY DROPDOWN
============================== */

function populateCategories(){

    const categories =
    [...new Set(
        allItems
            .map(x => x.category)
            .filter(Boolean)
    )]
    .sort();

    let html = "";

    categories.forEach(category => {

        html += `
        <div
            class="filter-chip"
            onclick="toggleCategory('${category}')">

            ${category}

        </div>
        `;
    });

    document
        .getElementById(
            "categoryChips"
        )
        .innerHTML = html;
}

function toggleCategory(category)

{

    if(
      selectedCategories.includes(
        category
      )
    ){

        selectedCategories =
        selectedCategories.filter(
            x => x !== category
        );

    }else{

        selectedCategories.push(
            category
        );
    }

    updateChipStyles();

    selectedTypes = [];
    
    populateItemTypes();
    
    applyFilters();
}

function populateItemTypes()
{

    let sourceItems =
    [...allItems];

    if(
       selectedCategories.length > 0
    ){

        sourceItems =
        sourceItems.filter(
            x =>
            selectedCategories.includes(
                x.category
            )
        );
    }

    const types =
    [...new Set(
        sourceItems.map(
            x => x.item_type
        )
    )]
    .sort();

    selectedTypes =
selectedTypes.filter(
    type => types.includes(type)
);

    let html = "";

    types.forEach(type => {

        html += `
        <div
            class="filter-chip"
            onclick="toggleType('${type}')">

            ${type}

        </div>
        `;
    });

    document
        .getElementById(
            "itemTypeChips"
        )
        .innerHTML = html;
}

function toggleType(type)
{

    if(
      selectedTypes.includes(type)
    ){

        selectedTypes =
        selectedTypes.filter(
          x => x !== type
        );

    }else{

        selectedTypes.push(type);
    }

    updateChipStyles();

    applyFilters();
}

function updateChipStyles(){

    document
        .querySelectorAll(
            ".filter-chip"
        )
        .forEach(chip => {

            chip.classList.remove(
                "active"
            );

            const value =
            chip.innerText.trim();

            if(
                selectedCategories.includes(
                    value
                )
                ||
                selectedTypes.includes(
                    value
                )
            ){

                chip.classList.add(
                    "active"
                );
            }
        });
}

/* ==============================
   RENDER ITEMS
============================== */

function renderItems(items){

    let html = "";

    items.forEach(item => {

        html += `
        <div class="item-row">

            <label>

            <input
            type="checkbox"
            value="${item.id}"
            ${selectedItemIds.includes(item.id) ? 'checked' : ''}
            onchange="toggleSelection('${item.id}')">

                <div class="item-details">

                    <div class="item-name">
                        ${item.item_name}
                    </div>

                    <div class="item-meta">
                        ${item.category}
                        •
                        ${item.item_type}
                    </div>

                    <div class="item-qty">
                        Qty Per Thaal:
                        ${item.qty_per_thaal}
                    </div>

                </div>

            </label>

        </div>
        `;
    });

    document
        .getElementById("items")
        .innerHTML = html;
}

/* ==============================
   FILTERING
============================== */

function applyFilters(){

    const search =
    document
      .getElementById(
         "searchBox"
      )
      .value
      .toLowerCase();

    let filtered =
    [...allItems];

    if(
      selectedCategories.length > 0
    ){

        filtered =
        filtered.filter(
            x =>
            selectedCategories.includes(
                x.category
            )
        );
    }

    if(
      selectedTypes.length > 0
    ){

        filtered =
        filtered.filter(
            x =>
            selectedTypes.includes(
                x.item_type
            )
        );
    }

    if(search){

        filtered =
        filtered.filter(
            x =>
            x.item_name
             .toLowerCase()
             .includes(search)
        );
    }

    renderItems(filtered);
}

document.addEventListener(
    "input",
    function (e) {

        if (
            e.target.id ===
            "searchBox"
        ) {
            applyFilters();
        }
    }
);

function toggleSelection(itemId){

    if(
        selectedItemIds.includes(itemId)
    ){

        selectedItemIds =
        selectedItemIds.filter(
            x => x !== itemId
        );

    }else{

        selectedItemIds.push(itemId);
    }

    updateSelectedItems();
}

/* ==============================
   SUBMIT FORM
============================== */

async function submitForm() {

    const selectedItems = selectedItemIds;

    if (selectedItems.length === 0) {
        alert("Please select at least one item.");
        return;
    }

    const name = document.getElementById("name").value.trim();
    const phone = document.getElementById("phone").value.trim();
    const email = document.getElementById("email").value.trim();
    const additional = document.getElementById("additional").value.trim();

    if (!name) {
        alert("Please enter your name.");
        return;
    }

    const phoneRegex = /^(?:\+61|0)4\d{8}$/;

    if (!phoneRegex.test(phone)) {
        alert("Please enter a valid Australian mobile number.");
        return;
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    if (!emailRegex.test(email)) {
        alert("Please enter a valid email address.");
        return;
    }

    const rows = buildExcelRows(name, phone, email, selectedItems, additional);

    if (BACKEND_URL && BACKEND_URL !== "PASTE_YOUR_GOOGLE_APPS_SCRIPT_URL_HERE") {
        try {
            const params = new URLSearchParams();
            params.append("action", "submitAllocation");
            params.append("centre", CENTRE_NAME);
            params.append("name", name);
            params.append("phone", phone);
            params.append("email", email);
            params.append("additional", additional || "");

            selectedItems.forEach(itemId => {
                const item = allItems.find(x => x.id === itemId);
                params.append("selectedItems", item ? item.item_name : itemId);
                params.append("selectedCategories", item ? item.category : "");
                params.append("selectedTypes", item ? item.item_type : "");
                params.append("selectedQuantities", item ? item.qty_per_thaal : "");
            });

            const backendResponse = await fetch(BACKEND_URL, {
                method: "POST",
                headers: { "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8" },
                body: params.toString()
            });

            const backendResult = await backendResponse.json();

            if (!backendResponse.ok || !backendResult.ok) {
                throw new Error("Backend submission failed.");
            }
        } catch (error) {
            console.error(error);
            alert("The live sheet update did not finish. Your local Excel download still completed.");
            return;
        }
    }

    markSelectedItemsUnavailable(selectedItems);
    renderItems(allItems.filter(item => item.available));
    exportToExcel(rows);

    alert("Jai Swaminarayan\n\nThank you for offering Annakut Seva.\n\nYour Excel submission file has been downloaded.");

    selectedItemIds = [];
    updateSelectedItems();
    document.getElementById("additional").value = "";
    document.getElementById("name").value = "";
    document.getElementById("phone").value = "";
    document.getElementById("email").value = "";
    document.querySelectorAll('input[type="checkbox"]').forEach(chk => {
        chk.checked = false;
    });
}

/* ==============================
   START APP
============================== */

loadItems();
setInterval(() => {
    if (document.visibilityState === "visible") {
        loadItems();
    }
}, 30000);

document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
        loadItems();
    }
});

function updateSelectedItems(){

    if(selectedItemIds.length === 0){

        document
            .getElementById(
                "selectedItems"
            )
            .innerHTML =
            "None Selected";

        return;
    }

    let html = "";

    selectedItemIds.forEach(itemId => {

        const item =
        allItems.find(
            x => x.id === itemId
        );

        if(item){

            html += `
            ✓ ${item.item_name}<br>
            `;
        }
    });

    document
        .getElementById(
            "selectedItems"
        )
        .innerHTML = html;
}
