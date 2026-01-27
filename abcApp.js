import { usageUrl, inventoryUrl, mainsapUrl } from './config.js';
import { toNumber, formatNumber, formatCurrency, formatShortCurrency, getPlantLabel, getPlantName, getPlantType } from './utils.js';

export let allData = [];
let filteredData = [];
let mode = "all", abcFilter = "", movingFilter = "";
let currentPage = 1;
const rowsPerPage = 25;
let lastBaseData = [];
let plantChart = null;
let sortKey = "CumPercent", sortDir = "asc";

function calculateKeepQty(r) {
    const avg = r.AvgMonthly;
    if (r.Moving === "Non Moving" || r.Moving === "Slowly") return 1;
    if (r.Moving === "Slow" && r.ABCValue === "C") return Math.max(1, Math.round(avg * 60));
    if (r.Moving === "Slow" && r.ABCValue === "B") return Math.max(1, Math.round(avg * 60));
    if (r.Moving === "Medium" && r.ABCValue === "C") return Math.max(2, Math.round(avg * 60));
    if (r.Moving === "Slow" && r.ABCValue === "A") return Math.max(3, Math.round(avg * 60));
    if (r.Moving === "Medium" && r.ABCValue === "B") return Math.max(4, Math.round(avg * 60));
    if (r.Moving === "Fast" && r.ABCValue === "C") return Math.max(5, Math.round(avg * 60));
    if (r.Moving === "Medium" && r.ABCValue === "A") return Math.max(6, Math.round(avg * 75));
    if (r.Moving === "Fast" && r.ABCValue === "B") return Math.max(8, Math.round(avg * 90));
    if (r.Moving === "Fast" && r.ABCValue === "A") return Math.max(10, Math.round(avg * 120));
    return Math.max(1, Math.round(avg * 40));
}

export async function loadData(isBackground = false) {
    const loader = document.getElementById("loader");
    const summarySection = document.getElementById("summarySection");
    const controlsSection = document.getElementById("controlsSection");
    const tableSection = document.getElementById("tableSection");
    const paginationSection = document.getElementById("pagination");

    if (!isBackground) {
        loader.style.display = "block";
        summarySection.style.display = "none";
        controlsSection.style.display = "none";
        tableSection.style.display = "none";
        paginationSection.style.display = "block";
    }

    try {
        const [uRes, iRes, mRes] = await Promise.all([fetch(usageUrl), fetch(inventoryUrl), fetch(mainsapUrl)]);
        const usage = await uRes.json();
        const inv = await iRes.json();
        const mainsap = await mRes.json();

        const useMap = new Map();
        const invMap = new Map();
        const mainsapMap = new Map();

        usage.forEach(r => {
            let plant = (r.Plant || '').trim();
            if (plant.length === 3) plant = "0" + plant;
            const material = (r.Material || '').trim();
            const key = `${plant}-${material}`;
            useMap.set(key, r);
        });
        inv.forEach(r => {
            let plant = (r.Plant || '').trim();
            if (plant.length === 3) plant = "0" + plant;
            const material = (r.Material || '').trim();
            const key = `${plant}-${material}`;
            invMap.set(key, r);
        });
        mainsap.forEach(r => {
            const material = (r.Material || '').trim();
            mainsapMap.set(material, {
                Note: r['หมายเหตุ'] || '',
                MultiplyUnit: toNumber(r['คูณหน่วย']),
                Product: r.Product || ''
            });
        });

        const allKeys = new Set([...useMap.keys(), ...invMap.keys()]);
        allData = Array.from(allKeys).map(key => {
            const [plant, material] = key.split('-');
            const u = useMap.get(key) || { Qtyissu: 0, Qtyissu6m: 0, '30Day': 0 };
            const i = invMap.get(key) || { Unrestricted: 0, "Value Unrestricted": 0, "Material description": '', "Base Unit of Measure": '' };
            const main = mainsapMap.get(material) || { Note: '', MultiplyUnit: 1, Product: '' };
            const qty4m = toNumber(u.Qtyissu);
            const qty6m = toNumber(u.Qtyissu6m) || Math.round(qty4m * 1.5);
            const thirtyDay = toNumber(u['30Day']);

            return {
                Plant: plant,
                Material: material,
                Description: i["Material description"] || '',
                Unit: i["Base Unit of Measure"] || '',
                Unrestricted: toNumber(i.Unrestricted),
                Value: toNumber(i["Value Unrestricted"]),
                Qty6Month: qty6m,
                Qty4Month: qty4m,
                ThirtyDay: thirtyDay,
                AvgDailyUse: qty4m / 120,
                AvgMonthly: qty4m / 4,
                SafetyStock: 0, ROP: 0, DOS: 0, RecommendedOrder: 0, Mean: 0,
                ABCValue: "", CumPercent: 0, Moving: "", IsNonMoving: qty4m === 0,
                ReturnQty: 0, ReturnValue: 0,
                Note: main.Note,
                MultiplyUnit: main.MultiplyUnit,
                Product: main.Product
            };
        });

        const descMap = new Map();
        allData.forEach(r => { if (r.Description) descMap.set(r.Material, r.Description); });
        allData.forEach(r => { if (!r.Description && descMap.has(r.Material)) r.Description = descMap.get(r.Material); });

        calcABCValue();
        fillPlantDropdown();
        setDefaultAndCalculate();

        if (!isBackground) {
            loader.style.display = "none";
            document.getElementById("mainCardsSection").style.display = "flex";
            summarySection.style.display = "flex";
            controlsSection.style.display = "flex";
            tableSection.style.display = "block";
            document.getElementById("pagination").style.display = "block";
        }
    } catch (e) {
        if (!isBackground) {
            loader.style.display = "none";
            summarySection.style.display = "flex";
            alert("โหลดข้อมูลไม่สำเร็จ");
        }
        console.error(e);
    }
}

function calcABCValue() {
    const sorted = [...allData].sort((a, b) => b.Value - a.Value);
    const total = sorted.reduce((s, r) => s + r.Value, 0);
    let cum = 0;
    sorted.forEach(r => {
        cum += r.Value;
        const pct = total ? (cum / total) * 100 : 0;
        const orig = allData.find(x => x.Material === r.Material && x.Plant === r.Plant);
        if (orig) {
            orig.CumPercent = pct;
            orig.ABCValue = pct <= 70 ? "A" : pct <= 90 ? "B" : "C";
        }
    });
}

function recalculateStockFields(data, params) {
    data.forEach(r => {
        const d = r.AvgDailyUse || 0;
        const leadTime = params.leadTime || 5;
        const safetyDays = params.safety || 3;
        const coverDays = params.cover || 40;
        const safetyStock = d * safetyDays;
        const rop = (d * leadTime) + safetyStock;
        const dos = d > 0 ? r.Unrestricted / d : 9999;
        r.SafetyStock = Math.round(safetyStock);
        r.ROP = Math.round(rop);
        r.DOS = dos > 9999 ? 9999 : parseFloat(dos.toFixed(1));

        let recommend = 0;
        if (r.Unrestricted < rop) {
            const needed = d * coverDays;
            recommend = needed - r.Unrestricted;
            if (recommend < 0) recommend = 0;
        }
        const multiply = r.MultiplyUnit || 1;
        if (multiply > 0) recommend = Math.ceil(recommend / multiply) * multiply;
        r.RecommendedOrder = Math.round(recommend);

        const avgMonthly = r.Qty4Month / 4;
        if (avgMonthly === 0) r.Moving = "Non Moving";
        else if (avgMonthly < 0.7) r.Moving = "Slowly";
        else if (avgMonthly < 12) r.Moving = "Slow";
        else if (avgMonthly < 30) r.Moving = "Medium";
        else r.Moving = "Fast";

        r.Mean = (r.ThirtyDay > r.AvgMonthly) ? r.ThirtyDay : r.AvgMonthly;

        if (r.Mean === 0 && (r.Moving === "Slow" || r.Moving === "Slowly")) r.RecommendedOrder = 0;
        if (r.ThirtyDay === 0 && (r.Moving === "Slow" || r.Moving === "Slowly")) r.RecommendedOrder = 0;
        if (r.ThirtyDay >= r.Qty4Month && r.Moving === "Slowly") r.RecommendedOrder = 0;

        const keepQty = calculateKeepQty(r);
        const returnQty = Math.max(0, r.Unrestricted - keepQty);
        const unitPrice = r.Unrestricted > 0 ? r.Value / r.Unrestricted : 0;
        r.ReturnQty = Math.round(returnQty);
        r.ReturnValue = returnQty * unitPrice;
    });
}

function fillPlantDropdown() {
    const sel = document.getElementById("plantFilter");
    const plants = [...new Set(allData.map(r => r.Plant))].filter(Boolean).sort();
    sel.innerHTML = '<option value="">ทุก Plant</option>';
    plants.forEach(code => sel.add(new Option(getPlantLabel(code), code)));
}

function setDefaultAndCalculate() {
    document.getElementById("plantFilter").value = "";
    mode = "all"; abcFilter = ""; movingFilter = "";
    applyFiltersAndRender();
}

function applyFiltersAndRender() {
    let data = [...allData];
    const plant = document.getElementById("plantFilter").value;
    if (plant) data = data.filter(r => r.Plant === plant);

    const search = document.getElementById("searchInput").value.toLowerCase();
    if (search) data = data.filter(r => r.Material.toLowerCase().includes(search) || r.Description.toLowerCase().includes(search));

    const params = {
        leadTime: parseInt(document.getElementById("leadTimeDays").value) || 5,
        safety: parseInt(document.getElementById("safetyDays").value) || 3,
        cover: parseInt(document.getElementById("coverDays").value) || 40
    };

    recalculateStockFields(data, params);
    const baseData = [...data];
    lastBaseData = baseData;

    let dataForABC = [...data], dataForMoving = [...data], dataForTable = [...data];

    if (mode === "onlyOrder") {
        const isOrder = (r) => Math.round(r.RecommendedOrder || 0) >= 1;
        dataForABC = dataForABC.filter(isOrder); dataForMoving = dataForMoving.filter(isOrder); dataForTable = dataForTable.filter(isOrder);
    } else if (mode === "returnable") {
        const isReturnable = (r) => r.ReturnQty > 0;
        dataForABC = dataForABC.filter(isReturnable); dataForMoving = dataForMoving.filter(isReturnable); dataForTable = dataForTable.filter(isReturnable);
    }

    if (abcFilter) {
        dataForMoving = dataForMoving.filter(r => r.ABCValue === abcFilter);
        dataForTable = dataForTable.filter(r => r.ABCValue === abcFilter);
    }
    if (movingFilter) {
        dataForABC = dataForABC.filter(r => r.Moving === movingFilter);
        dataForTable = dataForTable.filter(r => r.Moving === movingFilter);
    }

    dataForTable.sort((a, b) => {
        let x = a[sortKey] || 0, y = b[sortKey] || 0;
        if (typeof x === "string") x = x.toLowerCase();
        if (typeof y === "string") y = y.toLowerCase();
        return sortDir === "asc" ? (x > y ? 1 : -1) : (x < y ? 1 : -1);
    });

    filteredData = dataForTable;
    currentPage = 1;
    renderTable();
    renderPagination();
    updateAllCards(baseData, dataForABC, dataForMoving, dataForTable);
}

function updateAllCards(baseData, dataForABC, dataForMoving, tableData) {
    let totalStock = 0, orderItems = 0, orderValue = 0, returnCount = 0, totalReturnValue = 0, total_usage_4m_base = 0;
    baseData.forEach(r => {
        totalStock += r.Value;
        const unit_price = r.Unrestricted > 0 ? r.Value / r.Unrestricted : 0;
        const qtyOrder = Math.round(r.RecommendedOrder || 0);
        if (qtyOrder > 0) { orderItems++; orderValue += qtyOrder * unit_price; }
        if (r.ReturnQty > 0) { returnCount++; totalReturnValue += r.ReturnValue; }
        if (r.Qty4Month && unit_price > 0) total_usage_4m_base += r.Qty4Month * unit_price;
    });

    const monthly_withdrawal = total_usage_4m_base / 4 || 0;
    const daily_usage = total_usage_4m_base / 120 || 0;
    const stock_days = daily_usage > 0 ? Math.round(totalStock / daily_usage) : "N/A";
    const after_stock_days = daily_usage > 0 ? Math.round((totalStock - totalReturnValue) / daily_usage) : "N/A";

    const abc = { A: 0, B: 0, C: 0 }, cntABC = { A: 0, B: 0, C: 0 };
    dataForABC.forEach(r => { if (r.ABCValue && abc[r.ABCValue] !== undefined) { abc[r.ABCValue] += r.Value; cntABC[r.ABCValue]++; } });

    const mov = { Fast: 0, Medium: 0, Slow: 0, Slowly: 0, "Non Moving": 0 }, cntMov = { ...mov };
    dataForMoving.forEach(r => {
        let m = r.Moving === "Dead" ? "Non Moving" : r.Moving;
        if (mov[m] !== undefined) { mov[m] += r.Value; cntMov[m]++; }
    });

    document.getElementById("totalStockValue").textContent = formatShortCurrency(totalStock);
    document.getElementById("orderTotalValue").textContent = formatShortCurrency(orderValue);
    document.getElementById("returnValue").textContent = formatShortCurrency(totalReturnValue);
    document.getElementById("orderItemCount").textContent = orderItems;
    document.getElementById("returnItemCount").textContent = returnCount + " รายการ";
    document.getElementById("monthlyWithdrawalValue").textContent = formatShortCurrency(monthly_withdrawal);
    document.getElementById("stockDaysValue").textContent = typeof stock_days === "number" ? (stock_days > 9999 ? "มาก" : stock_days) : stock_days;
    document.getElementById("afterStockDaysValue").textContent = typeof after_stock_days === "number" ? (after_stock_days > 9999 ? "มาก" : after_stock_days) : after_stock_days;
    document.getElementById("totalCount").textContent = tableData.length + " รายการ";

    ["A", "B", "C"].forEach(k => {
        document.getElementById(`sum${k}`).textContent = formatShortCurrency(abc[k]);
        document.getElementById(`count${k}`).textContent = cntABC[k] + " รายการ";
    });
    ["Fast", "Medium", "Slow", "Slowly"].forEach(k => {
        document.getElementById(`val${k}`).textContent = formatShortCurrency(mov[k]);
        document.getElementById(`count${k}`).textContent = cntMov[k] + " รายการ";
    });
    document.getElementById("valDead").textContent = formatShortCurrency(mov["Non Moving"]);
    document.getElementById("countDead").textContent = cntMov["Non Moving"] + " รายการ";
}

function renderTable() {
    const container = document.getElementById("dataList");
    container.innerHTML = "";
    const start = (currentPage - 1) * rowsPerPage;
    const pageData = filteredData.slice(start, start + rowsPerPage);
    if (pageData.length === 0) {
        container.innerHTML = `<tr><td colspan="20" style="text-align:center;padding:50px;color:#95a5a6;">ไม่มีข้อมูลตามเงื่อนไข</td></tr>`;
        return;
    }
    pageData.forEach(r => {
        const orderQty = Math.round(r.RecommendedOrder);
        const orderClass = orderQty === 0 ? "order-0" : orderQty <= 10 ? "order-1-10" : orderQty <= 50 ? "order-11-50" : orderQty <= 200 ? "order-51-200" : "order-200";
        const row = document.createElement("tr");
        row.innerHTML = `
            <td>${getPlantLabel(r.Plant)}</td><td class="material">${r.Material}</td><td class="desc">${r.Description}</td><td>${r.Unit}</td>
            <td>${formatNumber(r.Unrestricted)}</td><td><span class="value">${formatCurrency(r.Value)}</span></td>
            <td>${formatNumber(r.Qty6Month)}</td><td>${formatNumber(r.Qty4Month)}</td>
            <td>${r.AvgMonthly.toFixed(1)}</td><td>${formatNumber(r.ThirtyDay)}</td><td>${r.Mean.toFixed(1)}</td>
            <td>${formatNumber(r.SafetyStock)}</td><td>${formatNumber(r.ROP)}</td><td>${r.DOS > 9999 ? 'มาก' : r.DOS.toFixed(1)}</td>
            <td><span class="${orderClass} order">${formatNumber(orderQty)}</span></td>
            <td>${r.Note}</td><td>${formatNumber(r.MultiplyUnit)}</td><td>${r.Product}</td>
            <td><strong>${r.ABCValue}</strong></td><td><span class="moving-${r.Moving.toLowerCase()}">${r.Moving}</span></td>
            <td><span class="return-qty">${formatNumber(r.ReturnQty)}</span></td><td>${r.CumPercent.toFixed(1)}%</td>
        `;
        container.appendChild(row);
    });
}

function renderPagination() {
    const totalPages = Math.max(1, Math.ceil(filteredData.length / rowsPerPage));
    const p = document.getElementById("pagination");
    p.innerHTML = "";
    const createBtn = (text, page, disabled) => {
        const btn = document.createElement("button");
        btn.className = "page-btn";
        if (page === currentPage) btn.classList.add("active");
        btn.textContent = text;
        btn.disabled = disabled;
        if (!disabled) btn.onclick = () => { currentPage = page; renderTable(); renderPagination(); };
        p.appendChild(btn);
    };
    createBtn("<<", 1, currentPage === 1);
    createBtn("<", currentPage - 1, currentPage === 1);
    createBtn(currentPage, currentPage, true);
    createBtn(">", currentPage + 1, currentPage === totalPages);
    createBtn(">>", totalPages, currentPage === totalPages);
}

function buildPlantChart() {
    const data = lastBaseData && lastBaseData.length ? lastBaseData : allData;
    const plantMap = new Map();
    data.forEach(r => {
        const code = r.Plant;
        if (!code) return;
        if (!plantMap.has(code)) plantMap.set(code, { plant: code, name: getPlantName(code), type: getPlantType(code), totalValue: 0, overValue: 0, stockValue: 0, usageValue4m: 0 });
        const agg = plantMap.get(code);
        agg.totalValue += r.Value;
        agg.overValue += r.ReturnValue;
        const unit_price = r.Unrestricted > 0 ? r.Value / r.Unrestricted : 0;
        agg.stockValue += r.Value;
        agg.usageValue4m += (r.Qty4Month || 0) * unit_price;
    });

    const summaries = [];
    plantMap.forEach(agg => {
        const dailyUsage = agg.usageValue4m / 120 || 0;
        summaries.push({ ...agg, stockDays: dailyUsage > 0 ? Math.round(agg.stockValue / dailyUsage) : 0 });
    });

    const ordered = [...summaries.filter(p => p.type === 'Company').sort((a, b) => b.stockDays - a.stockDays), ...summaries.filter(p => p.type === 'SA').sort((a, b) => b.stockDays - a.stockDays)];
    const labels = ordered.map(p => p.name ? `${p.plant} ${p.name}` : p.plant);
    
    const ctx = document.getElementById('plantChart').getContext('2d');
    if (plantChart) plantChart.destroy();
    plantChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                { label: 'มูลค่าคงเหลือ', type: 'bar', data: ordered.map(p => p.totalValue), yAxisID: 'y', borderWidth: 1 },
                { label: 'มูลค่า OverStock', type: 'line', data: ordered.map(p => p.overValue), yAxisID: 'y', borderWidth: 2, fill: false, borderDash: [5, 5] },
                { label: 'Stock Days', type: 'line', data: ordered.map(p => p.stockDays), yAxisID: 'y1', borderWidth: 2, fill: false },
                { label: 'เป้า Stock Days 60 วัน', type: 'line', data: new Array(labels.length).fill(60), yAxisID: 'y1', borderWidth: 2, fill: false, borderDash: [8, 4] }
            ]
        },
        options: { responsive: true, scales: { y: { beginAtZero: true, ticks: { callback: v => formatShortCurrency(v) } }, y1: { beginAtZero: true, position: 'right' } } }
    });
}

export function initABC() {
    document.getElementById("tableHeader").addEventListener("click", (e) => {
        const col = e.target.closest("th[data-sort]");
        if (!col) return;
        const key = col.getAttribute("data-sort");
        if (sortKey === key) sortDir = sortDir === "asc" ? "desc" : "asc";
        else { sortKey = key; sortDir = "desc"; }
        applyFiltersAndRender();
    });

    document.getElementById("cardTotalStock").onclick = () => { mode = "all"; abcFilter = ""; movingFilter = ""; applyFiltersAndRender(); };
    document.getElementById("cardOrderItems").onclick = () => { mode = "onlyOrder"; abcFilter = ""; movingFilter = ""; applyFiltersAndRender(); };
    document.getElementById("cardReturnValue").onclick = () => { mode = "returnable"; abcFilter = ""; movingFilter = ""; applyFiltersAndRender(); };
    
    ["A", "B", "C"].forEach(k => document.getElementById(`abc${k}`).onclick = () => { abcFilter = k; movingFilter = ""; applyFiltersAndRender(); });
    ["Fast", "Medium", "Slow", "Slowly"].forEach(k => document.getElementById(`card${k}`).onclick = () => { movingFilter = k; abcFilter = ""; applyFiltersAndRender(); });
    document.getElementById("cardDead").onclick = () => { movingFilter = "Non Moving"; abcFilter = ""; applyFiltersAndRender(); };

    document.getElementById("plantFilter").addEventListener("change", () => { mode = "all"; abcFilter = ""; movingFilter = ""; applyFiltersAndRender(); });
    ["leadTimeDays", "safetyDays", "coverDays", "searchInput"].forEach(id => document.getElementById(id).addEventListener("input", applyFiltersAndRender));

    document.getElementById("exportBtn").onclick = () => {
        const headers = ["Plant", "Material", "รายการอะไหล่", "หน่วย", "คงเหลือ", "มูลค่า", "ใช้ 6 เดือน", "ใช้ 4 เดือน", "เฉลี่ย/ด.", "30Day", "Mean", "Safety", "ROP", "DOS", "แนะนำสั่งซื้อ", "หมายเหตุ", "คูณหน่วย", "Product", "ABC", "Moving", "ส่งคืนได้ (ชิ้น)", "% สะสม"];
        let csv = "\uFEFF" + headers.join(",") + "\n";
        filteredData.forEach(r => {
            csv += [getPlantLabel(r.Plant), r.Material, r.Description.replace(/"/g, '""'), r.Unit, r.Unrestricted, r.Value, r.Qty6Month, r.Qty4Month, r.AvgDailyUse.toFixed(1), r.ThirtyDay, r.Mean.toFixed(1), Math.round(r.SafetyStock), Math.round(r.ROP), r.DOS > 9999 ? "มาก" : r.DOS.toFixed(1), Math.round(r.RecommendedOrder), r.Note, r.MultiplyUnit, r.Product, r.ABCValue, r.Moving, r.ReturnQty, r.CumPercent.toFixed(2)].map(v => `"${v}"`).join(",") + "\n";
        });
        const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
        const a = document.createElement("a"); a.href = url; a.download = `ABC_Analysis_${new Date().toISOString().slice(0, 10)}.csv`; a.click(); URL.revokeObjectURL(url);
    };

    document.getElementById('graphBtn').addEventListener('click', () => { document.getElementById('graphModal').style.display = 'flex'; buildPlantChart(); });
    document.getElementById('closeGraphBtn').addEventListener('click', () => { document.getElementById('graphModal').style.display = 'none'; });
    document.getElementById('helpBtn').addEventListener('click', () => { document.getElementById('helpModal').style.display = 'flex'; });
    document.getElementById('closeHelpBtn').addEventListener('click', () => { document.getElementById('helpModal').style.display = 'none'; });
}