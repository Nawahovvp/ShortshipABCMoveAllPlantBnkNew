import { loggedUser } from './auth.js';
import { shotShipReportDiffUrl, shotShipReportDailyUrl, shotShipReportMonthlyUrl, shotShipReportQuarterlyUrl, shotShipRemarksUrl, shotShipSaveUrl } from './config.js';
import { formatNumber, getPlantName, getPlantLabel } from './utils.js';
import { allData, loadData as loadABCData } from './abcApp.js';

let shotShipDiffData = [], shotShipDailyData = [], shotShipMonthlyData = [], shotShipQuarterlyData = [];
let filteredShotShipData = []; // For Table (Diff Items)
let currentShotShipPage = 1;
const shotShipRowsPerPage = 20;
let shotShipTypeFilter = "", shotShipPlantFilter = "", shotShipMovingFilter = "";
let shotShipChart = null;
let syncQueue = [];
let isSyncing = false;

export async function loadShotShipData() {
    const loader = document.getElementById("shotShipLoader");
    const tbody = document.getElementById("shotShipDataList");
    const content = document.getElementById("shotShipContent");

    loader.style.display = "block";
    content.style.display = "none";
    tbody.innerHTML = "";

    try {
        if (allData.length === 0) await loadABCData(true);

        // Fetch remarks in parallel
        const remarksPromise = fetch(shotShipRemarksUrl).then(r => r.json()).catch(() => []);

        const cachedReports = localStorage.getItem("shotShip_reports_cache");
        let diffs = [], daily = [], monthly = [], quarterly = [];

        if (cachedReports) {
            const parsed = JSON.parse(cachedReports);
            diffs = parsed.diffs; daily = parsed.daily; monthly = parsed.monthly; quarterly = parsed.quarterly;
        } else {
            const [resDiff, resDaily, resMonthly, resQuarterly] = await Promise.all([
                fetch(shotShipReportDiffUrl),
                fetch(shotShipReportDailyUrl),
                fetch(shotShipReportMonthlyUrl),
                fetch(shotShipReportQuarterlyUrl)
            ]);

            diffs = await resDiff.json();
            daily = await resDaily.json();
            monthly = await resMonthly.json();
            quarterly = await resQuarterly.json();

            try {
                localStorage.setItem("shotShip_reports_cache", JSON.stringify({ diffs, daily, monthly, quarterly }));
            } catch (e) { }
        }

        const remarksData = await remarksPromise;
        const remarksMap = new Map();
        if (Array.isArray(remarksData)) {
            remarksData.forEach(item => {
                // Support various column names from OpenSheet
                const d = item.DocNo || item.docNo || item['เลขที่ใบเบิก'];
                const m = item['Material Code'] || item.matCode || item.Material || item.material;
                const n = item.Note || item.note || item['หมายเหตุ'];
                const k = item.Key || item.key; // รองรับคอลัมน์ Key โดยตรง (เช่น 778761-30013193)
                if (k && n) {
                    remarksMap.set(String(k).trim(), n);
                } else if (d && m && n) {
                    remarksMap.set(`${String(d).trim()}-${String(m).trim()}`, n);
                }
            });
        }

        // --- PREPARE DIFF DATA (For Table & Moving Status) ---
        const movingMap = new Map();
        allData.forEach(r => movingMap.set(`${r.Plant}-${r.Material}`, r.Moving));

        diffs.forEach(r => {
            // Map Raw Report Columns to Internal Keys
            r.DocNo = r['เลขที่ใบเบิก'] || '';
            r.CreateDate = r['วันที่สร้างใบเบิก'] || '';
            r.ReceiveDate = r['วันที่รับโอน'] || ''; // New Column
            r.Date = r['วันที่'] || r['Date'] || '';
            r.Material = r['Material Code'] || '';
            r.Name = r['Material Name'] || '';
            r.Type = r['ประเภทอะไหล่'] || '';
            r.Requestor = r['ผู้ขอเบิก'] || '';
            r.Transferor = r['ผู้โอน'] || '';
            r.Req = parseFloat(r['จำนวนที่ขอเบิก'] || 0);
            r.Appr = parseFloat(r['จำนวนอนุมัติ'] || 0);
            r.Diff = parseFloat(r['ผลต่าง'] || 0);
            r.Plant = r['Plant'] || r['plant'] || '';
            if (r.Plant && String(r.Plant).length === 3) r.Plant = "0" + r.Plant;

            // Internal Logic
            r.RealAppr = r.Appr; // Use raw values for consistency with user request
            // If Diff exists but Appr is missing, maybe calculate? user didn't ask.

            // Date Cleaning
            if (r.Date && r.Date.includes(' ')) r.Date = r.Date.split(' ')[0];

            // Re-calc Moving Status
            const key = `${String(r.Plant || '').trim()}-${String(r.Material || '').trim()}`;
            let mv = movingMap.get(key) || "Non Moving";
            if (mv === "Dead") mv = "Non Moving";
            r.Moving = mv;

            // Note
            const keyNote = `${String(r.DocNo).trim()}-${String(r.Material).trim()}`;
            r.Note = remarksMap.get(keyNote) || r['หมายเหตุ'] || r['Note'] || '';
        });

        // Sync Queue Application
        syncQueue = JSON.parse(localStorage.getItem('shotShip_syncQueue') || '[]');
        syncQueue.forEach(item => {
            const row = diffs.find(r => String(r.DocNo).trim() === String(item.docNo).trim() && String(r.Material).trim() === String(item.matCode).trim());
            if (row) row.Note = item.note;
        });

        // Normalize Helper for Aggregated Reports
        const norm = (arr, dateKeyName) => {
            arr.forEach(r => {
                // Normalize Date
                const rawDate = r[dateKeyName] || r['Date'] || r['date'];
                if (rawDate) {
                    if (rawDate.includes(' ')) r.Date = rawDate.split(' ')[0]; // Split time
                    else r.Date = rawDate;
                }

                // Normalize Keys
                r.Plant = r['Plant'] || r['plant'] || '';
                if (r.Plant && String(r.Plant).length === 3) r.Plant = "0" + r.Plant;

                r['Part Type'] = r['Part Type'] || r['part type'] || r['Part_Type'] || r['part_type'] || '';
                r['Req Qty'] = parseFloat(r['Req Qty'] || r['req qty'] || r['Req_Qty'] || r['req_qty'] || 0);
                r['Real Appr Qty'] = parseFloat(r['Real Appr Qty'] || r['real appr qty'] || r['Real_Appr_Qty'] || r['real_appr_qty'] || 0);
                r['Req Items'] = parseInt(r['Req Items'] || r['req items'] || 0);
                r['Appr Items'] = parseInt(r['Appr Items'] || r['appr items'] || 0);
                r['Doc Count'] = parseInt(r['Doc Count'] || r['doc count'] || 0);

                // Ensure Month/Quarter keys exist if needed by specific logic
                if (dateKeyName === 'Month') r.Month = r.Date; // Logic uses r.Month
                if (dateKeyName === 'Quarter') r.Quarter = r.Date; // Logic uses r.Quarter
            });
        };

        if (daily) norm(daily, 'Date');
        if (monthly) norm(monthly, 'Month');
        if (quarterly) norm(quarterly, 'Quarter');

        // Helper to parse dd/mm/yyyy for sorting
        const parseSortDate = (dStr) => {
            if (!dStr) return 0;
            const parts = dStr.split('/');
            if (parts.length < 3) return 0;
            return new Date(parts[2], parts[1] - 1, parts[0]).getTime();
        };

        shotShipDiffData = diffs.sort((a, b) => parseSortDate(b.Date) - parseSortDate(a.Date));

        // Filter last 30 days
        if (shotShipDiffData.length > 0) {
            const latestTime = parseSortDate(shotShipDiffData[0].Date);
            const thirtyDaysMs = 30 * 24 * 60 * 60 * 1000;
            shotShipDiffData = shotShipDiffData.filter(r => (latestTime - parseSortDate(r.Date)) <= thirtyDaysMs);
        }

        shotShipDailyData = daily;
        shotShipMonthlyData = monthly;
        shotShipQuarterlyData = quarterly;

        populateShotShipFilters();
        shotShipTypeFilter = ""; shotShipPlantFilter = ""; shotShipMovingFilter = "";

        // Reset Filters UI
        document.getElementById('shotShipHeader').innerHTML = '<span style="color: #2c3e50; font-size: 1.5em; font-weight: 800;">ทุกพื้นที่</span>';
        const dateSelect = document.getElementById('shotShipDateFilter');

        let targetDate = "";
        // 1. Try to get latest date from Actual Table Data (Diffs)
        // shotShipDiffData is already sorted desc by date
        if (shotShipDiffData.length > 0 && shotShipDiffData[0].Date) {
            targetDate = shotShipDiffData[0].Date;
        }
        // 2. If no Table Data, fallback to latest from Filters (Daily Report)
        else if (dateSelect.options.length > 1) {
            targetDate = dateSelect.options[1].value;
        }

        if (targetDate) {
            // Verify option exists (it should now be in filters)
            dateSelect.value = targetDate;
            // Double check if setting value worked 
            if (dateSelect.value !== targetDate && dateSelect.options.length > 1) {
                // Format mismatch possibly? fallback to index 1
                dateSelect.selectedIndex = 1;
            }
        } else {
            dateSelect.value = "";
        }

        document.getElementById('shotShipNoteFilter').value = "empty";
        document.getElementById('shotShipQuarterFilter').value = "";

        applyShotShipFilters();
        updateSyncIndicator();

    } catch (e) {
        console.error(e);
        alert("โหลดข้อมูล ShotShip ไม่สำเร็จ");
    } finally {
        loader.style.display = "none";
        document.getElementById("shotShipContent").style.display = "block";
    }
}

function getShotShipDateParts(dateStr) {
    if (!dateStr) return { month: '', date: '' };
    const parts = dateStr.split(' ');
    const dParts = parts[0].split('/');
    if (dParts.length < 3) return { month: '', date: '' };
    return { month: `${dParts[1]}/${dParts[2]}`, date: parts[0] };
}

function populateShotShipFilters() {
    const monthSet = new Set(), dateSet = new Set(), quarterSet = new Set();

    // Dates from Daily Report
    shotShipDailyData.forEach(r => dateSet.add(r.Date));
    // Dates from Diff Data (Ensure table dates are selectable)
    shotShipDiffData.forEach(r => {
        if (r.Date) dateSet.add(r.Date);
    });

    // Months from Monthly Report
    shotShipMonthlyData.forEach(r => monthSet.add(r.Month));
    // Quarters from Quarterly Report
    shotShipQuarterlyData.forEach(r => quarterSet.add(r.Quarter));

    const monthSel = document.getElementById('shotShipMonthFilter');
    const dateSel = document.getElementById('shotShipDateFilter');
    const quarterSel = document.getElementById('shotShipQuarterFilter');

    monthSel.innerHTML = '<option value="">ทุกเดือน</option>';
    dateSel.innerHTML = '<option value="">ทุกวัน</option>';
    quarterSel.innerHTML = '<option value="">ทุกไตรมาส</option>';

    Array.from(monthSet).sort((a, b) => { const [m1, y1] = a.split('/'); const [m2, y2] = b.split('/'); return (y2 - y1) || (m2 - m1); }).forEach(m => monthSel.add(new Option(m, m)));
    // Sort Dates Descending
    Array.from(dateSet).sort((a, b) => { const [d1, m1, y1] = a.split('/'); const [d2, m2, y2] = b.split('/'); return (y2 - y1) || (m2 - m1) || (d2 - d1); }).forEach(d => dateSel.add(new Option(d, d)));
    Array.from(quarterSet).sort().forEach(q => quarterSel.add(new Option(q, q)));
}



function applyShotShipFilters(keepPage = false) {
    const search = document.getElementById('searchShotShipInput').value.toLowerCase();
    const month = document.getElementById('shotShipMonthFilter').value;
    const date = document.getElementById('shotShipDateFilter').value;
    const quarter = document.getElementById('shotShipQuarterFilter').value;
    const noteFilter = document.getElementById('shotShipNoteFilter').value;

    document.getElementById('shotShipHeader').innerHTML = shotShipPlantFilter ? `<span style="color: #e74c3c; font-size: 1.5em; font-weight: 800;">คลัง</span> <span style="color: #e74c3c; font-size: 1.5em; font-weight: 800; text-shadow: 2px 2px 4px rgba(0,0,0,0.1);">${getPlantName(shotShipPlantFilter)}</span>` : '<span style="color: #2c3e50; font-size: 1.5em; font-weight: 800;">ทุกพื้นที่</span>';

    // 1. FILTER DIFF ITEMS FOR TABLE (Same logic as before, but on diffData)
    const checkTableFilters = (r) => {
        const matchesSearch = !search || Object.values(r).some(val => String(val).toLowerCase().includes(search));

        // Date Matching Logic (Robust)
        // r.Date comes from Report_Diff_Items (cleaned to "d/m/yyyy" during load)
        // date filter comes from Report_Daily_30Days ("d/m/yyyy")
        // Just in case, trim both.
        const rDateClean = String(r.Date || '').trim();
        const filterDateClean = String(date || '').trim();

        const matchesDate = !date || rDateClean === filterDateClean;

        // Month Matching
        const { month: rMonth } = getShotShipDateParts(rDateClean);
        const matchesMonth = !month || rMonth === month;

        let matchesQuarter = true;
        if (quarter && rDateClean.includes('/')) {
            const parts = rDateClean.split('/');
            if (parts.length >= 2) {
                const m = parseInt(parts[1]);
                const y = parts[2]; // assuming d/m/yyyy
                matchesQuarter = `Q${Math.ceil(m / 3)}/${y}` === quarter;
            }
        }

        const matchesType = !shotShipTypeFilter || (r.Type || '').trim() === shotShipTypeFilter;
        // Moving Filter applies ONLY to Diff items (since we only calculate moving for diffs)
        const matchesMoving = !shotShipMovingFilter || (r.Moving || '').trim() === shotShipMovingFilter;

        const matchesPlant = !shotShipPlantFilter || (r.Plant || '').trim() === shotShipPlantFilter;

        // Note Filter
        const hasNote = r.Note && r.Note.trim() !== "";
        let matchesNote = true;
        if (noteFilter === 'empty') matchesNote = !hasNote;
        else if (noteFilter === 'not_empty') matchesNote = hasNote;

        return matchesSearch && matchesMonth && matchesDate && matchesQuarter && matchesType && matchesMoving && matchesPlant && matchesNote;
    };

    filteredShotShipData = shotShipDiffData.filter(checkTableFilters);

    // 2. PREPARE AGGREGATED DATA FOR DASHBOARD
    // Determine which Source to use based on filter hierarchy: Date > Month > Quarter > All (History?)
    // Actually, "All" history isn't requested, maybe default to Quarterly or Monthly?
    // Let's gather rows from the appropriate Summary Sheet.

    let summarySource = [];
    if (date) {
        summarySource = shotShipDailyData.filter(r => r.Date === date);
    } else if (month) {
        summarySource = shotShipMonthlyData.filter(r => r.Month === month);
    } else if (quarter) {
        summarySource = shotShipQuarterlyData.filter(r => r.Quarter === quarter);
    } else {
        // If no time filter, use Quarterly data (contains all quarters)
        summarySource = shotShipQuarterlyData;
    }

    // Filter Summary by Plant and Type (for dashboard calculation)
    const filteredSummary = summarySource.filter(r => {
        const matchesPlant = !shotShipPlantFilter || (r.Plant || '').trim() === shotShipPlantFilter;
        return matchesPlant;
        // Note: We don't filter Summary by Type here yet, we sum them up inside updateShotShipCards 
        // because the cards need to show breakdown of "General" vs "Consumable"
    });

    // We also need filteredDiffs for the Moving/Diff cards
    const filteredDiffsForCards = shotShipDiffData.filter(r => {
        const { month: rMonth, date: rDate } = getShotShipDateParts(r.Date);
        const matchesMonth = !month || rMonth === month;
        const matchesDate = !date || rDate === date;
        let matchesQuarter = true;
        if (quarter) {
            const parts = rDate.split('/');
            if (parts.length >= 3) {
                const m = parseInt(parts[1]);
                const y = parts[2];
                matchesQuarter = `Q${Math.ceil(m / 3)}/${y}` === quarter;
            }
        }
        const matchesPlant = !shotShipPlantFilter || (r.Plant || '').trim() === shotShipPlantFilter;
        return matchesMonth && matchesDate && matchesQuarter && matchesPlant;
    });

    // NEW: Diffs for Chart (Time filtered, ALL Plants) for adjusting chart calculations
    const diffsForChart = shotShipDiffData.filter(r => {
        const { month: rMonth, date: rDate } = getShotShipDateParts(r.Date);
        const matchesMonth = !month || rMonth === month;
        const matchesDate = !date || rDate === date;
        let matchesQuarter = true;
        if (quarter) { const parts = rDate.split('/'); if (parts.length >= 3) { const m = parseInt(parts[1]); const y = parts[2]; matchesQuarter = `Q${Math.ceil(m / 3)}/${y}` === quarter; } }
        return matchesMonth && matchesDate && matchesQuarter;
    });

    updateShotShipCards(filteredSummary, filteredDiffsForCards); // Pass both sources

    // For Chart: Use aggregated data (Plant grouped) AND diffs for adjustment
    renderShotShipChart(summarySource, shotShipPlantFilter, diffsForChart);

    if (!keepPage) currentShotShipPage = 1;
    renderShotShipTable();
    renderShotShipPagination();
}

// Updated to accept (SummaryData, DiffData)
function updateShotShipCards(summaryData, diffData) {
    let totalReq = 0, totalAppr = 0, totalDiff = 0, totalReqItems = 0, totalApprItems = 0;
    let uniqueDocs = 0; // Need logic to sum unique docs? 
    // Base Summation from Summary Data

    let totalGeneral = 0, totalApprGeneral = 0, totalConsumable = 0, totalApprConsumable = 0;

    summaryData.forEach(r => {
        // Only include if matches active Type filter (if set)?
        // Dashboard usually shows ALL types unless filtered?
        // UI code has specific cards for General/Consumable.
        // If Type filter is active, we should probably still show the Totals of that Type.

        if (shotShipTypeFilter && r['Part Type'] !== shotShipTypeFilter) return;

        totalReq += parseFloat(r['Req Qty'] || 0);
        totalAppr += parseFloat(r['Real Appr Qty'] || 0);
        // From summary, we don't have "Total Diff", but we know Diff = Req - RealAppr
        // totalDiff += (parseFloat(r['Req Qty']) - parseFloat(r['Real Appr Qty']));

        totalReqItems += parseInt(r['Req Items'] || 0);
        totalApprItems += parseInt(r['Appr Items'] || 0);

        // Doc Count - this is a sum of counts from groups. It might double count if a Doc has both types.
        // User accepted this limitation implicitly by moving to aggregated reports.
        uniqueDocs += parseInt(r['Doc Count'] || 0);

        if (r['Part Type'] === 'อะไหล่ทั่วไป') {
            totalGeneral += parseInt(r['Req Items'] || 0);
            totalApprGeneral += parseInt(r['Appr Items'] || 0);
        } else if (r['Part Type'] === 'อะไหล่สิ้นเปลือง') {
            totalConsumable += parseInt(r['Req Items'] || 0);
            totalApprConsumable += parseInt(r['Appr Items'] || 0);
        }
    });

    // ★ NEW: Calculate Adjustments from Diff Data (Items with Notes)
    // If Note exists -> Treat as No Diff -> Add Diff back to Appr
    let adjApprQty = 0;
    let adjApprItems = 0;
    let adjApprGeneral = 0;
    let adjApprConsumable = 0;

    diffData.forEach(r => {
        if (shotShipTypeFilter && r.Type !== shotShipTypeFilter) return;

        const hasNote = r.Note && r.Note.trim() !== "";
        if (hasNote) {
            const diff = parseFloat(r.Diff || 0);
            // Add Diff back to Appr (Treat as fully approved)
            adjApprQty += diff;

            // If it was 0 approved, it now becomes an approved item
            if (parseFloat(r.Appr || 0) === 0) {
                adjApprItems++;
                if (r.Type === 'อะไหล่ทั่วไป') adjApprGeneral++;
                else if (r.Type === 'อะไหล่สิ้นเปลือง') adjApprConsumable++;
            }
        }
    });

    // Apply Adjustments
    totalAppr += adjApprQty;
    totalApprItems += adjApprItems;
    totalApprGeneral += adjApprGeneral;
    totalApprConsumable += adjApprConsumable;

    totalDiff = totalReq - totalAppr;
    if (totalDiff < 0) totalDiff = 0;

    // Diff Cards (Source: diffData)
    let diffItemsCount = 0, diffGeneralCount = 0, diffConsumableCount = 0;
    let fastCount = 0, mediumCount = 0, slowCount = 0, slowlyCount = 0, deadCount = 0;

    diffData.forEach(r => {
        if (shotShipTypeFilter && r.Type !== shotShipTypeFilter) return;

        const diff = parseFloat(r.Diff || 0);
        const hasNote = r.Note && r.Note.trim() !== "";

        // Only count as Diff if Diff > 0 AND No Note
        if (diff > 0 && !hasNote) {
            diffItemsCount++;
            if (r.Type === 'อะไหล่ทั่วไป') diffGeneralCount++;
            else if (r.Type === 'อะไหล่สิ้นเปลือง') diffConsumableCount++;

            const moving = (r.Moving || '').trim();
            if (moving === 'Fast') fastCount++;
            else if (moving === 'Medium') mediumCount++;
            else if (moving === 'Slow') slowCount++;
            else if (moving === 'Slowly') slowlyCount++;
            else if (moving === 'Dead' || moving === 'Non Moving') deadCount++;
        }
    });

    document.getElementById('ssTotalDocs').textContent = formatNumber(uniqueDocs);
    document.getElementById('ssTotalReqItems').textContent = formatNumber(totalReqItems);
    document.getElementById('ssTotalApprItems').textContent = formatNumber(totalApprItems);
    document.getElementById('ssTotalReq').textContent = formatNumber(totalReq);
    document.getElementById('ssTotalAppr').textContent = formatNumber(totalAppr);
    document.getElementById('ssTotalDiff').textContent = formatNumber(totalDiff);
    document.getElementById('ssPercentAppr').textContent = (totalReq > 0 ? (totalAppr / totalReq) * 100 : 0).toFixed(2) + '%';

    document.getElementById('ssTypeGeneral').textContent = formatNumber(totalGeneral);
    document.getElementById('ssApprGeneral').textContent = formatNumber(totalApprGeneral);
    document.getElementById('ssTypeConsumable').textContent = formatNumber(totalConsumable);
    document.getElementById('ssApprConsumable').textContent = formatNumber(totalApprConsumable);

    document.getElementById('ssDiffItems').textContent = formatNumber(diffItemsCount);
    document.getElementById('ssDiffGeneral').textContent = formatNumber(diffGeneralCount);
    document.getElementById('ssDiffConsumable').textContent = formatNumber(diffConsumableCount);

    document.getElementById('ssFast').textContent = formatNumber(fastCount);
    document.getElementById('ssMedium').textContent = formatNumber(mediumCount);
    document.getElementById('ssSlow').textContent = formatNumber(slowCount);
    document.getElementById('ssSlowly').textContent = formatNumber(slowlyCount);
    document.getElementById('ssDead').textContent = formatNumber(deadCount);
}

function renderShotShipChart(data, highlightPlant = "", diffData = []) {
    const plantMap = new Map();
    // 1. Base Aggregation from Summary Data
    data.forEach(r => {
        if (shotShipTypeFilter && r['Part Type'] !== shotShipTypeFilter) return;

        const req = parseFloat(r['Req Qty'] || 0);
        const appr = parseFloat(r['Real Appr Qty'] || 0);
        const reqItems = parseInt(r['Req Items'] || 0);
        const apprItems = parseInt(r['Appr Items'] || 0);

        const plantCode = (r['Plant'] || 'Unknown').trim();
        if (!plantMap.has(plantCode)) plantMap.set(plantCode, { req: 0, appr: 0, reqItems: 0, apprItems: 0 });
        const p = plantMap.get(plantCode);
        p.req += req; p.appr += appr;
        p.reqItems += reqItems; p.apprItems += apprItems;
    });

    // 2. Apply Adjustments from Diff Data (Notes)
    if (diffData && diffData.length > 0) {
        diffData.forEach(r => {
            if (shotShipTypeFilter && r.Type !== shotShipTypeFilter) return;
            
            const hasNote = r.Note && r.Note.trim() !== "";
            if (hasNote) {
                const plantCode = (r.Plant || 'Unknown').trim();
                if (plantMap.has(plantCode)) {
                    const p = plantMap.get(plantCode);
                    const diff = parseFloat(r.Diff || 0);
                    p.appr += diff; // Add diff back to appr
                    
                    if (parseFloat(r.Appr || 0) === 0) {
                        p.apprItems++;
                    }
                }
            }
        });
    }

    const plantArray = [];
    plantMap.forEach((val, key) => plantArray.push({ plant: key, name: getPlantLabel(key), percent: val.req > 0 ? (val.appr / val.req) * 100 : 0, reqQty: val.req, apprQty: val.appr, ...val }));
    plantArray.sort((a, b) => b.percent - a.percent);

    const labels = plantArray.map(item => item.name);
    const values = plantArray.map(item => item.percent);
    const backgroundColors = values.map((v, i) => highlightPlant ? (plantArray[i].plant === highlightPlant ? (v >= 98 ? 'rgba(39, 174, 96, 1)' : 'rgba(231, 76, 60, 1)') : 'rgba(200, 200, 200, 0.3)') : (v >= 98 ? 'rgba(39, 174, 96, 0.7)' : 'rgba(231, 76, 60, 0.7)'));
    const borderColors = values.map((v, i) => highlightPlant ? (plantArray[i].plant === highlightPlant ? (v >= 98 ? 'rgba(39, 174, 96, 1)' : 'rgba(231, 76, 60, 1)') : 'rgba(200, 200, 200, 0.5)') : (v >= 98 ? 'rgba(39, 174, 96, 1)' : 'rgba(231, 76, 60, 1)'));

    const ctx = document.getElementById('shotShipChart').getContext('2d');
    if (shotShipChart) shotShipChart.destroy();
    shotShipChart = new Chart(ctx, {
        type: 'bar',
        data: { labels, datasets: [{ label: 'ประสิทธิภาพในการจ่าย (%)', data: values, backgroundColor: backgroundColors, borderColor: borderColors, borderWidth: 1, order: 2 }, { label: 'Target 98%', data: new Array(labels.length).fill(98), type: 'line', borderColor: '#27ae60', borderWidth: 2, borderDash: [5, 5], pointRadius: 0, fill: false, order: 1 }] },
        plugins: [{ id: 'barLabels', afterDatasetsDraw(chart) { const { ctx } = chart; chart.data.datasets.forEach((dataset, i) => { if (dataset.type === 'line') return; const meta = chart.getDatasetMeta(i); meta.data.forEach((bar, index) => { const value = dataset.data[index]; if (value != null) { ctx.fillStyle = '#2c3e50'; ctx.font = 'bold 12px sans-serif'; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(value.toFixed(2) + '%', bar.x, bar.y - 5); } }); }); } }],
        options: {
            responsive: true, maintainAspectRatio: false,
            onClick: (e, elements) => { if (elements.length > 0) { const index = elements[0].index; const selectedPlant = plantArray[index].plant; shotShipPlantFilter = (shotShipPlantFilter === selectedPlant) ? "" : selectedPlant; applyShotShipFilters(); } },
            scales: { y: { beginAtZero: true, title: { display: true, text: '% ประสิทธิภาพในการจ่าย' }, suggestedMax: 115 }, x: { ticks: { maxRotation: 45, minRotation: 45 } } },
            plugins: { tooltip: { callbacks: { label: (context) => { if (context.dataset.type === 'line') return `${context.dataset.label}: ${context.parsed.y}%`; const item = plantArray[context.dataIndex]; return [`${context.dataset.label}: ${context.parsed.y.toFixed(2)}%`, `รายการเบิก: ${formatNumber(item.reqItems)}`, `รายการที่จ่าย: ${formatNumber(item.apprItems)}`, `จำนวนขอเบิก: ${formatNumber(item.reqQty)} ชิ้น`, `จำนวนจ่าย: ${formatNumber(item.apprQty)} ชิ้น`]; } } } }
        }
    });
}

// --- Render Functions (Restored) ---

function renderShotShipTable() {
    const tbody = document.getElementById("shotShipDataList");
    tbody.innerHTML = "";
    const start = (currentShotShipPage - 1) * shotShipRowsPerPage;
    const pageData = filteredShotShipData.slice(start, start + shotShipRowsPerPage);
    if (pageData.length === 0) { tbody.innerHTML = `<tr><td colspan="14" style="text-align:center;padding:20px;">ไม่มีข้อมูล</td></tr>`; return; }

    pageData.forEach(r => {
        let noteHtml = r.Note || '';
        if (r.Note === 'จำนวนผิด') noteHtml = `<span class="note-badge note-qty">${r.Note}</span>`;
        else if (r.Note === 'เบิกผิด') noteHtml = `<span class="note-badge note-pick">${r.Note}</span>`;
        else if (r.Note === 'Code ผิด') noteHtml = `<span class="note-badge note-code">${r.Note}</span>`;
        const diffVal = parseFloat(r.Diff || 0); // Corrected to use internal key
        const diffHtml = `<span class="diff-badge ${diffVal > 0 ? 'diff-high' : 'diff-zero'} clickable-diff" title="คลิกเพื่อแจ้งปัญหา">${formatNumber(diffVal)}</span>`;
        const row = document.createElement("tr");
        row.innerHTML = `<td>${r.DocNo || ''}</td><td>${r.ReceiveDate || ''}</td><td>${r.Material || ''}</td><td style="text-align: left;">${r.Name || ''}</td><td>${r.Type || ''}</td><td style="text-align: left;">${r.Requestor || ''}</td><td style="text-align: left;">${r.Transferor || ''}</td><td>${r.Req || ''}</td><td>${r.Appr || ''}</td><td>${diffHtml}</td><td>${r.Plant || ''}</td><td><span class="moving-${(r.Moving || '').toLowerCase().replace(/ /g, '-')}">${r.Moving || ''}</span></td><td>${noteHtml}</td><td>${r.Date || ''}</td>`;
        tbody.appendChild(row);
    });
}

function renderShotShipPagination() {
    const totalPages = Math.max(1, Math.ceil(filteredShotShipData.length / shotShipRowsPerPage));
    const p = document.getElementById("shotShipPagination");
    p.style.display = "flex"; p.innerHTML = "";
    const createBtn = (text, page, disabled) => {
        const btn = document.createElement("button");
        btn.className = "page-btn";
        if (page === currentShotShipPage && !isNaN(text)) btn.classList.add("active");
        btn.textContent = text; btn.disabled = disabled;
        if (!disabled) btn.onclick = () => { currentShotShipPage = page; renderShotShipTable(); renderShotShipPagination(); };
        p.appendChild(btn);
    };
    createBtn("<<", 1, currentShotShipPage === 1); createBtn("<", currentShotShipPage - 1, currentShotShipPage === 1);
    createBtn(currentShotShipPage, currentShotShipPage, true);
    createBtn(">", currentShotShipPage + 1, currentShotShipPage >= totalPages); createBtn(">>", totalPages, currentShotShipPage >= totalPages);
}


// (Removed duplicate updateShotShipCards and renderShotShipChart)


// ★ NEW: Sync Queue Functions
function saveSyncQueue() {
    localStorage.setItem('shotShip_syncQueue', JSON.stringify(syncQueue));
    updateSyncIndicator();
}

function updateSyncIndicator() {
    const indicator = document.getElementById('syncStatusDisplay');
    if (!indicator) return;

    const count = syncQueue.length;
    if (count > 0) {
        indicator.style.display = 'flex';
        indicator.innerHTML = `<i class="fas fa-cloud-upload-alt ${isSyncing ? 'fa-pulse' : ''}" style="font-size: 16px;"></i><span style="font-size: 10px; font-weight: bold;">${count}</span>`;
        indicator.style.background = isSyncing ? '#27ae60' : '#f39c12';
    } else {
        indicator.style.display = 'none';
    }
}

async function processSyncQueue() {
    if (isSyncing || syncQueue.length === 0) return;
    if (!navigator.onLine) return;

    isSyncing = true;
    updateSyncIndicator();

    try {
        while (syncQueue.length > 0) {
            if (!navigator.onLine) break;

            const item = syncQueue[0];
            try {
                await fetch(shotShipSaveUrl, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(item)
                });

                syncQueue.shift();
                saveSyncQueue();
                await new Promise(r => setTimeout(r, 500));
            } catch (e) {
                console.error("Sync failed, retrying later", e);
                break;
            }
        }
    } finally {
        isSyncing = false;
        updateSyncIndicator();
    }
}

export function initShotShip() {
    document.getElementById('searchShotShipInput').addEventListener('input', () => applyShotShipFilters());
    document.getElementById('shotShipDateFilter').addEventListener('change', () => applyShotShipFilters());
    document.getElementById('shotShipQuarterFilter').addEventListener('change', () => applyShotShipFilters());
    document.getElementById('shotShipNoteFilter').addEventListener('change', () => applyShotShipFilters());
    document.getElementById('shotShipMonthFilter').addEventListener('change', () => {
        const selectedMonth = document.getElementById('shotShipMonthFilter').value;
        const dateSel = document.getElementById('shotShipDateFilter');
        const currentVal = dateSel.value;
        const dateSet = new Set();
        // Use shotShipDailyData for dates map
        shotShipDailyData.forEach(r => {
            const { month, date } = getShotShipDateParts(r.Date);
            if (!selectedMonth || month === selectedMonth) if (date) dateSet.add(date);
        });

        dateSel.innerHTML = '<option value="">ทุกวัน</option>';
        Array.from(dateSet).sort((a, b) => { const [d1, m1, y1] = a.split('/'); const [d2, m2, y2] = b.split('/'); return (y2 - y1) || (m2 - m1) || (d2 - d1); }).forEach(d => dateSel.add(new Option(d, d)));
        if (Array.from(dateSel.options).some(o => o.value === currentVal)) dateSel.value = currentVal;
        applyShotShipFilters();
    });

    document.getElementById('cardSsTotalDocs').addEventListener('click', () => { document.getElementById('searchShotShipInput').value = ""; shotShipTypeFilter = ""; shotShipMovingFilter = ""; shotShipPlantFilter = ""; document.getElementById('shotShipNoteFilter').value = "empty"; document.getElementById('shotShipQuarterFilter').value = ""; applyShotShipFilters(); });
    document.getElementById('cardSsTypeGeneral').addEventListener('click', () => { document.getElementById('searchShotShipInput').value = ""; shotShipTypeFilter = "อะไหล่ทั่วไป"; applyShotShipFilters(); });
    document.getElementById('cardSsTypeConsumable').addEventListener('click', () => { document.getElementById('searchShotShipInput').value = ""; shotShipTypeFilter = "อะไหล่สิ้นเปลือง"; applyShotShipFilters(); });
    document.getElementById('cardSsDiffItems').addEventListener('click', () => { document.getElementById('searchShotShipInput').value = ""; shotShipTypeFilter = ""; applyShotShipFilters(); });
    document.getElementById('cardSsDiffGeneral').addEventListener('click', () => { document.getElementById('searchShotShipInput').value = ""; shotShipTypeFilter = "อะไหล่ทั่วไป"; applyShotShipFilters(); });
    document.getElementById('cardSsDiffConsumable').addEventListener('click', () => { document.getElementById('searchShotShipInput').value = ""; shotShipTypeFilter = "อะไหล่สิ้นเปลือง"; applyShotShipFilters(); });
    ["Fast", "Medium", "Slow", "Slowly"].forEach(k => document.getElementById(`cardSs${k}`).addEventListener('click', () => { document.getElementById('searchShotShipInput').value = ""; shotShipMovingFilter = k; applyShotShipFilters(); }));
    document.getElementById('cardSsDead').addEventListener('click', () => { document.getElementById('searchShotShipInput').value = ""; shotShipMovingFilter = "Non Moving"; applyShotShipFilters(); });

    const editModal = document.getElementById('editShotShipModal');
    document.querySelectorAll('.note-btn').forEach(btn => btn.addEventListener('click', function () { document.querySelectorAll('.note-btn').forEach(b => b.classList.remove('active')); this.classList.add('active'); document.getElementById('editNote').value = this.getAttribute('data-value'); }));
    document.getElementById('shotShipDataList').addEventListener('click', (e) => {
        if (e.target.classList.contains('clickable-diff')) {
            const tr = e.target.closest('tr');
            document.getElementById('editDocNo').value = tr.cells[0].textContent;
            document.getElementById('editMatCode').value = tr.cells[2].textContent;
            document.getElementById('editMatName').value = tr.cells[3].textContent;
            document.getElementById('editPartType').value = tr.cells[4].textContent;
            const currentNote = tr.cells[12].textContent.trim();
            document.getElementById('editNote').value = currentNote;
            document.querySelectorAll('.note-btn').forEach(btn => { btn.classList.remove('active'); if (btn.getAttribute('data-value') === currentNote) btn.classList.add('active'); });
            editModal.style.display = 'flex';
        }
    });
    document.getElementById('btnCancelEdit').addEventListener('click', () => editModal.style.display = 'none');
    window.addEventListener('click', (e) => { if (e.target === editModal) editModal.style.display = 'none'; });
    document.getElementById('btnConfirmEdit').addEventListener('click', async () => {
        const payload = {
            action: 'save',
            docNo: document.getElementById('editDocNo').value,
            matCode: document.getElementById('editMatCode').value,
            matName: document.getElementById('editMatName').value,
            partType: document.getElementById('editPartType').value,
            note: document.getElementById('editNote').value,
            user: loggedUser ? loggedUser.Name : 'Unknown'
        };
        const btn = document.getElementById('btnConfirmEdit'); btn.textContent = 'กำลังบันทึก...'; btn.disabled = true;

        // ★ NEW: Optimistic UI Update & Queueing
        try {
            // 1. เพิ่มลงคิวและบันทึก LocalStorage
            syncQueue.push(payload);
            saveSyncQueue();

            // 2. อัปเดตข้อมูลในหน้าจอทันที (Update Memory)
            // Use shotShipDiffData here
            shotShipDiffData.forEach(r => {
                if (String(r.DocNo || '').trim() === String(payload.docNo).trim() &&
                    String(r.Material || '').trim() === String(payload.matCode).trim()) {
                    r.Note = payload.note;
                }
            });

            // 3. ปิด Modal และรีเฟรชหน้าจอทั้งหมด (Update DOM / Re-render)
            editModal.style.display = 'none';
            // เรียก applyShotShipFilters(true) เพื่อคำนวณ Cards, Chart และ Table ใหม่
            // โดย keepPage=true เพื่อให้ยังอยู่หน้าเดิมในตาราง
            applyShotShipFilters(true);

            // 4. แจ้งเตือน
            const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000, timerProgressBar: true });
            Toast.fire({ icon: 'success', title: 'บันทึกแล้ว (รอซิงค์)' });

            processSyncQueue();

        } catch (e) { console.error(e); Swal.fire({ icon: 'error', title: 'เกิดข้อผิดพลาด', text: 'ไม่สามารถบันทึกข้อมูลได้' }); } finally { btn.textContent = 'ยืนยัน'; btn.disabled = false; }
    });
    const clearBtn = document.getElementById('clearCacheBtn');
    if (clearBtn) clearBtn.addEventListener('click', () => { if (confirm("ต้องการล้างแคชข้อมูล ShotShip หรือไม่?")) { localStorage.removeItem("shotShip_reports_cache"); if (document.getElementById('shotShipApp').style.display !== 'none') loadShotShipData(); } });

    window.addEventListener('online', processSyncQueue);
    setTimeout(processSyncQueue, 2000);
}