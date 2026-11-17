// UI layer: rendering, event handlers, and app bootstrap.
import "./index.css";
import * as XLSX from "xlsx";
import { appState, logisticsService, bindUI } from "./services.js";
import { ICONS, GAS_STATION_FRIENDLY_NAMES, ROUTE_STATION_COLUMNS } from "./constants.js";
import { parseFuelBaseRows } from "./importer.js";
import { detectSheetType, parseUphillQueue, parseDownhillQueue, parsePartnerQueue, parseDailyRows } from "./importerTrips.js";
import { parseDrivers, parseRateTable, parseVehicles } from "./importerMasters.js";
import { parseRateSheet, buildAuditContext, auditTrip, applyCarry, flagDoubleDownhill, AUDIT_CLASSES, AUDIT_STATUS, DOWN_STATUS, normalizeOriginExtras, parseDriverAliasSheet, applyDriverAlias } from "./fuelAudit.js";
import { COLL_STATUS, downhillTripKeys, planDownhillFuel, withDownhillFuel, reconcileColumnL, createAllIfAbsent } from "./downhillFuel.js";
import { planDateFix, buildDateFixBatches, recheckBatch, prevDay } from "./dateFix.js";
import { TEMPLATE_SHEET, buildTemplateLists, generateTemplate, findTemplateHeader, parseTemplateRows, planTemplateSave, saveTemplateTrips } from "./tripTemplate.js";

const formatZeroAsEmpty = (value) => {
            const num = Number(value);
            if (!num || num === 0) return '';
            return num.toLocaleString();
        };

        // Local calendar date as YYYY-MM-DD (avoids the UTC off-by-one from toISOString()).
        const localDateStr = (d = new Date()) => {
            const tzOffsetMs = d.getTimezoneOffset() * 60000;
            return new Date(d.getTime() - tzOffsetMs).toISOString().split('T')[0];
        };

        // Loading overlay control — hidden once the first batch of Firebase data has loaded.
        let initialLoadDone = false;
        const hideLoadingOverlay = () => {
            if (initialLoadDone) return;
            initialLoadDone = true;
            const overlay = document.getElementById('loading-overlay');
            if (overlay) {
                overlay.style.opacity = '0';
                setTimeout(() => { overlay.style.display = 'none'; }, 300);
            }
        };

        // Undo toast: shown after a delete so the user can restore the item within a few seconds.
        let undoToastTimer = null;
        const showUndoToast = (message, onUndo) => {
            const existing = document.getElementById('undo-toast');
            if (existing) existing.remove();
            if (undoToastTimer) clearTimeout(undoToastTimer);

            const toast = document.createElement('div');
            toast.id = 'undo-toast';
            toast.className = 'fixed bottom-6 left-1/2 -translate-x-1/2 z-[70] flex items-center gap-4 bg-slate-800 border border-slate-600 text-white px-5 py-3 rounded-lg shadow-2xl';
            toast.innerHTML = `<span>${message}</span>
                <button id="undo-toast-btn" class="bg-accent hover:bg-secondary text-white font-bold py-1 px-4 rounded-md transition">เลิกทำ</button>`;
            document.body.appendChild(toast);

            const dismiss = () => {
                if (undoToastTimer) clearTimeout(undoToastTimer);
                if (toast.parentNode) toast.remove();
            };
            document.getElementById('undo-toast-btn').addEventListener('click', async () => {
                dismiss();
                try {
                    await onUndo();
                } catch (e) {
                    console.error('Undo failed:', e);
                    alert('เลิกทำไม่สำเร็จ: ' + e.message);
                }
            });
            undoToastTimer = setTimeout(dismiss, 7000);
        };

        const renderRouteDistributionRow = (item, index, unit) => {
            const unitText = unit === 'Liters' ? 'ลิตร' : 'บาท';
            const gasStationOptions = Object.entries(GAS_STATION_FRIENDLY_NAMES)
                .map(([key, name]) => `<option value="${key}" ${item.stationKey === key ? 'selected' : ''}>${name}</option>`)
                .join('');
            return `
                <div class="distribution-row grid grid-cols-12 gap-2 items-center mb-2" data-index="${index}" data-unit="${unit}">
                    <div class="col-span-6">
                        <select data-field="stationKey" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent">
                            <option value="">เลือกปั๊มน้ำมัน</option>
                            ${gasStationOptions}
                        </select>
                    </div>
                    <div class="col-span-5">
                        <input type="number" data-field="amount" value="${item.amount || ''}" placeholder="${unitText}" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" />
                    </div>
                    <div class="col-span-1 flex justify-center">
                        <button type="button" data-action="remove-distribution" class="text-red-500 hover:text-red-400 p-2 rounded-full transition-colors">${ICONS.Trash}</button>
                    </div>
                </div>
            `;
        };

        const renderEditDistributionRow = (item, index, unit) => {
            // This is identical to renderRouteDistributionRow
            return renderRouteDistributionRow(item, index, unit);
        };
        
        const renderApp = () => {
            const appContainer = document.getElementById('app-container');
            let viewHtml = '';
            if (appState.view === 'form') viewHtml = renderFuelRecordForm();
            if (appState.view === 'dashboard') viewHtml = renderDashboard();
            if (appState.view === 'data') viewHtml = renderDataManagement();
            if (appState.view === 'vehicles') viewHtml = renderVehicleManagement();
            if (appState.view === 'import') viewHtml = renderImport();
            if (appState.view === 'queue') viewHtml = renderQueue();
            if (appState.view === 'drivers') viewHtml = renderDrivers();
            if (appState.view === 'payroll') viewHtml = renderPayroll();
            if (appState.view === 'reports') viewHtml = renderReports();
            if (appState.view === 'audit') viewHtml = renderAudit();

            appContainer.innerHTML = `
                ${renderHeader()}
                <main class="container mx-auto px-6 py-8">
                    ${viewHtml}
                </main>
                <footer class="text-center py-4 text-gray-500 text-sm">
                    <p>NeoSiamLogistics & Transport Fuel Management © 2024</p>
                </footer>
                ${renderModal()}
                ${renderAuthModal()}
            `;
            attachEventListeners();
        };

        const renderHeader = () => {
            const navButton = (targetView, icon, text) => `
                <button
                    data-view="${targetView}"
                    class="nav-button flex items-center gap-2 px-4 py-2 rounded-md transition-colors duration-200 ${
                        appState.view === targetView
                            ? 'bg-accent text-white font-semibold shadow-lg'
                            : 'bg-neutral/60 hover:bg-neutral text-gray-300'
                    }"
                >
                    ${icon}
                    <span class="hidden sm:inline">${text}</span>
                </button>
            `;

            return `
                <header class="bg-neutral shadow-lg sticky top-0 z-30">
                    <nav class="container mx-auto px-4 sm:px-6 py-3 flex flex-wrap justify-between items-center gap-3">
                        <div class="flex items-center gap-3">
                            <div class="bg-gray-200 p-2 rounded-full">
                               <svg xmlns="http://www.w3.org/2000/svg" class="w-8 h-8" viewBox="0 0 24 24" fill="#0369a1"><path d="M12 22a7 7 0 0 0 7-7c0-2-1-3.9-3-5.5s-3.5-4-4-6.5c-.5 2.5-2 4.9-4 6.5C6 11.1 5 13 5 15a7 7 0 0 0 7 7z"></path></svg>
                            </div>
                            <h1 class="text-xl sm:text-2xl font-bold text-white tracking-wider">NeoSiamLogistics & Transport Fuel</h1>
                        </div>
                        <div class="flex flex-wrap items-center gap-2">
                            ${navButton('form', ICONS.Truck, 'บันทึกน้ำมัน')}
                            ${navButton('dashboard', ICONS.ChartBar, 'แดชบอร์ด')}
                            ${navButton('queue', ICONS.Truck, 'คิวรถ')}
                            ${navButton('payroll', ICONS.ChartBar, 'การเงิน')}
                            ${navButton('reports', ICONS.ChartBar, 'รายงาน')}
                            ${navButton('audit', ICONS.Truck, 'ตรวจการเติม')}
                            ${navButton('drivers', ICONS.IdCard, 'คนขับ')}
                            ${navButton('vehicles', ICONS.IdCard, 'จัดการทะเบียนรถ')}
                            ${navButton('data', ICONS.Database, 'จัดการเส้นทาง')}
                            ${navButton('import', ICONS.Download, 'นำเข้า Excel')}
                        </div>
                    </nav>
                </header>
            `;
        };

        const renderCard = (content, className = '') => `<div class="bg-neutral rounded-xl shadow-lg p-6 ${className}">${content}</div>`;

        const renderAutocomplete = (id, label, placeholder, options) => {
            return `
                <div class="relative autocomplete-container" data-id="${id}">
                    <label class="block text-sm font-medium text-gray-400 mb-1">${label}</label>
                    <input
                        type="text"
                        id="${id}"
                        placeholder="${placeholder}"
                        autocomplete="off"
                        class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent"
                    />
                    <div id="${id}-suggestions" class="hidden absolute z-20 w-full mt-1 bg-slate-800 border border-slate-600 rounded-md shadow-lg max-h-60 overflow-y-auto"></div>
                </div>
            `;
        };

        // --- View: Trip entry (บันทึกน้ำมัน) — full trip form + live route card, saves to trips ---
        const renderFuelRecordForm = () => {
            const driverNames = [...new Set(logisticsService.getDrivers().map(d => d.name).filter(Boolean))];
            const runTypes = (() => { const rt = [...new Set(logisticsService.getRateTable().map(r => r.runType).filter(Boolean))]; return rt.length ? rt : ['หนักล่อง', 'ตีเปล่า', 'คิวพิเศษ']; })();
            const money2 = (id, label) => `<div><label class="block text-xs text-gray-400 mb-1">${label}</label><input type="number" id="${id}" class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>`;
            return `
                <div class="max-w-6xl mx-auto grid grid-cols-1 md:grid-cols-3 gap-8">
                    <div class="md:col-span-2">
                        ${renderCard(`
                            <h2 class="text-2xl font-bold text-white mb-4 border-b border-slate-600 pb-2 flex items-center gap-2">${ICONS.Truck} เพิ่มคิวรถ / บันทึกงานวิ่งใหม่</h2>
                            <form id="fuel-record-form" class="space-y-5">
                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                    <div><label class="block text-sm font-medium text-gray-400 mb-1">วันที่</label>
                                        <input type="date" id="form-date" value="${localDateStr()}" class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>
                                    <div><label class="block text-sm font-medium text-gray-400 mb-1">ทิศทาง</label>
                                        <select id="form-direction" class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white"><option value="ขาขึ้น">ขาขึ้น</option><option value="ขาล่อง">ขาล่อง</option></select></div>
                                </div>
                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                    ${renderAutocomplete('form-vehicle-registration', 'ทะเบียนรถ', 'ค้นหา/เลือกทะเบียน', logisticsService.getVehicleRegistrations())}
                                    ${renderAutocomplete('form-vehicle-type', 'ประเภทรถ', 'ค้นหา/เลือกประเภท', logisticsService.getUniqueValues('vehicleType'))}
                                </div>
                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                    ${renderAutocomplete('form-driver', 'คนขับ', 'ค้นหา/เลือกคนขับ', driverNames)}
                                    ${renderAutocomplete('form-queue-type', 'ประเภทคิว/วิ่ง', 'หนักล่อง/ตีเปล่า/คิวพิเศษ', runTypes)}
                                </div>
                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                    ${renderAutocomplete('form-origin', 'ต้นทาง', 'ค้นหา/เลือกต้นทาง', logisticsService.getUniqueValues('origin'))}
                                    ${renderAutocomplete('form-destination', 'ปลายทาง', 'ค้นหา/เลือกปลายทาง', logisticsService.getUniqueValues('destination'))}
                                </div>
                                <div><label class="block text-sm font-medium text-gray-400 mb-1">รายการงาน</label>
                                    <input type="text" id="form-job-detail" class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>

                                <div class="bg-sky-900/20 border border-sky-800 rounded-lg p-2 text-xs text-sky-200">💡 กรอก ประเภทรถ + ต้นทาง + ปลายทาง → ดูการ์ด "ข้อมูลเส้นทาง (อัตโนมัติ)" ทางขวาว่าต้องเติมปั๊มไหน เท่าไร แล้วกด "ใช้ค่าจากเส้นทางนี้" เพื่อเติมให้อัตโนมัติ (กันเติมผิดปั๊ม/ผิดจำนวน)</div>

                                <div class="space-y-3 rounded-lg border border-slate-600 p-4">
                                    <div class="flex justify-between items-center"><h3 class="font-semibold text-lg text-white">น้ำมัน (ลิตร)</h3>
                                        <div class="flex items-center gap-2"><label class="text-sm text-gray-400">รวม</label><input type="number" id="form-total-liters" readonly class="w-24 text-right bg-slate-900 border border-slate-600 rounded-md px-3 py-1 text-white" /></div></div>
                                    <div id="distribution-list-liters"></div>
                                    <button type="button" id="add-distribution-btn-liters" class="w-full border-2 border-dashed border-slate-500 hover:bg-slate-700/50 text-gray-300 text-sm py-1.5 rounded-lg">+ เพิ่มปั๊ม (ลิตร)</button>
                                </div>
                                <div class="space-y-3 rounded-lg border border-slate-600 p-4">
                                    <div class="flex justify-between items-center"><h3 class="font-semibold text-lg text-white">น้ำมัน (บาท)</h3>
                                        <div class="flex items-center gap-2"><label class="text-sm text-gray-400">รวม</label><input type="number" id="form-total-baht" readonly class="w-24 text-right bg-slate-900 border border-slate-600 rounded-md px-3 py-1 text-white" /></div></div>
                                    <div id="distribution-list-baht"></div>
                                    <button type="button" id="add-distribution-btn-baht" class="w-full border-2 border-dashed border-slate-500 hover:bg-slate-700/50 text-gray-300 text-sm py-1.5 rounded-lg">+ เพิ่มปั๊ม (บาท)</button>
                                </div>

                                <div class="grid grid-cols-2 sm:grid-cols-3 gap-3">
                                    ${money2('form-revenue', 'รายรับ / เงินที่ต้องได้')}
                                    ${money2('form-net', 'สุทธิโอน / รวมเงิน')}
                                    ${money2('form-fuelcost', 'ค่าน้ำมัน / ค่าแก๊ส')}
                                    ${money2('form-allowance', 'เบี้ยขับ / ค่าเดินทาง')}
                                    ${money2('form-deduction', 'หักเงิน')}
                                    ${money2('form-distance', 'ระยะทาง (กม.)')}
                                </div>
                                <button type="button" data-action="calc-allowance-form" class="text-sm bg-sky-700 hover:bg-sky-600 text-white py-1.5 px-3 rounded-lg">🧮 คำนวณ เบี้ยเลี้ยง / หัก / สุทธิโอน จากตารางเบี้ยเลี้ยง</button>

                                <div><label class="block text-sm font-medium text-gray-400 mb-1">หมายเหตุ</label>
                                    <textarea id="form-notes" rows="2" class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white"></textarea></div>
                                <div class="flex justify-end pt-2 gap-4">
                                    <button type="button" id="reset-fuel-form-btn" class="bg-slate-600 hover:bg-slate-500 text-white font-bold py-2 px-6 rounded-lg">ล้างข้อมูล</button>
                                    <button type="submit" class="bg-primary hover:bg-secondary text-white font-bold py-2 px-6 rounded-lg">บันทึกคิวรถ</button>
                                </div>
                                <div id="form-status" class="mt-2"></div>
                            </form>
                        `)}
                    </div>
                    <div class="md:col-span-1">
                        ${renderCard('<div id="vlookup-card-content"></div>', 'sticky top-28')}
                    </div>
                </div>
            `;
        };

        const renderDashboardContent = () => {
            const filters = appState.dashboardFilters;
            const day = filters.day || '';
            const plateFilter = (filters.vehicleRegistration || '').toUpperCase();
            let trips = logisticsService.getTrips();
            if (day) trips = trips.filter(t => t.date === day);
            if (plateFilter) trips = trips.filter(t => (t.plate || '').toUpperCase().includes(plateFilter));

            const totalLiters = trips.reduce((s, t) => s + (t.totalLiters || 0), 0);
            const totalBaht = trips.reduce((s, t) => s + (t.totalBaht || 0), 0);

            const friendlyNames = logisticsService.getAllGasStationFriendlyNames();

            // Start from the known station catalog, then union any station keys that actually
            // appear in the trips — so refuels at an unrecognized/new pump are still shown and
            // counted instead of being dropped silently (which would make the per-pump totals
            // add up to less than the grand total above).
            const stationKeys = [...logisticsService.getAllGasStationKeys()];
            trips.forEach(t => {
                new Set([...Object.keys(t.gasStationsLiters || {}), ...Object.keys(t.gasStationsBaht || {})])
                    .forEach(k => { if (!stationKeys.includes(k)) stationKeys.push(k); });
            });

            // Group trips by the station(s) they refueled at.
            const grp = {}; stationKeys.forEach(k => grp[k] = []);
            trips.forEach(t => {
                new Set([...Object.keys(t.gasStationsLiters || {}), ...Object.keys(t.gasStationsBaht || {})])
                    .forEach(k => { if (grp[k]) grp[k].push(t); });
            });

            const stationTablesHtml = stationKeys.map(key => {
                const list = grp[key];
                const sumL = list.reduce((s, t) => s + (t.gasStationsLiters?.[key] || 0), 0);
                const sumB = list.reduce((s, t) => s + (t.gasStationsBaht?.[key] || 0), 0);
                const rows = list.map(t => {
                    const aL = t.gasStationsLiters?.[key] || 0, aB = t.gasStationsBaht?.[key] || 0;
                    const txt = [aL > 0 ? `${aL.toLocaleString()} ลิตร` : '', aB > 0 ? `${aB.toLocaleString()} บาท` : ''].filter(Boolean).join('<br>');
                    return `<tr class="border-b border-slate-700 hover:bg-slate-700/20">
                        <td class="p-2 whitespace-nowrap">${t.date || ''}</td>
                        <td class="p-2 font-mono whitespace-nowrap">${t.plate || ''}</td>
                        <td class="p-2 whitespace-nowrap">${t.origin || ''} → ${t.destination || ''}</td>
                        <td class="p-2 text-right font-semibold text-accent whitespace-nowrap">${txt}</td></tr>`;
                }).join('');
                return renderCard(`
                    <div class="flex justify-between items-start mb-3 gap-2">
                        <h3 class="text-lg font-bold flex items-center gap-2">⛽ ${friendlyNames[key] || key}</h3>
                        <div class="text-right"><div class="text-accent font-bold whitespace-nowrap">${sumL.toLocaleString()} ลิตร${sumB > 0 ? ` / ${sumB.toLocaleString()} บาท` : ''}</div><div class="text-xs text-gray-500">${list.length} ครั้ง</div></div>
                    </div>
                    <div class="overflow-x-auto max-h-[280px]">
                        <table class="w-full text-left text-sm">
                            <thead class="bg-slate-800 text-gray-400 uppercase sticky top-0 z-10"><tr><th class="p-2">วันที่</th><th class="p-2">ทะเบียน</th><th class="p-2">เส้นทาง</th><th class="p-2 text-right">ปริมาณ</th></tr></thead>
                            <tbody>${list.length ? rows : `<tr><td colspan="4" class="text-center py-6 text-gray-500">ไม่มีการเติมที่ปั๊มนี้</td></tr>`}</tbody>
                        </table>
                    </div>`);
            }).join('');

            const period = day ? `วันที่ ${day}` : `เดือน ${appState.tripsMonth || '(เลือกเดือน)'}`;
            return `
                <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                    ${renderCard(`<h3 class="text-lg text-gray-400">จำนวนคิว</h3><p class="text-4xl font-bold text-accent">${trips.length.toLocaleString()}</p>`, 'text-center')}
                    ${renderCard(`<h3 class="text-lg text-gray-400">น้ำมันรวม</h3><p class="text-4xl font-bold text-accent">${totalLiters.toLocaleString()} <span class="text-2xl">ลิตร</span></p>`, 'text-center')}
                    ${renderCard(`<h3 class="text-lg text-gray-400">น้ำมันรวม (บาท)</h3><p class="text-4xl font-bold text-accent">${totalBaht.toLocaleString()}<span class="text-2xl"> บาท</span></p>`, 'text-center')}
                </div>
                <div class="space-y-4">
                    <h2 class="text-2xl font-bold text-white mt-6 border-b border-slate-600 pb-2">⛽ สรุปยอดเติมน้ำมันรายปั๊ม <span class="text-base text-gray-400 font-normal">(${period})</span></h2>
                    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">${stationTablesHtml}</div>
                </div>
            `;
        }


        // --- View: Dashboard (fuel-by-pump summary from trips) ---
        const renderDashboard = () => {
            const month = appState.tripsMonth || localDateStr().slice(0, 7);
            const f = appState.dashboardFilters;
            return `
                <div class="space-y-6">
                    ${renderCard(`
                        <div class="flex flex-wrap justify-between items-end gap-4 mb-2">
                            <h2 class="text-2xl font-bold text-white flex items-center gap-3">${ICONS.ChartBar} แดชบอร์ดน้ำมัน (รายปั๊ม)</h2>
                            <div class="flex flex-wrap items-end gap-3">
                                <div><label class="block text-xs text-gray-400 mb-1">เดือน</label>
                                    <input type="month" id="dashboard-month" value="${month}" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>
                                <div><label class="block text-xs text-gray-400 mb-1">เจาะจงวัน (ถ้าต้องการ)</label>
                                    <input type="date" id="dashboard-day" value="${f.day || ''}" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>
                                <div><label class="block text-xs text-gray-400 mb-1">กรองทะเบียน</label>
                                    <input type="text" id="dashboard-plate" value="${f.vehicleRegistration || ''}" placeholder="ทะเบียน" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white w-40" /></div>
                                <button id="dashboard-clear-day" class="bg-slate-600 hover:bg-slate-500 text-white py-2 px-3 rounded-lg text-sm">ดูทั้งเดือน</button>
                                <button id="export-excel-btn" class="flex items-center gap-2 bg-success hover:bg-green-600 text-white font-bold py-2 px-4 rounded-lg shadow-md">${ICONS.Download} Export</button>
                            </div>
                        </div>
                        <div id="dashboard-content" class="space-y-6">${renderDashboardContent()}</div>
                    `)}
                </div>
            `;
        };

        // Routes as shown on the จัดการเส้นทาง page (search filter + sort) — shared by the table and the Excel export.
        const getVisibleRoutes = () => {
            const filter = appState.dataManagementFilter.toLowerCase();
            const routes = logisticsService.getRoutes();
            const filteredRoutes = filter ? routes.filter(r =>
                (r.vehicleType || '').toLowerCase().includes(filter) ||
                (r.origin || '').toLowerCase().includes(filter) ||
                (r.destination || '').toLowerCase().includes(filter)
            ) : routes; // getRoutes() already returns a copy, so the in-place sort below is safe

            // Sort routes for consistent display
            filteredRoutes.sort((a, b) => {
                const vehicleTypeCompare = (a.vehicleType || '').localeCompare(b.vehicleType || '', 'th');
                if (vehicleTypeCompare !== 0) return vehicleTypeCompare;

                const originCompare = (a.origin || '').localeCompare(b.origin || '', 'th');
                if (originCompare !== 0) return originCompare;

                return (a.destination || '').localeCompare(b.destination || '', 'th');
            });
            return filteredRoutes;
        };

        const renderDataManagementContent = () => {
            const filteredRoutes = getVisibleRoutes();
            const stationCols = ROUTE_STATION_COLUMNS;

            const tableRows = filteredRoutes.map(route => {
                const stationCells = stationCols.map(col => {
                    return col.units.map(unit => {
                        const value = (unit === 'Liters')
                            ? route.gasStationsLiters?.[col.key]
                            : route.gasStationsBaht?.[col.key];
                        return `<td class="px-2 py-2 text-right font-mono text-xs border-r border-slate-700">${formatZeroAsEmpty(value)}</td>`;
                    }).join('');
                }).join('');

                const totalLitersText = formatZeroAsEmpty(route.totalLiters);
                const totalBahtText = formatZeroAsEmpty(route.totalBaht);

                return `
                <tr key="${route.id}" class="border-b border-slate-600 hover:bg-slate-700/20">
                    <td class="p-3 border-r border-slate-700 whitespace-nowrap">${route.vehicleType || ''}</td>
                    <td class="p-3 border-r border-slate-700 whitespace-nowrap">${route.origin || ''}</td>
                    <td class="p-3 border-r border-slate-700 whitespace-nowrap">${route.destination || ''}</td>
                    <td class="p-3 border-r border-slate-700 text-center whitespace-nowrap">${route.direction || ''}</td>
                    <td class="p-3 border-r border-slate-700 text-center whitespace-nowrap">${route.crossingFee || ''}</td>
                    <td class="p-3 border-r border-slate-700 text-center whitespace-nowrap">${route.waitingFee || ''}</td>
                    <td class="p-3 text-right font-semibold text-accent border-r border-slate-700 whitespace-nowrap">
                        ${totalLitersText ? `<div>${totalLitersText} L</div>` : ''}
                        ${totalBahtText ? `<div>${totalBahtText} B</div>` : ''}
                    </td>
                    ${stationCells}
                    <td class="p-3 text-center">
                        <div class="flex justify-center gap-2">
                            <button data-action="edit-route" data-id="${route.id}" class="text-blue-400 hover:text-blue-300 p-2 rounded-full transition-colors">${ICONS.Pencil}</button>
                            <button data-action="delete-route" data-id="${route.id}" class="text-red-500 hover:text-red-400 p-2 rounded-full transition-colors">${ICONS.Trash}</button>
                        </div>
                    </td>
                </tr>
            `}).join('');

            // Data-driven station header rows (auto-adjusts to ROUTE_STATION_COLUMNS).
            const stationHeadTop = stationCols.map(col =>
                `<th colspan="${col.units.length}" class="px-2 py-2 text-center border-b border-r border-slate-600 whitespace-nowrap text-xs">${GAS_STATION_FRIENDLY_NAMES[col.key] || col.key}</th>`
            ).join('');
            const stationHeadSub = stationCols.map(col =>
                col.units.map(unit =>
                    `<th class="p-2 text-center bg-yellow-600/80 text-white border-r border-slate-600">${unit === 'Liters' ? 'ลิตร' : 'บาท'}</th>`
                ).join('')
            ).join('');

            return renderCard(`
                <p class="text-sm text-sky-300 mb-2 flex items-center gap-2">${ICONS.Refresh} ตารางมีปั๊มน้ำมันทั้งหมด ${stationCols.length} ปั๊ม — เลื่อนตารางไปทางขวา ⟶ เพื่อดูให้ครบ</p>
                <div class="overflow-x-auto">
                    <table class="w-full text-left">
                        <thead class="bg-slate-800 text-sm text-gray-400 uppercase sticky top-0 z-10">
                            <tr>
                                <th rowspan="2" class="p-3 border-r border-slate-600 align-middle whitespace-nowrap">ประเภทรถ</th>
                                <th rowspan="2" class="p-3 border-r border-slate-600 align-middle whitespace-nowrap">ต้นทาง</th>
                                <th rowspan="2" class="p-3 border-r border-slate-600 align-middle whitespace-nowrap">ปลายทาง</th>
                                <th rowspan="2" class="p-3 border-r border-slate-600 align-middle text-center whitespace-nowrap">ทิศทาง</th>
                                <th rowspan="2" class="p-3 border-r border-slate-600 align-middle text-center whitespace-nowrap">ค่าข้าม</th>
                                <th rowspan="2" class="p-3 border-r border-slate-600 align-middle text-center whitespace-nowrap">ค่าค้าง</th>
                                <th rowspan="2" class="p-3 text-right border-r border-slate-600 align-middle whitespace-nowrap">จำนวน<br>น้ำมัน</th>
                                ${stationHeadTop}
                                <th rowspan="2" class="p-3 text-center align-middle whitespace-nowrap">จัดการ</th>
                            </tr>
                            <tr>
                                ${stationHeadSub}
                            </tr>
                        </thead>
                        <tbody id="routes-table-body">${tableRows}</tbody>
                    </table>
                    ${filteredRoutes.length === 0 ? '<p class="text-center py-8 text-gray-500">ไม่พบข้อมูลเส้นทาง</p>' : ''}
                </div>
            `);
        };

        // --- View: Data Management ---
        const renderDataManagement = () => {
            return `
                <div class="space-y-8">
                    ${renderCard(`
                        <div class="flex flex-wrap justify-between items-center gap-4">
                            <h2 class="text-2xl font-bold text-white flex items-center gap-3">${ICONS.Database}จัดการข้อมูลเส้นทาง</h2>
                            <div class="flex flex-wrap items-center gap-4">
                                <input type="text" id="data-management-filter" value="${appState.dataManagementFilter}" placeholder="ค้นหา..." class="w-full md:w-64 bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" />
                                <button id="refresh-routes-btn" title="รีเฟรชข้อมูล" class="bg-sky-600 hover:bg-sky-700 text-white font-bold p-2 rounded-lg shadow-md transition-transform transform hover:scale-105">
                                    ${ICONS.Refresh}
                                </button>
                                <button id="backup-data-btn" title="สำรองข้อมูลทั้งหมด (เส้นทาง + บันทึกน้ำมัน + ทะเบียนรถ) เป็นไฟล์ JSON" class="flex items-center gap-2 bg-success hover:bg-green-600 text-white font-bold py-2 px-4 rounded-lg shadow-md transition-transform transform hover:scale-105">
                                    ${ICONS.Download} สำรองข้อมูล
                                </button>
                                <button id="routes-export-btn" title="Export เส้นทางที่แสดงอยู่ (ตามคำค้นหา) เป็นไฟล์ Excel" class="flex items-center gap-2 bg-emerald-700 hover:bg-emerald-800 text-white font-bold py-2 px-4 rounded-lg shadow-md transition-transform transform hover:scale-105 whitespace-nowrap">
                                    ${ICONS.Download} Export Excel
                                </button>
                                <button id="add-new-route-btn" class="flex items-center gap-2 bg-primary hover:bg-secondary text-white font-bold py-2 px-4 rounded-lg shadow-md transition-transform transform hover:scale-105">
                                    ${ICONS.Plus} เพิ่มเส้นทางใหม่
                                </button>
                            </div>
                        </div>
                    `)}
                    <div id="data-management-content">
                        ${renderDataManagementContent()}
                    </div>
                </div>
            `;
        };

        // --- View: Import from Excel ---
        let importWorkbook = null;
        let importParsedRoutes = null;
        let importParsedTrips = null;
        let importParsedMaster = null;
        let importParsedStandards = null;
        let importParsedTemplate = null; // { rows, header, fileName } from the Excel trip template
        let importParsedAliases = null;  // { aliases, skipped, unknown } from the confirmed driver-name file

        // Escape text taken from Excel/Firebase before putting it into innerHTML.
        const esc = (v) => String(v ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
        const pumpName = (key) => GAS_STATION_FRIENDLY_NAMES[key] || key;
        const fmtPumps = (obj) => Object.entries(obj || {}).filter(([, v]) => Number(v)).map(([k, v]) => `${esc(pumpName(k))} ${Number(v).toLocaleString()}`).join(' + ') || '-';
        // A rate's uphill liters; trailer-drop rates have parts (แม่/ลูก) and a total but no pump.
        const fmtRateUp = (s) => s.extraLiters
            ? `อัตราสิบล้อ + เพิ่มไม่เกิน ${Number(s.extraLiters)} ลิตร <span class="text-xs text-gray-400">(${esc(s.ruleText)})</span>`
            : s.parts?.length
            ? `${s.parts.map(p => `${esc(p.part)}→${esc(p.destination)} ${Number(p.liters).toLocaleString()}`).join(' + ')} = ${Number(s.upTotal).toLocaleString()} (ไม่ระบุปั๊ม)`
            : fmtPumps(s.up);

        const renderRateImportPreview = ({ standards, warnings, skippedSheets, fileName }) => {
            const flagged = standards.filter(s => s.flags.length);
            const cur = appState.rateStandardsMeta;
            const rows = standards.map(s => `
                <tr class="border-b border-slate-700 ${s.flags.length ? 'bg-amber-900/20' : ''}">
                    <td class="p-2 whitespace-nowrap">${esc(s.sheet)} · ${s.row}</td>
                    <td class="p-2 whitespace-nowrap">${esc(AUDIT_CLASSES[s.cls])}${s.variant === 'sino' ? ' (sino)' : ''}</td>
                    <td class="p-2 whitespace-nowrap">${esc(s.origin)} → ${esc(s.destinationLabel)}</td>
                    <td class="p-2">${fmtRateUp(s)}</td>
                    <td class="p-2 text-right">${s.down || '-'}${s.downVictor !== s.down ? ` <span class="text-xs text-gray-400">(victor ${s.downVictor})</span>` : ''}</td>
                    <td class="p-2 text-xs text-amber-300">${s.flags.map(esc).join('<br>')}</td>
                </tr>`).join('');
            return `
                <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
                    ${importStatCard('อัตราที่อ่านได้', standards.length)}
                    ${importStatCard('แถวที่ควรตรวจในไฟล์', flagged.length)}
                    ${importStatCard('ชีตที่ไม่ใช่ตารางอัตรา', skippedSheets.length)}
                    ${importStatCard('อัตราในระบบตอนนี้', appState.rateStandards.length)}
                </div>
                ${warnings.length ? `<p class="text-yellow-400 mb-3">⚠️ ${warnings.map(esc).join(', ')}</p>` : ''}
                ${skippedSheets.length ? `<p class="text-xs text-gray-400 mb-3">ไม่ได้อ่าน: ${skippedSheets.map(esc).join(', ')} (ไม่มีหัวคอลัมน์ "รวมยอดขาขึ้น/ลิตร")</p>` : ''}
                ${cur ? `<p class="text-xs text-gray-400 mb-3">ฉบับในระบบตอนนี้: ${esc(cur.fileName)} · บันทึกเมื่อ ${esc(cur.importedAt)}</p>` : ''}
                <div class="overflow-x-auto rounded-lg border border-slate-700 mb-4 max-h-[28rem]">
                    <table class="w-full text-sm">
                        <thead class="bg-slate-800 text-gray-400 sticky top-0"><tr>
                            <th class="p-2 text-left">ชีต · แถว</th><th class="p-2 text-left">ประเภทรถ</th><th class="p-2 text-left">เส้นทาง</th><th class="p-2 text-left">ขาขึ้น (ลิตร)</th><th class="p-2 text-right">ขาล่องไพศาล</th><th class="p-2 text-left">ข้อสังเกต</th>
                        </tr></thead>
                        <tbody>${rows}</tbody>
                    </table>
                </div>
                <div class="bg-sky-900/30 border border-sky-700 rounded-lg p-3 mb-4 text-sm text-sky-200">ไฟล์นี้จะเป็น <b>อัตรามาตรฐาน</b> ที่หน้า "ตรวจการเติม" ใช้เทียบ — บันทึกแล้วจะแทนอัตราฉบับเดิมทั้งชุด (ไม่กระทบคิวรถหรือเส้นทาง)</div>
                <button id="import-rates-save-btn" class="bg-primary hover:bg-secondary text-white font-bold py-2 px-5 rounded-lg shadow-md">บันทึกเป็นอัตรามาตรฐาน (${standards.length} อัตรา)</button>
                <div id="import-result" class="mt-4"></div>`;
        };

        const handleImportRatesCommit = async () => {
            if (!importParsedStandards?.standards?.length) return;
            const { standards, fileName } = importParsedStandards;
            if (!confirm(`บันทึก ${standards.length} อัตราจาก "${fileName}" เป็นอัตรามาตรฐาน?\n\nอัตราฉบับเดิมในระบบจะถูกแทนที่ทั้งชุด`)) return;
            const btn = document.getElementById('import-rates-save-btn');
            const resultDiv = document.getElementById('import-result');
            btn.disabled = true; btn.classList.add('opacity-60', 'cursor-not-allowed');
            try {
                const meta = { fileName, importedAt: new Date().toLocaleString('th-TH'), count: standards.length, flagged: standards.filter(s => s.flags.length).length };
                await logisticsService.replaceRateStandards(standards, meta);
                resultDiv.innerHTML = `<p class="text-green-400 font-bold">✓ บันทึกอัตรามาตรฐาน ${standards.length} อัตราแล้ว — ไปที่แท็บ "ตรวจการเติม" เพื่อดูผลตรวจ</p>`;
            } catch (err) {
                console.error('Save rate standards failed:', err);
                resultDiv.innerHTML = `<p class="text-red-400">บันทึกไม่สำเร็จ: ${esc(err.message)}</p>`;
                btn.disabled = false; btn.classList.remove('opacity-60', 'cursor-not-allowed');
            }
        };

        const renderImport = () => {
            const start = appState.templateStartMonth;
            return `
                <div class="max-w-5xl mx-auto space-y-6">
                    ${renderCard(`
                        <h2 class="text-xl font-bold text-white mb-2 flex items-center gap-3">${ICONS.DocumentAdd} แม่แบบนำเข้าเที่ยวรถ (ขาขึ้น / ขาล่อง)</h2>
                        <p class="text-gray-400 text-sm mb-4">คนจัดรถกรอกทุกวันแทนการกดบันทึกทีละเที่ยว: 1 แถว = 1 เที่ยว → อัปโหลดไฟล์ที่ช่องด้านล่าง → ระบบตรวจและแสดงผลตรวจน้ำมันก่อนบันทึกเสมอ</p>
                        <div class="flex flex-wrap items-end gap-3">
                            <button id="tpl-download-btn" class="bg-primary hover:bg-secondary text-white font-bold py-2 px-4 rounded-lg shadow-md flex items-center gap-2">${ICONS.Download} ดาวน์โหลดแม่แบบ</button>
                            <div>
                                <label for="tpl-start-month" class="block text-sm text-gray-400 mb-1">ใช้แม่แบบตั้งแต่เดือน</label>
                                <input type="month" id="tpl-start-month" value="${esc(start)}" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" />
                            </div>
                            <button id="tpl-start-save" class="bg-slate-600 hover:bg-slate-500 text-white font-bold py-2 px-4 rounded-lg">บันทึกเดือนเริ่ม</button>
                        </div>
                        <p class="text-xs mt-2 ${start ? 'text-gray-500' : 'text-yellow-400'}">${start
                            ? `ตั้งแต่เดือนนี้ เที่ยวรถมาจากแม่แบบเท่านั้น — ถ้านำเข้ากระดานคิว/ไฟล์ Transport ของเดือนนี้ขึ้นไป ระบบจะเตือนเพราะอาจนับเที่ยวซ้ำ`
                            : `⚠️ ยังไม่ได้ตั้งเดือนเริ่มใช้แม่แบบ — ต้องตั้งก่อนจึงจะนำเข้าแม่แบบได้`}</p>
                        <div id="tpl-msg" class="mt-2 text-sm"></div>
                    `)}
                    ${renderCard(`
                        <h2 class="text-xl font-bold text-white mb-2">แก้วันที่เที่ยวจากไฟล์ Transport (บันทึกไว้เร็วไป 1 วัน)</h2>
                        <p class="text-gray-400 text-sm mb-4">เที่ยวที่นำเข้าจากไฟล์ Transport ถูกบันทึกวันที่เร็วกว่าใน Excel 1 วัน (พบ 17 พ.ย. 2569) · เลือกไฟล์ Transport ล่าสุด → ตรวจ (ยังไม่บันทึก) → ดาวน์โหลดสำรองข้อมูล → กดแก้ · เที่ยวจากกระดานคิวรายวันและแม่แบบไม่ถูกแตะ</p>
                        <div class="flex flex-wrap items-end gap-3">
                            <input type="file" id="dfix-file" accept=".xlsx,.xls" class="text-sm text-gray-300 file:mr-3 file:py-2 file:px-4 file:rounded-lg file:border-0 file:bg-slate-600 file:text-white file:font-bold" />
                            <button id="dfix-check-btn" class="bg-sky-600 hover:bg-sky-700 text-white font-bold py-2 px-4 rounded-lg">ตรวจ (ยังไม่บันทึก)</button>
                        </div>
                        <div id="dfix-out" class="mt-4 text-sm"></div>
                    `)}
                    ${renderCard(`
                        <h2 class="text-2xl font-bold text-white mb-2 flex items-center gap-3">${ICONS.Download} นำเข้าเส้นทางจาก Excel</h2>
                        <p class="text-gray-400 mb-6">อัปโหลดไฟล์ Excel แล้วเลือกชีต — ระบบตรวจชนิดให้อัตโนมัติ:<br>• <span class="text-accent">แม่แบบนำเข้าเที่ยวรถ</span> → คิวรถ (ขาขึ้น/ขาล่อง) &nbsp;•&nbsp; <span class="text-accent">ฐานคำนวนน้ำมันขาขึ้น</span> → เส้นทาง &nbsp;•&nbsp; <span class="text-accent">คิวขาขึ้น/ขาล่อง, คิวรถร่วม</span> → คิวรถ &nbsp;•&nbsp; <span class="text-accent">ทะเบียน, คนขับ, ฐานคำนวน</span> → รถ/คนขับ/เบี้ยเลี้ยง</p>
                        <div class="rounded-lg border-2 border-dashed border-slate-500 p-6 text-center">
                            <input type="file" id="import-file" accept=".xlsx,.xls" class="block w-full text-sm text-gray-300 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:bg-primary file:text-white file:font-bold hover:file:bg-secondary cursor-pointer" />
                        </div>
                        <div id="import-controls" class="hidden mt-4 flex flex-wrap items-end gap-3">
                            <div>
                                <label class="block text-sm text-gray-400 mb-1">เลือกชีต</label>
                                <select id="import-sheet" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white"></select>
                            </div>
                            <button id="import-parse-btn" class="bg-sky-600 hover:bg-sky-700 text-white font-bold py-2 px-4 rounded-lg shadow-md">อ่าน &amp; ตรวจสอบ</button>
                        </div>
                        <div id="import-preview" class="mt-6"></div>
                    `)}
                </div>
            `;
        };

        const importStatCard = (label, val) =>
            `<div class="bg-slate-800 rounded-lg p-3 text-center"><div class="text-2xl font-bold text-accent">${(val || 0).toLocaleString()}</div><div class="text-xs text-gray-400">${label}</div></div>`;

        const renderImportPreview = (result) => {
            const { routes, stats, warnings } = result;
            if (!routes.length) {
                return `<div class="text-red-400">ไม่พบเส้นทางที่นำเข้าได้${warnings.length ? ' — ' + warnings.join(', ') : ''}</div>`;
            }
            const sample = routes.slice(0, 12).map(r => `
                <tr class="border-b border-slate-700">
                    <td class="p-2 whitespace-nowrap">${r.vehicleType}</td>
                    <td class="p-2 whitespace-nowrap">${r.origin} → ${r.destination}</td>
                    <td class="p-2 text-center whitespace-nowrap">${r.direction || ''}</td>
                    <td class="p-2 text-center whitespace-nowrap">${r.crossingFee || ''}</td>
                    <td class="p-2 text-center whitespace-nowrap">${r.waitingFee || ''}</td>
                    <td class="p-2 text-right text-accent whitespace-nowrap">${r.totalLiters ? r.totalLiters.toLocaleString() + ' L' : ''} ${r.totalBaht ? r.totalBaht.toLocaleString() + ' B' : ''}</td>
                </tr>`).join('');
            return `
                <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
                    ${importStatCard('เส้นทางสรุปได้', stats.total)}
                    ${importStatCard('มีค่าน้ำมัน', stats.withFuel)}
                    ${importStatCard('ไม่มีน้ำมัน (ค่าข้าม/ค้างเท่านั้น)', stats.zeroFuel)}
                    ${importStatCard('ค่าน้ำมันต่างกัน→เลือก mode', stats.variedRoutes)}
                </div>
                <p class="text-xs text-gray-500 mb-3">จาก ${(stats.rowsUsed || 0).toLocaleString()} แถวดิบ · ข้ามแถวไม่ครบ ${stats.skipped} แถว · พบคอลัมน์ปั๊ม ${stats.stationColumns} คอลัมน์</p>
                ${warnings.length ? `<p class="text-yellow-400 mb-3">⚠️ ${warnings.join(', ')}</p>` : ''}
                <div class="overflow-x-auto rounded-lg border border-slate-700 mb-2">
                    <table class="w-full text-sm">
                        <thead class="bg-slate-800 text-gray-400"><tr>
                            <th class="p-2 text-left">ประเภทรถ</th><th class="p-2 text-left">เส้นทาง</th><th class="p-2">ทิศทาง</th><th class="p-2">ค่าข้าม</th><th class="p-2">ค่าค้าง</th><th class="p-2 text-right">น้ำมัน</th>
                        </tr></thead>
                        <tbody>${sample}</tbody>
                    </table>
                </div>
                <p class="text-xs text-gray-500 mb-4">แสดงตัวอย่าง 12 จาก ${stats.total.toLocaleString()} เส้นทาง</p>
                <div class="bg-yellow-900/30 border border-yellow-700 rounded-lg p-3 mb-4 text-sm text-yellow-200">💡 แนะนำ: กด "สำรองข้อมูล" ในหน้าจัดการเส้นทางก่อนนำเข้า เพื่อกันพลาด</div>
                <div class="flex flex-wrap gap-3">
                    <button id="import-replace-btn" class="bg-error hover:bg-red-600 text-white font-bold py-2 px-5 rounded-lg shadow-md">แทนที่เส้นทางทั้งหมด (${stats.total.toLocaleString()})</button>
                    <button id="import-merge-btn" class="bg-success hover:bg-green-600 text-white font-bold py-2 px-5 rounded-lg shadow-md">เพิ่มเข้าไป (รวมกับของเดิม)</button>
                </div>
                <div id="import-result" class="mt-4"></div>
            `;
        };

        const handleImportFile = async (e) => {
            const file = e.target.files?.[0];
            if (!file) return;
            const preview = document.getElementById('import-preview');
            try {
                const buf = await file.arrayBuffer();
                importWorkbook = XLSX.read(buf, { type: 'array', cellDates: true });
                const sel = document.getElementById('import-sheet');
                sel.innerHTML = importWorkbook.SheetNames.map(n => `<option value="${n}">${n}</option>`).join('');
                const def = importWorkbook.SheetNames.find(n => n === TEMPLATE_SHEET) || importWorkbook.SheetNames.find(n => n.includes('ฐานคำนวนน้ำมันขาขึ้น')) || importWorkbook.SheetNames.find(n => n.includes('ฐานคำนวน'));
                if (def) sel.value = def;
                document.getElementById('import-controls').classList.remove('hidden');
                preview.innerHTML = '<p class="text-gray-400">เลือกชีตแล้วกด "อ่าน &amp; ตรวจสอบ"</p>';
            } catch (err) {
                console.error('Read file failed:', err);
                preview.innerHTML = `<p class="text-red-400">อ่านไฟล์ไม่สำเร็จ: ${err.message}</p>`;
            }
        };

        const handleImportParse = () => {
            const preview = document.getElementById('import-preview');
            if (!importWorkbook) { preview.innerHTML = '<p class="text-red-400">ยังไม่ได้เลือกไฟล์</p>'; return; }
            const sheetName = document.getElementById('import-sheet').value;
            const ws = importWorkbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: '' });
            importParsedRoutes = null;
            importParsedTrips = null;
            importParsedStandards = null;
            importParsedTemplate = null;
            importParsedAliases = null;

            // Excel trip template: read its own data sheet whichever sheet is selected (the help/lists sheets are not data).
            if (importWorkbook.Sheets[TEMPLATE_SHEET] || findTemplateHeader(rows) >= 0) {
                handleTemplateParse(importWorkbook.Sheets[TEMPLATE_SHEET] ? TEMPLATE_SHEET : sheetName);
                return;
            }

            // Confirmed driver-name proposal file (ชื่อในกระดานคิว → ชื่อเต็ม) → settings/fuelAudit.driverAliases.
            if (rows.slice(0, 5).some(r => { const h = (r || []).map(c => String(c).trim()); return h.includes('ชื่อในกระดานคิว') && h.includes('ยืนยัน'); })) {
                const res = parseDriverAliasSheet(rows, appState.drivers.map(d => d.name));
                importParsedAliases = res.error ? null : res;
                preview.innerHTML = renderDriverAliasPreview(res);
                document.getElementById('alias-save-btn')?.addEventListener('click', handleDriverAliasSave);
                return;
            }

            // Official rate file ("อัตราเติมน้ำมันประจำ ปั๊ม ต่างๆ"): when the selected sheet is a rate table, read every
            // sheet of the workbook. blankrows:true + the used-range start keep real Excel row numbers for each rate.
            const isRateSheet = rows.some(r => (r || []).some(c => String(c).trim() === 'รวมยอดขาขึ้น/ลิตร'));
            const rate = isRateSheet ? parseRateSheet(Object.fromEntries(importWorkbook.SheetNames.map(n => {
                const s = importWorkbook.Sheets[n];
                return [n, { rows: XLSX.utils.sheet_to_json(s, { header: 1, blankrows: true, defval: '' }), startRow: s['!ref'] ? XLSX.utils.decode_range(s['!ref']).s.r : 0 }];
            }))) : { standards: [] };
            if (rate.standards.length) {
                importParsedStandards = { ...rate, fileName: document.getElementById('import-file')?.files?.[0]?.name || '' };
                preview.innerHTML = renderRateImportPreview(importParsedStandards);
                document.getElementById('import-rates-save-btn')?.addEventListener('click', handleImportRatesCommit);
                return;
            }

            const kind = detectSheetType(rows);
            if (kind === 'fuelbase') {
                const result = parseFuelBaseRows(rows);
                importParsedRoutes = result.routes;
                preview.innerHTML = renderImportPreview(result);
                document.getElementById('import-replace-btn')?.addEventListener('click', () => handleImportCommit('replace'));
                document.getElementById('import-merge-btn')?.addEventListener('click', () => handleImportCommit('merge'));
            } else if (kind === 'uphill' || kind === 'downhill' || kind === 'partnerqueue') {
                const result = kind === 'uphill' ? parseUphillQueue(rows) : kind === 'downhill' ? parseDownhillQueue(rows) : parsePartnerQueue(rows);
                importParsedTrips = result.trips;
                preview.innerHTML = (kind === 'downhill' ? '<div id="coll-section" class="mb-6"><p class="text-sky-400">กำลังอ่านลิตรขาล่องจากช่อง "คิวรถล่าสุด"...</p></div>' : '') + renderTripsImportPreview(result, kind);
                document.getElementById('import-trips-replace-btn')?.addEventListener('click', () => handleImportTripsCommit('replace'));
                document.getElementById('import-trips-merge-btn')?.addEventListener('click', () => handleImportTripsCommit('merge'));
                if (kind === 'downhill') prepareColumnL(result);
            } else if (kind === 'drivers' || kind === 'ratetable' || kind === 'vehicles') {
                const result = kind === 'drivers' ? parseDrivers(rows) : kind === 'vehicles' ? parseVehicles(rows) : parseRateTable(rows);
                const items = kind === 'drivers' ? result.drivers : kind === 'vehicles' ? result.vehicles : result.rates;
                importParsedMaster = { kind, items };
                preview.innerHTML = renderMasterImportPreview(result, kind);
                document.getElementById('import-master-replace-btn')?.addEventListener('click', () => handleImportMasterCommit('replace'));
                document.getElementById('import-master-merge-btn')?.addEventListener('click', () => handleImportMasterCommit('merge'));
            } else if (kind === 'daily') {
                // Parse ALL daily-board sheets in the workbook (คิวขาขึ้น ปี 2026.xlsx) — these carry the real fuel.
                const allTrips = [];
                let skipped = 0, sheets = 0;
                const failedSheets = [];
                for (const name of importWorkbook.SheetNames) {
                    if (!/^\d{2}\.\d{2}\.\d{2}$/.test(name)) continue;
                    // Guard each sheet independently: one malformed sheet must not abort the whole
                    // import (files can have ~16k rows across dozens of sheets).
                    try {
                        const rws = XLSX.utils.sheet_to_json(importWorkbook.Sheets[name], { header: 1, blankrows: false, defval: '' });
                        const res = parseDailyRows(name, rws);
                        allTrips.push(...res.trips); skipped += res.skipped; sheets++;
                    } catch (err) {
                        console.error(`Parse daily sheet "${name}" failed:`, err);
                        failedSheets.push(name);
                    }
                }
                importParsedTrips = allTrips;
                const dates = allTrips.map(t => t.date).filter(Boolean).sort();
                const months = new Set(allTrips.map(t => t.yearMonth));
                const stats = { total: allTrips.length, withFuel: allTrips.filter(t => t.totalLiters > 0 || t.totalBaht > 0).length, months: months.size, minDate: dates[0] || '', maxDate: dates[dates.length - 1] || '', skipped };
                const dailyWarnings = [`อ่านจาก ${sheets} ชีตรายวัน — คิวขาขึ้นที่มีข้อมูลน้ำมันจริง (แกะจากคอลัมน์ "ปั๊มน้ำมัน")`];
                if (failedSheets.length) dailyWarnings.push(`⚠️ อ่านไม่สำเร็จ ${failedSheets.length} ชีต: ${failedSheets.join(', ')}`);
                preview.innerHTML = renderTripsImportPreview({ trips: allTrips, stats, warnings: dailyWarnings }, 'daily');
                document.getElementById('import-trips-replace-btn')?.addEventListener('click', () => handleImportTripsCommit('replace'));
                document.getElementById('import-trips-merge-btn')?.addEventListener('click', () => handleImportTripsCommit('merge'));
            } else {
                preview.innerHTML = `
                    <div class="bg-red-900/30 border border-red-700 rounded-lg p-4 text-red-300">
                        <p class="font-bold mb-2">ไม่รู้จักรูปแบบชีตนี้ (เลือกไฟล์หรือชีตผิด)</p>
                        <p class="text-sm">ใช้ไฟล์ <b class="text-white">Transport 2026.xlsx</b> แล้วเลือกชีต: <b>ฐานคำนวนน้ำมันขาขึ้น</b> (เส้นทาง), <b>คิวขาขึ้น</b> หรือ <b>คิวขาล่อง</b> (คิวรถ)</p>
                    </div>`;
            }
        };

        const renderTripsImportPreview = (result, kind) => {
            const { trips, stats, warnings } = result;
            const label = kind === 'uphill' ? 'คิวขาขึ้น' : kind === 'downhill' ? 'คิวขาล่อง' : kind === 'daily' ? 'คิวรายวัน (มีน้ำมัน)' : 'คิวรถร่วม';
            if (!trips.length) return `<div class="text-red-400">ไม่พบคิวรถที่นำเข้าได้${warnings.length ? ' — ' + warnings.join(', ') : ''}</div>`;
            const sample = trips.slice(0, 12).map(t => `
                <tr class="border-b border-slate-700">
                    <td class="p-2 whitespace-nowrap">${t.date}</td>
                    <td class="p-2 whitespace-nowrap">${t.driver}</td>
                    <td class="p-2 font-mono whitespace-nowrap">${t.plate}</td>
                    <td class="p-2 whitespace-nowrap">${t.origin} → ${t.destination}</td>
                    <td class="p-2 text-right whitespace-nowrap text-accent">${t.totalLiters ? t.totalLiters.toLocaleString() + ' L' : ''}</td>
                    <td class="p-2 text-right whitespace-nowrap">${t.revenue ? t.revenue.toLocaleString() : ''}</td>
                    <td class="p-2 text-right whitespace-nowrap text-green-400">${t.netTransfer ? t.netTransfer.toLocaleString() : ''}</td>
                </tr>`).join('');
            return `
                <div class="mb-3 text-lg font-bold text-white">ชีตประเภท: <span class="text-accent">${label}</span></div>
                <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
                    ${importStatCard('คิวรถทั้งหมด', stats.total)}
                    ${importStatCard('มีข้อมูลน้ำมัน', stats.withFuel)}
                    ${importStatCard('จำนวนเดือน', stats.months)}
                    ${importStatCard('ข้ามแถวไม่ครบ', stats.skipped)}
                </div>
                <p class="text-xs text-gray-500 mb-3">ช่วงวันที่: ${stats.minDate} ถึง ${stats.maxDate} · เก็บแยกตามเดือน (trips/YYYY-MM)</p>
                ${warnings.length ? `<p class="text-yellow-400 mb-3">⚠️ ${warnings.join(', ')}</p>` : ''}
                ${templateOverlapNote(trips) ? `<p class="text-red-300 bg-red-900/30 border border-red-700 rounded-lg p-3 mb-3">⚠️ ${esc(templateOverlapNote(trips))}</p>` : ''}
                <div class="overflow-x-auto rounded-lg border border-slate-700 mb-2">
                    <table class="w-full text-sm"><thead class="bg-slate-800 text-gray-400"><tr>
                        <th class="p-2 text-left">วันที่</th><th class="p-2 text-left">คนขับ</th><th class="p-2 text-left">ทะเบียน</th><th class="p-2 text-left">เส้นทาง</th><th class="p-2 text-right">น้ำมัน</th><th class="p-2 text-right">รายรับ</th><th class="p-2 text-right">สุทธิโอน</th>
                    </tr></thead><tbody>${sample}</tbody></table>
                </div>
                <p class="text-xs text-gray-500 mb-4">แสดงตัวอย่าง 12 จาก ${stats.total.toLocaleString()} คิว</p>
                <div class="bg-yellow-900/30 border border-yellow-700 rounded-lg p-3 mb-4 text-sm text-yellow-200">💡 "แทนที่ทั้งหมด" จะลบคิวรถทุกเดือนที่มีอยู่ แล้วใส่ชุดนี้แทน · "เพิ่มเข้าไป" ค่อยๆ เติมโดยไม่ลบ</div>
                <div class="flex flex-wrap gap-3">
                    <button id="import-trips-replace-btn" class="bg-error hover:bg-red-600 text-white font-bold py-2 px-5 rounded-lg shadow-md">แทนที่คิวรถทั้งหมด (${stats.total.toLocaleString()})</button>
                    <button id="import-trips-merge-btn" class="bg-success hover:bg-green-600 text-white font-bold py-2 px-5 rounded-lg shadow-md">เพิ่มเข้าไป (รวมกับของเดิม)</button>
                </div>
                <div id="import-result" class="mt-4"></div>
            `;
        };

        const handleImportTripsCommit = async (mode) => {
            if (!importParsedTrips || !importParsedTrips.length) return;
            const n = importParsedTrips.length;
            const msg = mode === 'replace'
                ? `แทนที่คิวรถทั้งหมดด้วย ${n.toLocaleString()} คิวจาก Excel?\n\n⚠️ คิวรถเดิมทุกเดือนจะถูกลบ (ควรสำรองข้อมูลก่อน)${appState.templateStartMonth ? '\n⚠️ รวมถึงเที่ยวที่คนจัดรถนำเข้าจากแม่แบบด้วย — ถ้าไม่แน่ใจ ให้กด "เพิ่มเข้าไป" แทน' : ''}`
                : `เพิ่ม ${n.toLocaleString()} คิวเข้าไป (ไม่ลบของเดิม)?`;
            const overlap = templateOverlapNote(importParsedTrips);
            if (!confirm(overlap ? `${msg}\n\n⚠️ ${overlap}` : msg)) return;
            const resultDiv = document.getElementById('import-result');
            const buttons = document.querySelectorAll('#import-trips-replace-btn, #import-trips-merge-btn');
            buttons.forEach(b => { b.disabled = true; b.classList.add('opacity-60', 'cursor-not-allowed'); });
            resultDiv.innerHTML = '<p class="text-sky-400">กำลังนำเข้า... (อาจใช้เวลาสักครู่)</p>';
            try {
                if (mode === 'replace') await logisticsService.replaceAllTrips(importParsedTrips);
                else await logisticsService.addTripsBatch(importParsedTrips);
                resultDiv.innerHTML = `<p class="text-green-400 font-bold">✓ นำเข้าสำเร็จ ${n.toLocaleString()} คิว! ไปที่หน้า "คิวรถ" แล้วเลือกเดือนเพื่อดูข้อมูล</p>`;
            } catch (err) {
                console.error('Import trips failed:', err);
                resultDiv.innerHTML = `<p class="text-red-400">นำเข้าไม่สำเร็จ: ${err.message}</p>`;
                buttons.forEach(b => { b.disabled = false; b.classList.remove('opacity-60', 'cursor-not-allowed'); });
            }
        };

        const handleImportCommit = async (mode) => {
            if (!importParsedRoutes || !importParsedRoutes.length) return;
            const n = importParsedRoutes.length;
            const msg = mode === 'replace'
                ? `แทนที่เส้นทางทั้งหมดด้วย ${n} เส้นทางจาก Excel?\n\n⚠️ เส้นทางเดิมทั้งหมดจะถูกลบ (ควรสำรองข้อมูลก่อน)`
                : `เพิ่ม ${n} เส้นทางเข้าไปในระบบ (ไม่ลบของเดิม)?`;
            if (!confirm(msg)) return;
            const resultDiv = document.getElementById('import-result');
            const buttons = document.querySelectorAll('#import-replace-btn, #import-merge-btn');
            buttons.forEach(b => { b.disabled = true; b.classList.add('opacity-60', 'cursor-not-allowed'); });
            resultDiv.innerHTML = '<p class="text-sky-400">กำลังนำเข้า...</p>';
            try {
                if (mode === 'replace') await logisticsService.replaceAllRoutes(importParsedRoutes);
                else await logisticsService.addRoutesBatch(importParsedRoutes);
                resultDiv.innerHTML = `<p class="text-green-400 font-bold">✓ นำเข้าสำเร็จ ${n} เส้นทาง! ไปที่หน้า "จัดการเส้นทาง" เพื่อดูข้อมูล</p>`;
            } catch (err) {
                console.error('Import failed:', err);
                resultDiv.innerHTML = `<p class="text-red-400">นำเข้าไม่สำเร็จ: ${err.message}</p>`;
                buttons.forEach(b => { b.disabled = false; b.classList.remove('opacity-60', 'cursor-not-allowed'); });
            }
        };

        const renderMasterImportPreview = (result, kind) => {
            const label = kind === 'drivers' ? 'คนขับ' : kind === 'vehicles' ? 'ทะเบียนรถ' : 'ตารางเบี้ยเลี้ยง';
            const items = kind === 'drivers' ? result.drivers : kind === 'vehicles' ? result.vehicles : result.rates;
            if (!items.length) return `<div class="text-red-400">ไม่พบข้อมูล${result.warnings.length ? ' — ' + result.warnings.join(', ') : ''}</div>`;
            let head, sample, statCards, allowMerge = true;
            if (kind === 'drivers') {
                head = '<tr><th class="p-2 text-left">ชื่อ</th><th class="p-2 text-left">เลขบัญชี</th></tr>';
                sample = items.slice(0, 15).map(d => `<tr class="border-b border-slate-700"><td class="p-2">${d.name}</td><td class="p-2 font-mono">${d.bankAccount || '<span class="text-gray-600">—</span>'}</td></tr>`).join('');
                statCards = `${importStatCard('จำนวนทั้งหมด', result.stats.total)}${importStatCard('มีเลขบัญชี', result.stats.withBank)}${importStatCard('ข้ามแถวไม่ครบ', result.stats.skipped)}`;
            } else if (kind === 'vehicles') {
                head = '<tr><th class="p-2 text-left">ทะเบียน</th><th class="p-2 text-left">ประเภทรถ</th><th class="p-2 text-left">ชนิด</th><th class="p-2 text-left">บริษัท</th></tr>';
                sample = items.slice(0, 15).map(v => `<tr class="border-b border-slate-700"><td class="p-2 font-mono">${v.registrationNumber}</td><td class="p-2">${v.vehicleType || ''}</td><td class="p-2">${v.driveType || ''}</td><td class="p-2">${v.company || ''}</td></tr>`).join('');
                statCards = `${importStatCard('จำนวนทั้งหมด', result.stats.total)}${importStatCard('รถบริษัท', result.stats.own)}${importStatCard('รถร่วม', result.stats.partner)}`;
            } else {
                head = '<tr><th class="p-2 text-left">ประเภทวิ่ง</th><th class="p-2 text-left">ประเภทรถ</th><th class="p-2 text-right">เบี้ยเลี้ยง</th><th class="p-2 text-right">หักเงินสะสม</th></tr>';
                sample = items.slice(0, 25).map(r => `<tr class="border-b border-slate-700"><td class="p-2">${r.runType}</td><td class="p-2">${r.vehicleType}</td><td class="p-2 text-right">${r.allowance.toLocaleString()}</td><td class="p-2 text-right">${r.deduction.toLocaleString()}</td></tr>`).join('');
                statCards = `${importStatCard('จำนวนทั้งหมด', result.stats.total)}${importStatCard('ตัดซ้ำออก', result.stats.duplicates)}${importStatCard('ข้ามแถวไม่ครบ', result.stats.skipped)}`;
                allowMerge = false;
            }
            return `
                <div class="mb-3 text-lg font-bold text-white">ชีตประเภท: <span class="text-accent">${label}</span></div>
                <div class="grid grid-cols-2 md:grid-cols-3 gap-3 mb-4">${statCards}</div>
                ${kind === 'drivers' ? '<div class="bg-red-900/30 border border-red-700 rounded-lg p-3 mb-4 text-sm text-red-200">🔒 มีเลขบัญชีธนาคาร — Firebase ยังเปิดสาธารณะ ควรตั้ง Security Rules โดยเร็ว</div>' : ''}
                <div class="overflow-x-auto rounded-lg border border-slate-700 mb-4">
                    <table class="w-full text-sm"><thead class="bg-slate-800 text-gray-400">${head}</thead><tbody>${sample}</tbody></table>
                </div>
                <div class="flex flex-wrap gap-3">
                    <button id="import-master-replace-btn" class="bg-error hover:bg-red-600 text-white font-bold py-2 px-5 rounded-lg shadow-md">แทนที่ทั้งหมด (${result.stats.total})</button>
                    ${allowMerge ? `<button id="import-master-merge-btn" class="bg-success hover:bg-green-600 text-white font-bold py-2 px-5 rounded-lg shadow-md">เพิ่มเข้าไป (รวมกับของเดิม)</button>` : ''}
                </div>
                <div id="import-result" class="mt-4"></div>
            `;
        };

        const handleImportMasterCommit = async (mode) => {
            if (!importParsedMaster || !importParsedMaster.items.length) return;
            const { kind, items } = importParsedMaster;
            const n = items.length;
            const label = kind === 'drivers' ? 'คนขับ' : kind === 'vehicles' ? 'ทะเบียนรถ' : 'อัตราเบี้ยเลี้ยง';
            const msg = mode === 'replace' ? `แทนที่ข้อมูล${label}ทั้งหมดด้วย ${n} รายการ?` : `เพิ่ม ${n} รายการ${label}?`;
            if (!confirm(msg)) return;
            const resultDiv = document.getElementById('import-result');
            const buttons = document.querySelectorAll('#import-master-replace-btn, #import-master-merge-btn');
            buttons.forEach(b => { b.disabled = true; b.classList.add('opacity-60', 'cursor-not-allowed'); });
            resultDiv.innerHTML = '<p class="text-sky-400">กำลังนำเข้า...</p>';
            try {
                if (kind === 'drivers') {
                    if (mode === 'replace') await logisticsService.replaceAllDrivers(items);
                    else await logisticsService.addDriversBatch(items);
                } else if (kind === 'vehicles') {
                    // The Excel ทะเบียน sheet has no รุ่นรถ column — keep the sino/victor set in the app so a re-import
                    // does not silently switch those trucks back to the general fuel rate.
                    const modelOf = {};
                    logisticsService.getVehicles().forEach(v => { if (v.model) modelOf[String(v.registrationNumber || '').replace(/\s+/g, '')] = v.model; });
                    const withModel = items.map(v => {
                        const m = modelOf[String(v.registrationNumber || '').replace(/\s+/g, '')];
                        return m && !v.model ? { ...v, model: m } : v;
                    });
                    if (mode === 'replace') await logisticsService.replaceAllVehicles(withModel);
                    else await logisticsService.addVehiclesBatch(withModel);
                } else {
                    await logisticsService.replaceAllRates(items); // rate table is always replace
                }
                resultDiv.innerHTML = `<p class="text-green-400 font-bold">✓ นำเข้าสำเร็จ ${n} รายการ!</p>`;
            } catch (err) {
                console.error('Import master failed:', err);
                resultDiv.innerHTML = `<p class="text-red-400">นำเข้าไม่สำเร็จ: ${err.message}</p>`;
                buttons.forEach(b => { b.disabled = false; b.classList.remove('opacity-60', 'cursor-not-allowed'); });
            }
        };

        // ---- Driver name aliases (confirmed proposal file) ----
        const renderDriverAliasPreview = (res) => {
            if (res.error) return `<p class="text-red-400">${esc(res.error)}</p>`;
            const cur = Object.keys(appState.auditState.driverAliases || {}).length;
            const rows = res.aliases.map(a => `<tr class="border-b border-slate-700"><td class="p-2">${esc(a.from)}</td><td class="p-2">→</td><td class="p-2">${esc(a.to)}</td></tr>`).join('');
            return `
                <div class="mb-3 text-lg font-bold text-white">ไฟล์ประเภท: <span class="text-accent">จับคู่ชื่อคนขับ (ชื่อในกระดานคิว → ชื่อเต็ม)</span></div>
                <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
                    ${importStatCard('จับคู่ที่ยืนยันแล้ว', res.aliases.length)}
                    ${importStatCard('ไม่ได้ยืนยัน (ข้าม)', res.skipped.length)}
                    ${importStatCard('ต้องแก้ในไฟล์ (ไม่บันทึก)', res.conflicts.length)}
                    ${importStatCard('ที่ใช้อยู่ในระบบตอนนี้', cur)}
                </div>
                ${res.conflicts.length ? `<div class="text-yellow-300 text-sm mb-3">⚠️ ไม่บันทึก ${res.conflicts.length} รายการ — แก้ในไฟล์แล้วนำเข้าใหม่:<ul class="mt-1 ml-4 list-disc">${res.conflicts.map(c => `<li>${esc(c.from)} → ${esc(c.to)} · ${esc(c.reason)}</li>`).join('')}</ul></div>` : ''}
                <div class="bg-sky-900/30 border border-sky-700 rounded-lg p-3 mb-4 text-sm text-sky-200">ใช้ตอนตรวจการเติมเท่านั้น: รวมยอดยกลิตรและสรุปรายคนขับของชื่อในกระดานคิวกับชื่อเต็มเป็นคนเดียวกัน · <b class="text-white">ไม่แก้ชื่อในเที่ยวรถเดิม</b> · บันทึกแล้วแทนชุดเดิมทั้งชุด</div>
                <div class="overflow-x-auto max-h-80 overflow-y-auto rounded-lg border border-slate-700 mb-4"><table class="w-full text-sm"><tbody>${rows}</tbody></table></div>
                ${res.aliases.length || cur
                    ? `<button id="alias-save-btn" class="${res.aliases.length ? 'bg-primary hover:bg-secondary' : 'bg-error hover:bg-red-600'} text-white font-bold py-2 px-5 rounded-lg shadow-md">${res.aliases.length ? `บันทึกการจับคู่ชื่อ (${res.aliases.length})` : `ล้างการจับคู่ชื่อที่ใช้อยู่ (${cur})`}</button>`
                    : '<p class="text-gray-400">ยังไม่มีชื่อที่ยืนยันในไฟล์</p>'}
                <div id="import-result" class="mt-4"></div>`;
        };
        const handleDriverAliasSave = async () => {
            const res = importParsedAliases;
            if (!res) return;
            const n = res.aliases.length;
            if (!confirm(n ? `บันทึกการจับคู่ชื่อคนขับ ${n} ชื่อ?\n\nใช้ตอนตรวจการเติมเท่านั้น ไม่แก้ชื่อในเที่ยวรถเดิม · แทนชุดเดิมทั้งชุด` : 'ล้างการจับคู่ชื่อคนขับทั้งหมด?\n\nหน้าตรวจการเติมจะแยกชื่อในกระดานกับชื่อเต็มเหมือนเดิม')) return;
            const btn = document.getElementById('alias-save-btn');
            const out = document.getElementById('import-result');
            btn.disabled = true; btn.classList.add('opacity-60', 'cursor-not-allowed');
            try {
                // false = none (Firebase drops an empty array). The settings listener updates appState.
                await logisticsService.saveAuditSettings({ driverAliases: n ? res.aliases : false });
                out.innerHTML = n ? `<p class="text-green-400 font-bold">✓ บันทึกการจับคู่ชื่อแล้ว ${n} ชื่อ — ดูผลที่หน้า "ตรวจการเติม"</p>` : '<p class="text-green-400 font-bold">✓ ล้างการจับคู่ชื่อแล้ว</p>';
            } catch (err) {
                console.error('Save driver aliases failed:', err);
                out.innerHTML = `<p class="text-red-400">บันทึกไม่สำเร็จ: ${esc(err.message)}</p>`;
                btn.disabled = false; btn.classList.remove('opacity-60', 'cursor-not-allowed');
            }
        };

        // ---- Excel trip template (แม่แบบนำเข้าเที่ยวรถ) ----
        const tplMsg = (html) => { const el = document.getElementById('tpl-msg'); if (el) el.innerHTML = html; };

        const handleTemplateDownload = async () => {
            const btn = document.getElementById('tpl-download-btn');
            if (btn) btn.disabled = true;
            tplMsg('<span class="text-sky-400">กำลังสร้างไฟล์แม่แบบ...</span>');
            try {
                const mod = await import('exceljs'); // loaded only when needed (large library)
                const ExcelJS = mod.default || mod;
                const lists = buildTemplateLists({
                    drivers: appState.drivers, vehicles: appState.vehicles, rateStandards: appState.rateStandards,
                    routes: appState.routes, runTypes: logisticsService.getRateTable().map(r => r.runType),
                });
                const buf = await generateTemplate(ExcelJS, lists).xlsx.writeBuffer();
                const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
                const a = document.createElement('a');
                a.href = url; a.download = `แม่แบบนำเข้าเที่ยวรถ_${localDateStr()}.xlsx`;
                document.body.appendChild(a); a.click(); a.remove();
                setTimeout(() => URL.revokeObjectURL(url), 1000);
                tplMsg(`<span class="text-green-400">✓ ดาวน์โหลดแล้ว — รายการให้เลือก: คนขับ ${lists.drivers.length} · ทะเบียนหัว ${lists.heads.length} · หาง ${lists.trailers.length} · ปลายทาง ${lists.destinations.length}</span>`);
            } catch (err) {
                console.error('Template download failed:', err);
                tplMsg(`<span class="text-red-400">สร้างแม่แบบไม่สำเร็จ: ${esc(err.message)}</span>`);
            } finally { if (btn) btn.disabled = false; }
        };

        const handleTemplateStartSave = async () => {
            const v = (document.getElementById('tpl-start-month')?.value || '').trim();
            if (!/^\d{4}-\d{2}$/.test(v)) { tplMsg('<span class="text-red-400">เลือกเดือนก่อน</span>'); return; }
            const [y, m] = v.split('-').map(Number);
            if (!confirm(`ตั้งให้ใช้แม่แบบตั้งแต่เดือน ${m}/${y + 543}?\n\nตั้งแต่เดือนนี้ เที่ยวรถควรมาจากแม่แบบเท่านั้น และแม่แบบจะไม่รับเที่ยวที่ก่อนเดือนนี้`)) return;
            try {
                await logisticsService.saveTemplateSettings({ startMonth: v });
                appState.templateStartMonth = v;
                tplMsg(`<span class="text-green-400">✓ บันทึกแล้ว: ใช้แม่แบบตั้งแต่เดือน ${m}/${y + 543}</span>`);
            } catch (err) {
                tplMsg(`<span class="text-red-400">บันทึกไม่สำเร็จ: ${esc(err.message)}</span>`);
            }
        };

        const templateDupSource = (d) => d?.source === 'file' ? `แถว ${d.row} ในไฟล์นี้` : d?.source === 'template' ? 'แม่แบบที่นำเข้าไปแล้ว' : d?.source === 'transport' ? 'ไฟล์ Transport' : d?.source === 'daily' ? 'กระดานคิว' : 'ระบบ';
        const TH_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
        // '2026-11-17' → '17 พ.ย. 2569' — so a day/month swapped by the PC's date setting is easy to spot.
        const thaiDate = (d) => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d || ''); return m ? `${+m[3]} ${TH_MONTHS[+m[2] - 1]} ${+m[1] + 543}` : ''; };

        const renderTemplatePreview = ({ rows }, ctx) => {
            const tol = appState.auditState.tolerance || 0;
            const count = (f) => rows.filter(f).length;
            const nErr = count(p => p.status === 'error'), nNew = count(p => p.importState === 'new');
            const nAlready = count(p => p.importState === 'already'), nChanged = count(p => p.changed), nDup = count(p => p.importState === 'possible-dup');
            const importCell = (p) => {
                if (p.status === 'error') return '<span class="text-red-400">ไม่บันทึก</span>';
                if (p.importState === 'new') return '<span class="text-green-400">ใหม่</span>';
                if (p.importState === 'already') return `<span class="text-gray-400">นำเข้าแล้ว</span>${p.changed ? '<br><span class="text-amber-300 text-xs">ตัวเลขต่างจากที่บันทึกไว้ — ไม่ทับ (แก้ที่หน้าคิวรถ)</span>' : ''}`;
                return `<label class="flex items-start gap-2 cursor-pointer"><input type="checkbox" class="tpl-dup-cb w-4 h-4 mt-0.5" data-key="${esc(p.key)}">
                    <span><span class="text-amber-300">อาจซ้ำ</span> <span class="text-xs text-gray-400">(ติ๊กถ้าเป็นคนละเที่ยว)</span><br><span class="text-xs text-gray-400">มีรถคันนี้ วันเดียวกัน ทิศเดียวกันแล้ว (จาก${esc(templateDupSource(p.dupOf))}${p.dupOf?.destination ? ` → ${esc(p.dupOf.destination)}` : ''})</span></span></label>`;
            };
            const auditCell = (p) => {
                if (p.status === 'error' || !p.trip.totalLiters && !p.trip.totalBaht) return '<span class="text-gray-600">—</span>';
                const r = auditTrip(p.trip, ctx, tol);
                if (!r) return '<span class="text-gray-600">—</span>';
                const why = [r.reason, ...(r.flags || [])].filter(Boolean).join(' · ');
                const diff = r.status === 'no-rate' ? '' : ` ${r.diff > 0 ? '+' : ''}${r.diff}`;
                return `<span title="${esc(why)}" class="px-2 py-0.5 rounded text-xs text-white whitespace-nowrap ${AUDIT_STATUS[r.status].cls}">${AUDIT_STATUS[r.status].label}${diff}</span>`;
            };
            const body = rows.slice(0, 3000).map(p => {
                const t = p.trip;
                const msgs = [...p.errors.map(e => `<div class="text-red-400">✗ ${esc(e)}</div>`), ...p.warnings.map(w => `<div class="text-yellow-300">⚠ ${esc(w)}</div>`)].join('');
                return `<tr class="border-b border-slate-700 align-top">
                    <td class="p-2 text-gray-400">${p.excelRow}</td>
                    <td class="p-2 whitespace-nowrap">${esc(thaiDate(t.date))}</td>
                    <td class="p-2 whitespace-nowrap">${esc(t.direction)}</td>
                    <td class="p-2 font-mono whitespace-nowrap">${esc(t.plate)}${t.trailerDrop && t.trailerDrop !== 'ไม่ตัด' ? `<div class="text-xs text-sky-300 font-sans">ตัดหาง ${esc(t.trailerDrop)}</div>` : ''}</td>
                    <td class="p-2 whitespace-nowrap">${esc(t.driver)}</td>
                    <td class="p-2">${esc(t.origin)}${t.origin || t.destination ? ' → ' : ''}${esc(t.destination)}</td>
                    <td class="p-2 text-right whitespace-nowrap text-accent">${t.totalLiters ? t.totalLiters.toLocaleString() + ' L' : ''}${t.totalBaht ? `<div>${t.totalBaht.toLocaleString()} B</div>` : ''}</td>
                    <td class="p-2">${auditCell(p)}</td>
                    <td class="p-2">${importCell(p)}</td>
                    <td class="p-2 text-xs">${msgs}</td>
                </tr>`;
            }).join('');
            return `
                <div class="mb-3 text-lg font-bold text-white">ไฟล์ประเภท: <span class="text-accent">แม่แบบนำเข้าเที่ยวรถ</span></div>
                <div class="grid grid-cols-2 md:grid-cols-5 gap-3 mb-4">
                    ${importStatCard('แถวทั้งหมด', rows.length)}
                    ${importStatCard('เที่ยวใหม่ (จะบันทึก)', nNew)}
                    ${importStatCard('นำเข้าแล้ว (ข้าม)', nAlready)}
                    ${importStatCard('อาจซ้ำ (ข้าม)', nDup)}
                    ${importStatCard('แถวที่ผิด (ไม่บันทึก)', nErr)}
                </div>
                ${nChanged ? `<p class="text-amber-300 text-sm mb-3">⚠️ ${nChanged} แถวเคยนำเข้าแล้วแต่ตัวเลขในไฟล์ต่างจากที่บันทึกไว้ — ระบบไม่ทับข้อมูลเดิม ถ้าต้องแก้ ให้แก้ที่หน้า "คิวรถ"</p>` : ''}
                <div class="overflow-x-auto rounded-lg border border-slate-700 mb-2 max-h-[32rem] overflow-y-auto">
                    <table class="w-full text-sm"><thead class="bg-slate-800 text-gray-400 sticky top-0"><tr>
                        <th class="p-2 text-left">แถว</th><th class="p-2 text-left">วันที่</th><th class="p-2 text-left">ขาขึ้น/ขาล่อง</th><th class="p-2 text-left">ทะเบียน</th><th class="p-2 text-left">คนขับ</th><th class="p-2 text-left">เส้นทาง</th><th class="p-2 text-right">น้ำมัน</th><th class="p-2 text-left">ผลตรวจน้ำมัน</th><th class="p-2 text-left">นำเข้า</th><th class="p-2 text-left">ข้อความ</th>
                    </tr></thead><tbody>${body}</tbody></table>
                </div>
                ${rows.length > 3000 ? `<p class="text-xs text-gray-500 mb-2">แสดง 3,000 แถวแรกจาก ${rows.length.toLocaleString()} แถว — แนะนำใช้ไฟล์ใหม่ทุกเดือน</p>` : ''}
                <div class="bg-sky-900/30 border border-sky-700 rounded-lg p-3 my-4 text-sm text-sky-200">บันทึกเฉพาะ "เที่ยวใหม่" และแถว "อาจซ้ำ" ที่ติ๊กไว้ · ไม่ลบและไม่ทับข้อมูลเดิม · แถวที่ผิดไม่ถูกบันทึก (ดาวน์โหลดไปแก้แล้วนำเข้าใหม่ได้)</div>
                <div class="flex flex-wrap gap-3">
                    <button id="tpl-save-btn" class="bg-success hover:bg-green-600 text-white font-bold py-2 px-5 rounded-lg shadow-md ${nNew ? '' : 'opacity-60 cursor-not-allowed'}" ${nNew ? '' : 'disabled'}>บันทึกเที่ยวใหม่ (<span id="tpl-save-count">${nNew}</span>)</button>
                    ${nErr ? `<button id="tpl-errors-btn" class="bg-slate-600 hover:bg-slate-500 text-white font-bold py-2 px-5 rounded-lg">ดาวน์โหลดแถวที่ผิด (${nErr})</button>` : ''}
                </div>
                <div id="import-result" class="mt-4"></div>`;
        };

        const handleTemplateParse = async (sheetName) => {
            const preview = document.getElementById('import-preview');
            const start = appState.templateStartMonth;
            if (!start) {
                preview.innerHTML = '<p class="text-red-400">ยังไม่ได้ตั้ง "ใช้แม่แบบตั้งแต่เดือน" (ด้านบน) — ตั้งก่อนแล้วกด "อ่าน &amp; ตรวจสอบ" อีกครั้ง</p>';
                return;
            }
            const ws = importWorkbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: true, defval: '' });
            const firstRow = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s.r + 1 : 1;
            const fileName = document.getElementById('import-file')?.files?.[0]?.name || '';
            const res = parseTemplateRows(rows, { startMonth: start, vehicles: appState.vehicles, drivers: appState.drivers, fileName, today: localDateStr() }, firstRow);
            if (res.error) { preview.innerHTML = `<p class="text-red-400">${esc(res.error)}</p>`; return; }
            if (!res.rows.length) { preview.innerHTML = '<p class="text-yellow-400">แม่แบบยังไม่มีข้อมูล — กรอกเที่ยวรถในชีต "นำเข้าเที่ยวรถ" ก่อน</p>'; return; }
            preview.innerHTML = '<p class="text-sky-400">กำลังเทียบกับข้อมูลในระบบ...</p>';
            try {
                const months = [...new Set(res.rows.filter(p => p.status !== 'error').map(p => p.trip.yearMonth))];
                const existing = {};
                for (const ym of months) existing[ym] = (await logisticsService.templateTripsDb.fetchMonth(ym)) || {};
                planTemplateSave(res.rows, existing);
            } catch (err) {
                console.error('Template compare failed:', err);
                preview.innerHTML = `<p class="text-red-400">อ่านข้อมูลเดิมในระบบไม่สำเร็จ จึงยังบันทึกไม่ได้: ${esc(err.message)}</p>`;
                return;
            }
            importParsedTemplate = { ...res, fileName };
            preview.innerHTML = renderTemplatePreview(res, buildAuditContext(appState.rateStandards, appState.vehicles, { originExtras: appState.auditState.originExtras }));
            const nNew = res.rows.filter(p => p.importState === 'new').length;
            const btn = document.getElementById('tpl-save-btn');
            document.querySelectorAll('.tpl-dup-cb').forEach(cb => cb.addEventListener('change', () => {
                const n = nNew + document.querySelectorAll('.tpl-dup-cb:checked').length;
                document.getElementById('tpl-save-count').textContent = n;
                btn.disabled = !n; btn.classList.toggle('opacity-60', !n); btn.classList.toggle('cursor-not-allowed', !n);
            }));
            btn?.addEventListener('click', handleTemplateSave);
            document.getElementById('tpl-errors-btn')?.addEventListener('click', handleTemplateErrorsDownload);
        };

        const handleTemplateSave = async () => {
            if (!importParsedTemplate || importParsedTemplate.saved) return;
            const includeKeys = [...document.querySelectorAll('.tpl-dup-cb:checked')].map(cb => cb.dataset.key);
            const n = importParsedTemplate.rows.filter(p => p.importState === 'new').length + includeKeys.length;
            if (!n || !confirm(`บันทึก ${n.toLocaleString()} เที่ยวจาก "${importParsedTemplate.fileName}"?${includeKeys.length ? `\n(รวมแถว "อาจซ้ำ" ที่ติ๊กไว้ ${includeKeys.length} แถว)` : ''}\n\nไม่ลบและไม่ทับข้อมูลเดิม`)) return;
            const btn = document.getElementById('tpl-save-btn');
            const resultDiv = document.getElementById('import-result');
            btn.disabled = true; btn.classList.add('opacity-60', 'cursor-not-allowed');
            document.querySelectorAll('.tpl-dup-cb').forEach(cb => { cb.disabled = true; });
            resultDiv.innerHTML = '<p class="text-sky-400">กำลังบันทึก...</p>';
            try {
                const { added, skipped, failed, error } = await saveTemplateTrips(logisticsService.templateTripsDb, importParsedTemplate.rows, { includeKeys });
                // One save per preview (read the file again before another save); the error-rows download stays usable.
                importParsedTemplate.saved = true;
                const skippedText = skipped ? ` · ข้าม ${skipped} เที่ยวที่มีคนบันทึกไปก่อนหน้านี้` : '';
                if (failed) {
                    console.error('Template save failed for some trips:', error);
                    resultDiv.innerHTML = `<p class="text-red-400 font-bold">บันทึกได้ ${added.toLocaleString()} เที่ยว แต่ไม่สำเร็จ ${failed} เที่ยว${skippedText} (${esc(error?.message || error)}) — กด "อ่าน &amp; ตรวจสอบ" แล้วบันทึกไฟล์เดิมอีกครั้ง ระบบจะเติมเฉพาะที่ยังขาด ไม่สร้างซ้ำ</p>`;
                } else {
                    resultDiv.innerHTML = `<p class="text-green-400 font-bold">✓ บันทึกแล้ว ${added.toLocaleString()} เที่ยว${skippedText} — ดูผลที่หน้า "ตรวจการเติม" / "คิวรถ"</p>`;
                }
            } catch (err) {
                document.querySelectorAll('.tpl-dup-cb').forEach(cb => { cb.disabled = false; });
                console.error('Template save failed:', err);
                resultDiv.innerHTML = `<p class="text-red-400">บันทึกไม่สำเร็จ: ${esc(err.message)} — นำเข้าไฟล์เดิมอีกครั้งได้ ระบบไม่สร้างเที่ยวซ้ำ</p>`;
                btn.disabled = false; btn.classList.remove('opacity-60', 'cursor-not-allowed');
            }
        };

        // Rows with errors, in the file's own column order plus the reason — fix and import again.
        const handleTemplateErrorsDownload = () => {
            if (!importParsedTemplate) return;
            const { header, rows } = importParsedTemplate;
            const bad = rows.filter(p => p.status === 'error');
            const aoa = [[...header, 'แถวในไฟล์เดิม', 'สิ่งที่ต้องแก้'], ...bad.map(p => [...header.map((_, i) => p.raw[i] ?? ''), p.excelRow, p.errors.join(' / ')])];
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: 'dd/mm/yyyy' }), TEMPLATE_SHEET);
            XLSX.writeFile(wb, `แถวที่ผิด_${(importParsedTemplate.fileName || 'แม่แบบ').replace(/\.xlsx?$/i, '')}.xlsx`);
        };

        // ---- Downhill liters from column L ("คิวรถล่าสุด") of the Transport คิวขาล่อง sheet (option ก, 17 Nov 2026) ----
        // Saved to downhillFuel/{ym}/{tripKey}; the downhill trips themselves are never modified.
        let collPlan = null, collGen = 0;
        const sha1hex = async (s) => [...new Uint8Array(await crypto.subtle.digest('SHA-1', new TextEncoder().encode(s)))].map(b => b.toString(16).padStart(2, '0')).join('');

        const prepareColumnL = async (result) => {
            const gen = ++collGen;
            collPlan = null;
            const box = () => (gen === collGen ? document.getElementById('coll-section') : null);
            try {
                const lastQueue = result.lastQueue || [];
                if (!lastQueue.some(Boolean)) { if (box()) box().innerHTML = ''; return; }
                const keys = await downhillTripKeys(result.trips, sha1hex);
                // Only months that have column L text; only tdh_ trips (the Transport downhill import) can be linked.
                const months = [...new Set(result.trips.filter((t, i) => lastQueue[i]).map(t => t.yearMonth))].sort();
                const existingTrips = {}, existingFuel = {};
                await Promise.all(months.map(async (ym) => {
                    [existingTrips[ym], existingFuel[ym]] = await Promise.all([logisticsService.fetchDownhillTripsMonth(ym), logisticsService.fetchDownhillFuelMonth(ym)]);
                }));
                if (gen !== collGen) return;
                // tdh_ trips still carrying the one-day-early date (not fixed yet) have keys built from that wrong date:
                // a row could hash onto the previous day's identical trip. Leave them out until "แก้วันที่" has run.
                let unfixed = 0;
                for (const ym of months) for (const [k, t] of Object.entries(existingTrips[ym] || {})) {
                    if (!t?.dateFixedFrom) { unfixed++; delete existingTrips[ym][k]; }
                }
                const fileName = document.getElementById('import-file')?.files?.[0]?.name || '';
                const plan = planDownhillFuel(result.trips, lastQueue, keys, existingTrips, existingFuel, { importedFrom: fileName, importedAt: new Date().toLocaleString('th-TH') });
                plan.unfixed = unfixed;
                if (!box()) return;
                collPlan = plan;
                box().innerHTML = renderColumnLSection(plan);
                document.getElementById('coll-save-btn')?.addEventListener('click', handleColumnLSave);
            } catch (err) {
                console.error('Column L read failed:', err);
                if (box()) box().innerHTML = `<p class="text-red-400">อ่านลิตรขาล่องจากช่อง "คิวรถล่าสุด" ไม่สำเร็จ: ${esc(err.message)}</p>`;
            }
        };

        const renderColumnLSection = (plan) => {
            const n = plan.toSave.length;
            const statusRows = Object.entries(COLL_STATUS).map(([k, label]) => {
                const c = plan.byStatus[k] || 0;
                return c ? `<tr class="border-b border-slate-700"><td class="p-2">${label}</td><td class="p-2 text-right">${c.toLocaleString()}</td></tr>` : '';
            }).join('');
            const liters = plan.toSave.reduce((s, e) => s + (e.record.status === 'filled' ? e.record.totalLiters : 0), 0);
            const unparsed = plan.toSave.filter(e => e.record.status === 'unparsed');
            return `
                <div class="bg-slate-800/60 border border-sky-700 rounded-lg p-4">
                    <div class="text-lg font-bold text-white mb-1">ลิตรขาล่องจากช่อง "คิวรถล่าสุด" (คอลัมน์ L)</div>
                    <p class="text-sm text-gray-400 mb-3">เก็บแยกจากเที่ยวรถ — <b class="text-white">ไม่แก้ไขเที่ยวขาล่องเดิม</b> · หน้า "ตรวจการเติม" จะใช้ลิตรนี้ตรวจขาล่องที่ยังไม่มีลิตร</p>
                    <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-3">
                        ${importStatCard('ผูกกับเที่ยวในระบบได้ (จะบันทึก)', n)}
                        ${importStatCard('ลิตรรวม (เฉพาะที่มีลิตร)', liters)}
                        ${importStatCard('บันทึกไว้แล้ว (ข้าม)', plan.already)}
                        ${importStatCard('ไม่มีเที่ยวนี้ในระบบ (ข้าม)', plan.noTrip)}
                    </div>
                    ${plan.unfixed ? `<p class="text-red-300 bg-red-900/30 border border-red-700 rounded-lg p-3 text-sm mb-3">⚠️ เที่ยวขาล่องในระบบ ${plan.unfixed.toLocaleString()} เที่ยวยังบันทึกวันที่เร็วไป 1 วัน — ยังผูกลิตรให้ไม่ได้ ให้กด "แก้วันที่" ในการ์ดด้านบนก่อน แล้วอ่านไฟล์นี้ใหม่</p>` : ''}
                    ${plan.changed ? `<p class="text-amber-300 text-sm mb-3">⚠️ ${plan.changed} รายการเคยบันทึกแล้วแต่ข้อความในไฟล์เปลี่ยนไป — ระบบไม่ทับของเดิม</p>` : ''}
                    ${statusRows ? `<table class="text-sm mb-3"><tbody>${statusRows}</tbody></table>` : ''}
                    ${unparsed.length ? `<details class="mb-3 text-sm"><summary class="cursor-pointer text-yellow-300">อ่านไม่ได้ ${unparsed.length} รายการ — ตรวจเอง</summary><ul class="mt-2 space-y-1 text-gray-300">${unparsed.slice(0, 100).map(e => `<li>${esc(e.record.date)} ${esc(e.record.plate)} · ${esc(e.record.raw)}</li>`).join('')}</ul></details>` : ''}
                    <p class="text-xs text-yellow-300 mb-3">ถ้าเที่ยวขาล่องมีในระบบอยู่แล้ว ใช้ปุ่มนี้อย่างเดียว — ไม่ต้องกด "เพิ่มเข้าไป" ด้านล่าง (จะทำให้เที่ยวซ้ำ)</p>
                    <button id="coll-save-btn" class="bg-success hover:bg-green-600 text-white font-bold py-2 px-5 rounded-lg shadow-md ${n ? '' : 'opacity-60 cursor-not-allowed'}" ${n ? '' : 'disabled'}>บันทึกลิตรขาล่อง (${n.toLocaleString()} รายการ)</button>
                    <div id="coll-result" class="mt-3"></div>
                </div>`;
        };

        const handleColumnLSave = async () => {
            if (!collPlan || collPlan.saved || !collPlan.toSave.length) return;
            const n = collPlan.toSave.length;
            if (!confirm(`บันทึกลิตรขาล่องจากช่อง "คิวรถล่าสุด" ${n.toLocaleString()} รายการ?\n\nเก็บแยกจากเที่ยวรถ ไม่แก้ไขเที่ยวขาล่องเดิม และไม่ทับรายการที่บันทึกไว้แล้ว`)) return;
            const btn = document.getElementById('coll-save-btn');
            const out = document.getElementById('coll-result');
            btn.disabled = true; btn.classList.add('opacity-60', 'cursor-not-allowed');
            out.innerHTML = '<p class="text-sky-400">กำลังบันทึก...</p>';
            const plan = collPlan;
            const { added, skipped, failed, error } = await createAllIfAbsent(plan.toSave, (e) => logisticsService.createDownhillFuel(e.ym, e.key, e.record));
            plan.saved = true;
            if (failed) {
                console.error('Column L save failed for some rows:', error);
                out.innerHTML = `<p class="text-red-400 font-bold">บันทึกได้ ${added.toLocaleString()} รายการ แต่ไม่สำเร็จ ${failed} รายการ (${esc(error?.message || error)}) — กด "อ่าน &amp; ตรวจสอบ" แล้วบันทึกอีกครั้ง ระบบจะเติมเฉพาะที่ยังขาด</p>`;
            } else {
                resetDownhillFuel(); // the audit reloads the months it needs
                out.innerHTML = `<p class="text-green-400 font-bold">✓ บันทึกลิตรขาล่องแล้ว ${added.toLocaleString()} รายการ${skipped ? ` · ข้าม ${skipped} รายการที่มีอยู่แล้ว` : ''} — ดูผลที่หน้า "ตรวจการเติม"</p>`;
            }
        };

        // ---- One-time date fix (user decision A, 17 Nov 2026): Transport trips stored one day early ----
        let dfix = null; // { plan, backupDone, fileName }
        const dfixOut = (html) => { const el = document.getElementById('dfix-out'); if (el) el.innerHTML = html; };
        const thMoney = (n) => Number(n || 0).toLocaleString('th-TH', { maximumFractionDigits: 2 });

        const handleDateFixCheck = async () => {
            const f = document.getElementById('dfix-file')?.files?.[0];
            if (!f) { dfixOut('<p class="text-red-400">เลือกไฟล์ Transport ก่อน</p>'); return; }
            if (dfix?.running) return; // never re-plan while a fix is being written
            dfix = null;
            const btn = document.getElementById('dfix-check-btn');
            btn.disabled = true;
            dfixOut('<p class="text-sky-400">กำลังอ่านไฟล์และข้อมูลในระบบทั้งหมด (อาจใช้เวลา 1–2 นาที)...</p>');
            try {
                const wb = XLSX.read(await f.arrayBuffer(), { type: 'array', cellDates: true });
                const rowsOf = (name) => wb.Sheets[name] ? XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false, defval: '' }) : null;
                const upRows = rowsOf('คิวขาขึ้น'), downRows = rowsOf('คิวขาล่อง');
                if (!upRows || !downRows) { dfixOut('<p class="text-red-400">ไฟล์นี้ไม่มีชีต "คิวขาขึ้น" และ "คิวขาล่อง" — ใช้ไฟล์ Transport</p>'); return; }
                const up = parseUphillQueue(upRows).trips, down = parseDownhillQueue(downRows).trips;
                // tdh_ keys (Transport downhill import, 17 Nov 2026) contain the date: old key (date one day early) → new key.
                const [oldKeys, newKeys] = await Promise.all([downhillTripKeys(down.map(t => ({ ...t, date: prevDay(t.date) })), sha1hex), downhillTripKeys(down, sha1hex)]);
                const tdhKeyMap = Object.fromEntries(oldKeys.map((k, i) => [k, newKeys[i]]));
                const all = await logisticsService.fetchAllTrips();
                const plan = planDateFix([...up, ...down], all, tdhKeyMap);
                dfix = { plan, all, backupDone: false, fileName: f.name };
                dfixOut(renderDateFixPreview(plan));
                document.getElementById('dfix-backup-btn')?.addEventListener('click', handleDateFixBackup);
                document.getElementById('dfix-apply-btn')?.addEventListener('click', handleDateFixApply);
            } catch (err) {
                console.error('Date fix check failed:', err);
                dfixOut(`<p class="text-red-400">ตรวจไม่สำเร็จ: ${esc(err.message)}</p>`);
            } finally { btn.disabled = false; }
        };

        const renderDateFixPreview = (plan) => {
            const c = plan.counts;
            const byMonth = {};
            for (const m of plan.moves) { const b = byMonth[m.fromYm] ||= { n: 0, out: 0, net: 0 }; b.n++; if (m.toYm !== m.fromYm) { b.out++; b.net += Number(m.trip.netTransfer) || 0; } }
            const monthRows = Object.keys(byMonth).sort().map(ym => `<tr class="border-b border-slate-700"><td class="p-2">${esc(ym)}</td><td class="p-2 text-right">${byMonth[ym].n.toLocaleString()}</td><td class="p-2 text-right">${byMonth[ym].out ? `${byMonth[ym].out} เที่ยว → เดือนถัดไป` : '-'}</td><td class="p-2 text-right">${byMonth[ym].out ? thMoney(byMonth[ym].net) : '-'}</td></tr>`).join('');
            const moved = plan.moves.filter(m => m.toYm !== m.fromYm);
            const movedNet = moved.reduce((s, m) => s + (Number(m.trip.netTransfer) || 0), 0);
            const amb = plan.ambiguous.slice(0, 200).map(a => `<li>${esc(a.date)} ${esc(a.direction)} ${esc(a.plate)} ${esc(a.driver)} — ${esc(a.reason)}</li>`).join('');
            return `
                <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
                    ${importStatCard('จะแก้วันที่ (+1 วัน)', c.shift)}
                    ${importStatCard('ย้ายไปเดือนถัดไป', c.monthChange)}
                    ${importStatCard('ไม่พบในไฟล์ (ไม่แตะ)', c.notInFile)}
                    ${importStatCard('บอกไม่ได้ / ชนกัน (ไม่แตะ)', c.ambiguous + c.conflict)}
                </div>
                <p class="mb-3 text-amber-300">⚠️ ${c.monthChange.toLocaleString()} เที่ยวจะย้ายเดือน (วันจริงคือวันที่ 1 ของเดือนถัดไป) · โอนสุทธิรวม ${thMoney(movedNet)} บาท — ยอดรายเดือนในหน้าการเงิน/รายงานของเดือนเหล่านี้จะเปลี่ยน</p>
                <div class="overflow-x-auto max-h-80 overflow-y-auto rounded-lg border border-slate-700 mb-3">
                    <table class="w-full text-sm"><thead class="bg-slate-800 text-gray-400 sticky top-0"><tr><th class="p-2 text-left">เดือน (เดิม)</th><th class="p-2 text-right">แก้วันที่</th><th class="p-2 text-right">ย้ายเดือน</th><th class="p-2 text-right">โอนสุทธิที่ย้าย (บาท)</th></tr></thead><tbody>${monthRows}</tbody></table>
                </div>
                ${amb ? `<details class="mb-3"><summary class="cursor-pointer text-yellow-300">รายการที่ไม่แตะ ${plan.ambiguous.length} รายการ</summary><ul class="mt-2 space-y-1 text-gray-300">${amb}</ul></details>` : ''}
                ${plan.conflicts.length ? `<p class="mb-3 text-yellow-300">ชนกับเที่ยวอื่น ${plan.conflicts.length} รายการ — ไม่แตะ</p>` : ''}
                <ol class="list-decimal ml-5 space-y-1 text-gray-300 mb-3">
                    <li>กด "ดาวน์โหลดสำรองข้อมูล" (บังคับ) — เก็บไฟล์ไว้ ใช้กู้คืนได้</li>
                    <li>กด "แก้วันที่" แล้วยืนยัน · ระหว่างแก้อย่าให้คนอื่นนำเข้าหรือแก้คิวรถ</li>
                </ol>
                <div class="flex flex-wrap gap-3">
                    <button id="dfix-backup-btn" class="bg-slate-600 hover:bg-slate-500 text-white font-bold py-2 px-4 rounded-lg">ดาวน์โหลดสำรองข้อมูล</button>
                    <button id="dfix-apply-btn" class="bg-error hover:bg-red-600 text-white font-bold py-2 px-4 rounded-lg opacity-60 cursor-not-allowed" disabled>แก้วันที่ ${c.shift.toLocaleString()} เที่ยว</button>
                </div>
                <div id="dfix-result" class="mt-3"></div>`;
        };

        const handleDateFixBackup = () => {
            if (!dfix) return;
            const blob = new Blob([JSON.stringify(dfix.all)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = `trips-backup-before-datefix-${localDateStr()}.json`;
            document.body.appendChild(a); a.click(); a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 2000);
            dfix.backupDone = true;
            const btn = document.getElementById('dfix-apply-btn');
            if (btn && dfix.plan.moves.length) { btn.disabled = false; btn.classList.remove('opacity-60', 'cursor-not-allowed'); }
        };

        const handleDateFixApply = async () => {
            const cur = dfix; // the plan being applied — a new "ตรวจ" must not swap it out mid-run
            if (!cur || !cur.backupDone || cur.running || cur.done) return;
            const { plan } = cur;
            if (!confirm(`แก้วันที่ ${plan.counts.shift.toLocaleString()} เที่ยว (+1 วัน ให้ตรงกับไฟล์ Excel)?\n\n${plan.counts.monthChange} เที่ยวจะย้ายไปเดือนถัดไป\nเก็บวันเดิมไว้ในช่อง dateFixedFrom ของแต่ละเที่ยว\n\nสำรองข้อมูลแล้วใช่ไหม?`)) return;
            cur.running = true;
            const btn = document.getElementById('dfix-apply-btn');
            const checkBtn = document.getElementById('dfix-check-btn');
            btn.disabled = true; btn.classList.add('opacity-60', 'cursor-not-allowed');
            if (checkBtn) checkBtn.disabled = true;
            const out = document.getElementById('dfix-result');
            const committedFrom = new Set(); // source paths of moves actually written so far
            const batches = buildDateFixBatches(plan.moves);
            let done = 0, dropped = 0, failedAt = '';
            for (const b0 of batches) {
                if (out) out.innerHTML = `<p class="text-sky-400">กำลังแก้เดือน ${esc(b0.month)} (${done.toLocaleString()} / ${plan.counts.shift.toLocaleString()} เที่ยว)...</p>`;
                try {
                    // Re-read the months this batch touches; skip trips changed since the preview.
                    const months = [...new Set(b0.moves.flatMap(m => [m.fromYm, m.toYm]))];
                    const fresh = Object.fromEntries(await Promise.all(months.map(async ym => [ym, await logisticsService.fetchTripsMonthRaw(ym)])));
                    const b = recheckBatch(b0, fresh, committedFrom);
                    if (Object.keys(b.updates).length) await logisticsService.applyTripsUpdate(b.updates); // atomic per month
                    for (const m of b.moves) committedFrom.add(`${m.fromYm}/${m.fromKey}`);
                    done += b.moves.length; dropped += b.dropped.length;
                } catch (err) {
                    console.error('Date fix batch failed:', b0.month, err);
                    failedAt = `${b0.month}: ${err.message}`;
                    break;
                }
            }
            cur.running = false; cur.done = !failedAt;
            if (checkBtn) checkBtn.disabled = false;
            const entry = { at: new Date().toISOString(), file: cur.fileName, planned: plan.counts.shift, fixed: done, skipped: dropped, monthChange: plan.counts.monthChange, error: failedAt || '' };
            logisticsService.logDateFix(entry).catch(err => console.error('Date fix log failed:', err));
            if (out) out.innerHTML = failedAt
                ? `<p class="text-red-400 font-bold">หยุดที่เดือน ${esc(failedAt)} — แก้ไปแล้ว ${done.toLocaleString()} เที่ยว (เดือนที่สำเร็จแก้ครบทั้งเดือน) · กด "ตรวจ" ใหม่แล้วแก้ต่อได้ ระบบจะแก้เฉพาะที่เหลือ</p>`
                : `<p class="text-green-400 font-bold">✓ แก้วันที่แล้ว ${done.toLocaleString()} เที่ยว${dropped ? ` · ข้าม ${dropped} เที่ยวที่ถูกแก้ระหว่างนั้น` : ''} — กด "ตรวจ" อีกครั้งเพื่อยืนยัน (ควรเหลือ 0)</p>`;
        };

        // Daily-board / Transport imports overlapping the template period → the same trip may be counted twice.
        const templateOverlapNote = (trips) => {
            const start = appState.templateStartMonth;
            if (!start) return '';
            const months = [...new Set((trips || []).map(t => t.yearMonth).filter(ym => ym && ym >= start))].sort();
            if (!months.length) return '';
            const [y, m] = start.split('-').map(Number);
            return `เดือน ${months.join(', ')} อยู่ในช่วงที่ใช้แม่แบบแล้ว (ตั้งแต่ ${m}/${y + 543}) — นำเข้าไฟล์นี้อาจนับเที่ยวซ้ำกับที่คนจัดรถกรอก`;
        };

        const attachImportListeners = () => {
            document.getElementById('import-file')?.addEventListener('change', handleImportFile);
            document.getElementById('import-parse-btn')?.addEventListener('click', handleImportParse);
            document.getElementById('tpl-download-btn')?.addEventListener('click', handleTemplateDownload);
            document.getElementById('dfix-check-btn')?.addEventListener('click', handleDateFixCheck);
            document.getElementById('tpl-start-save')?.addEventListener('click', handleTemplateStartSave);
        };

        // --- View: Queue / Dispatch (คิวรถ) — loads one month at a time ---
        const renderQueue = () => {
            const month = appState.tripsMonth || localDateStr().slice(0, 7);
            return `
                <div class="space-y-6">
                    ${renderCard(`
                        <div class="flex flex-wrap justify-between items-end gap-4 mb-4">
                            <h2 class="text-2xl font-bold text-white flex items-center gap-3">${ICONS.Truck} คิวรถ (Dispatch)</h2>
                            <div class="flex flex-wrap items-end gap-3">
                                <div><label class="block text-xs text-gray-400 mb-1">เดือน</label>
                                    <input type="month" id="queue-month" value="${month}" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>
                                <div><label class="block text-xs text-gray-400 mb-1">ทิศทาง</label>
                                    <select id="queue-direction" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white">
                                        <option value="">ทั้งหมด</option>
                                        <option value="ขาขึ้น" ${appState.queueDirectionFilter === 'ขาขึ้น' ? 'selected' : ''}>ขาขึ้น</option>
                                        <option value="ขาล่อง" ${appState.queueDirectionFilter === 'ขาล่อง' ? 'selected' : ''}>ขาล่อง</option>
                                    </select></div>
                                <div><label class="block text-xs text-gray-400 mb-1">ค้นหา</label>
                                    <input type="text" id="queue-search" value="${appState.queueSearch || ''}" placeholder="คนขับ/ทะเบียน/เส้นทาง" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white w-56" /></div>
                                <button id="queue-add-btn" class="flex items-center gap-2 bg-primary hover:bg-secondary text-white font-bold py-2 px-4 rounded-lg shadow-md">${ICONS.Plus} เพิ่มคิว</button>
                            </div>
                        </div>
                        <div id="queue-content">${renderQueueContent()}</div>
                    `)}
                </div>
            `;
        };

        const renderQueueContent = () => {
            const dirFilter = appState.queueDirectionFilter;
            const q = (appState.queueSearch || '').toLowerCase();
            let trips = logisticsService.getTrips();
            if (dirFilter) trips = trips.filter(t => t.direction === dirFilter);
            if (q) trips = trips.filter(t => [t.driver, t.plate, t.origin, t.destination, t.vehicleType].some(v => (v || '').toLowerCase().includes(q)));

            const sumRevenue = trips.reduce((s, t) => s + (t.revenue || 0), 0);
            const sumNet = trips.reduce((s, t) => s + (t.netTransfer || 0), 0);
            const sumLiters = trips.reduce((s, t) => s + (t.totalLiters || 0), 0);

            const rows = trips.map(t => `
                <tr class="border-b border-slate-700 hover:bg-slate-700/20">
                    <td class="px-2 py-2 whitespace-nowrap">${t.date || ''}</td>
                    <td class="px-2 py-2 text-center whitespace-nowrap"><span class="px-2 py-0.5 rounded text-xs ${t.direction === 'ขาขึ้น' ? 'bg-sky-700' : 'bg-amber-700'} text-white">${t.direction || ''}</span></td>
                    <td class="px-2 py-2 whitespace-nowrap">${t.driver || ''}</td>
                    <td class="px-2 py-2 font-mono whitespace-nowrap">${t.plate || ''}</td>
                    <td class="px-2 py-2 whitespace-nowrap">${t.vehicleType || ''}</td>
                    <td class="px-2 py-2 whitespace-nowrap">${t.origin || ''} → ${t.destination || ''}</td>
                    <td class="px-2 py-2 text-right whitespace-nowrap text-accent">${t.totalLiters ? t.totalLiters.toLocaleString() + ' L' : ''}${t.totalBaht ? ' ' + t.totalBaht.toLocaleString() + ' B' : ''}</td>
                    <td class="px-2 py-2 text-right whitespace-nowrap">${t.revenue ? t.revenue.toLocaleString() : ''}</td>
                    <td class="px-2 py-2 text-right whitespace-nowrap font-semibold text-green-400">${t.netTransfer ? t.netTransfer.toLocaleString() : ''}</td>
                    <td class="px-2 py-2 text-center">
                        <div class="flex justify-center gap-1">
                            <button data-action="edit-trip" data-id="${t.id}" class="text-blue-400 hover:text-blue-300 p-1">${ICONS.Pencil}</button>
                            <button data-action="delete-trip" data-id="${t.id}" class="text-red-500 hover:text-red-400 p-1">${ICONS.Trash}</button>
                        </div>
                    </td>
                </tr>`).join('');

            return `
                <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
                    ${importStatCard('จำนวนคิว', trips.length)}
                    ${importStatCard('น้ำมันรวม (ลิตร)', sumLiters)}
                    ${importStatCard('รายรับรวม (บาท)', sumRevenue)}
                    ${importStatCard('สุทธิโอนรวม (บาท)', sumNet)}
                </div>
                <div class="overflow-x-auto">
                    <table class="w-full text-left text-sm">
                        <thead class="bg-slate-800 text-gray-400 uppercase sticky top-0 z-10"><tr>
                            <th class="px-2 py-2 whitespace-nowrap">วันที่</th><th class="px-2 py-2 text-center">ทิศทาง</th><th class="px-2 py-2">คนขับ</th><th class="px-2 py-2">ทะเบียน</th><th class="px-2 py-2">ประเภท</th><th class="px-2 py-2">เส้นทาง</th><th class="px-2 py-2 text-right">น้ำมัน</th><th class="px-2 py-2 text-right">รายรับ</th><th class="px-2 py-2 text-right">สุทธิโอน</th><th class="px-2 py-2 text-center">จัดการ</th>
                        </tr></thead>
                        <tbody>${rows}</tbody>
                    </table>
                    ${trips.length === 0 ? `<p class="text-center py-8 text-gray-500">ไม่มีคิวรถในเดือน ${appState.tripsMonth || '(ยังไม่เลือก)'} — เลือกเดือนอื่น หรือกด "เพิ่มคิว"</p>` : ''}
                </div>
            `;
        };

        const attachQueueListeners = () => {
            const monthInput = document.getElementById('queue-month');
            const targetMonth = appState.tripsMonth || (monthInput ? monthInput.value : '');
            // Load the month's data on first entry (or if it differs from what's loaded).
            if (targetMonth && appState.tripsMonth !== targetMonth) {
                logisticsService.loadTripsForMonth(targetMonth);
            }
            monthInput?.addEventListener('change', (e) => logisticsService.loadTripsForMonth(e.target.value));
            document.getElementById('queue-direction')?.addEventListener('change', (e) => {
                appState.queueDirectionFilter = e.target.value;
                const c = document.getElementById('queue-content'); if (c) c.innerHTML = renderQueueContent();
            });
            document.getElementById('queue-search')?.addEventListener('input', (e) => {
                appState.queueSearch = e.target.value;
                const c = document.getElementById('queue-content'); if (c) c.innerHTML = renderQueueContent();
            });
            document.getElementById('queue-add-btn')?.addEventListener('click', handleAddTrip);
        };

        const renderTripFormContent = (trip) => {
            const litersItems = trip?.gasStationsLiters ? Object.entries(trip.gasStationsLiters).map(([stationKey, amount]) => ({ stationKey, amount })) : [];
            const bahtItems = trip?.gasStationsBaht ? Object.entries(trip.gasStationsBaht).map(([stationKey, amount]) => ({ stationKey, amount })) : [];
            const money = (name, label, val) => `
                <div><label class="block text-xs text-gray-400 mb-1">${label}</label>
                    <input type="number" name="${name}" value="${val || ''}" class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>`;

            // Dropdown options pulled from real data (pick from the list OR type a new value).
            const uniq = (arr) => [...new Set(arr.filter(Boolean))];
            const runTypes = (() => { const rt = uniq(logisticsService.getRateTable().map(r => r.runType)); return rt.length ? rt : ['หนักล่อง', 'ตีเปล่า', 'คิวพิเศษ']; })();
            const vehTypes = uniq([...logisticsService.getRateTable().map(r => r.vehicleType), ...logisticsService.getVehicles().map(v => v.vehicleType), ...logisticsService.getRoutes().map(r => r.vehicleType)]).sort();
            const plates = uniq(logisticsService.getVehicles().filter(v => !(v.driveType || '').includes('ร่วม')).map(v => v.registrationNumber));
            const driverNames = uniq(logisticsService.getDrivers().map(d => d.name));
            const origins = logisticsService.getUniqueValues('origin');
            const dests = logisticsService.getUniqueValues('destination');
            const dl = (name, label, val, opts) => `
                <div><label class="block text-xs text-gray-400 mb-1">${label}</label>
                    <input type="text" name="${name}" value="${val || ''}" list="dl-${name}" autocomplete="off" class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" />
                    <datalist id="dl-${name}">${opts.map(o => `<option value="${o}"></option>`).join('')}</datalist></div>`;

            return `
                <form id="trip-form" class="space-y-5">
                    <input type="hidden" name="id" value="${trip?.id || ''}" />
                    <input type="hidden" name="originalMonth" value="${trip?.yearMonth || ''}" />
                    <div class="grid grid-cols-2 sm:grid-cols-4 gap-3">
                        <div><label class="block text-xs text-gray-400 mb-1">วันที่</label>
                            <input type="date" name="date" value="${trip?.date || localDateStr()}" required class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>
                        <div><label class="block text-xs text-gray-400 mb-1">ทิศทาง</label>
                            <select name="direction" class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white">
                                <option value="ขาขึ้น" ${(trip?.direction || 'ขาขึ้น') === 'ขาขึ้น' ? 'selected' : ''}>ขาขึ้น</option>
                                <option value="ขาล่อง" ${trip?.direction === 'ขาล่อง' ? 'selected' : ''}>ขาล่อง</option>
                            </select></div>
                        ${dl('plate', 'ทะเบียน', trip?.plate, plates)}
                        ${dl('vehicleType', 'ประเภทรถ', trip?.vehicleType, vehTypes)}
                    </div>
                    <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                        ${dl('driver', 'คนขับ', trip?.driver, driverNames)}
                        ${dl('origin', 'ต้นทาง (ขึ้นที่)', trip?.origin, origins)}
                        ${dl('destination', 'ปลายทาง (ลงที่)', trip?.destination, dests)}
                    </div>
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        ${dl('queueType', 'ประเภทคิว/วิ่ง', trip?.queueType, runTypes)}
                        <div><label class="block text-xs text-gray-400 mb-1">รายการงาน</label>
                            <input type="text" name="jobDetail" value="${trip?.jobDetail || ''}" class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>
                    </div>

                    <button type="button" data-action="apply-route-to-trip" class="text-sm bg-success hover:bg-green-600 text-white py-1.5 px-3 rounded-lg">⛽ ดึงน้ำมันมาตรฐานจากเส้นทาง (ตาม ประเภทรถ + ต้นทาง + ปลายทาง)</button>
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div class="space-y-3 rounded-lg border border-slate-600 p-3">
                            <h4 class="font-semibold text-white">น้ำมัน (ลิตร)</h4>
                            <div id="route-distribution-list-liters">${litersItems.map((it, i) => renderRouteDistributionRow(it, i, 'Liters')).join('')}</div>
                            <button type="button" data-action="add-route-distribution" data-unit="Liters" class="w-full border-2 border-dashed border-slate-500 hover:bg-slate-700/50 text-gray-300 text-sm py-1.5 rounded-lg">+ เพิ่มปั๊ม (ลิตร)</button>
                        </div>
                        <div class="space-y-3 rounded-lg border border-slate-600 p-3">
                            <h4 class="font-semibold text-white">น้ำมัน (บาท)</h4>
                            <div id="route-distribution-list-baht">${bahtItems.map((it, i) => renderRouteDistributionRow(it, i, 'Baht')).join('')}</div>
                            <button type="button" data-action="add-route-distribution" data-unit="Baht" class="w-full border-2 border-dashed border-slate-500 hover:bg-slate-700/50 text-gray-300 text-sm py-1.5 rounded-lg">+ เพิ่มปั๊ม (บาท)</button>
                        </div>
                    </div>

                    <div class="grid grid-cols-2 sm:grid-cols-3 gap-3">
                        ${money('revenue', 'รายรับ / เงินที่ต้องได้', trip?.revenue)}
                        ${money('netTransfer', 'สุทธิโอน / รวมเงิน', trip?.netTransfer)}
                        ${money('fuelCost', 'ค่าน้ำมัน / ค่าแก๊ส', trip?.fuelCost)}
                        ${money('driverAllowance', 'เบี้ยขับ / ค่าเดินทาง', trip?.driverAllowance)}
                        ${money('deduction', 'หักเงิน', trip?.deduction)}
                        ${money('distance', 'ระยะทาง (กม.)', trip?.distance)}
                    </div>
                    <button type="button" data-action="calc-allowance" class="text-sm bg-sky-700 hover:bg-sky-600 text-white py-1.5 px-3 rounded-lg">🧮 คำนวณ เบี้ยเลี้ยง / หัก / สุทธิโอน จากตารางเบี้ยเลี้ยง (ตาม ประเภทคิว + ประเภทรถ)</button>
                    <div><label class="block text-xs text-gray-400 mb-1">หมายเหตุ</label>
                        <input type="text" name="notes" value="${trip?.notes || ''}" class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>

                    <footer class="flex justify-end gap-4 pt-2">
                        <button type="button" data-action="close-modal" class="bg-slate-600 hover:bg-slate-500 text-white font-bold py-2 px-4 rounded-lg">ยกเลิก</button>
                        <button type="submit" class="bg-primary hover:bg-secondary text-white font-bold py-2 px-6 rounded-lg">บันทึก</button>
                    </footer>
                </form>
            `;
        };

        // Fill the trip's fuel-per-pump from the matching route (ประเภทรถ+ต้นทาง+ปลายทาง).
        const handleApplyRouteToTrip = (button) => {
            const form = button.closest('#trip-form');
            if (!form) return;
            const vehicleType = form.querySelector('[name="vehicleType"]').value;
            const origin = form.querySelector('[name="origin"]').value;
            const destination = form.querySelector('[name="destination"]').value;
            const route = logisticsService.findMatchingRoute(vehicleType, origin, destination);
            if (!route) {
                alert(`ไม่พบเส้นทางที่ตรงกับ:\n${vehicleType || '(ว่าง)'} · ${origin || '(ว่าง)'} → ${destination || '(ว่าง)'}\n\n(ต้องมี ประเภทรถ + ต้นทาง + ปลายทาง ตรงกับข้อมูลเส้นทาง 492 เส้นทาง)`);
                return;
            }
            const fill = (listId, stations, unit) => {
                const c = form.querySelector(`#${listId}`);
                if (!c) return;
                c.innerHTML = '';
                Object.entries(stations || {}).forEach(([stationKey, amount], i) => {
                    c.insertAdjacentHTML('beforeend', renderRouteDistributionRow({ stationKey, amount }, i, unit));
                });
            };
            fill('route-distribution-list-liters', route.gasStationsLiters, 'Liters');
            fill('route-distribution-list-baht', route.gasStationsBaht, 'Baht');
            const nL = Object.keys(route.gasStationsLiters || {}).length, nB = Object.keys(route.gasStationsBaht || {}).length;
            if (nL + nB === 0) alert('เส้นทางนี้ไม่มีข้อมูลน้ำมัน (ไม่ต้องเติม)');
        };

        // Auto-fill เบี้ยเลี้ยง/หัก/สุทธิโอน from the imported rate table (ตารางเบี้ยเลี้ยง). silent = no alert.
        const handleCalcAllowance = (form, silent = false) => {
            form = form?.closest ? form.closest('#trip-form') : form;
            if (!form) return;
            const queueType = form.querySelector('[name="queueType"]').value;
            const vehicleType = form.querySelector('[name="vehicleType"]').value;
            const rate = logisticsService.findRate(queueType, vehicleType);
            if (!rate) {
                if (!silent) alert(`ไม่พบอัตราสำหรับ ประเภทคิว "${queueType || '(ว่าง)'}" + ประเภทรถ "${vehicleType || '(ว่าง)'}"\n\nประเภทวิ่งในตาราง: หนักล่อง, ตีเปล่า, คิวพิเศษ`);
                return;
            }
            form.querySelector('[name="driverAllowance"]').value = rate.allowance;
            form.querySelector('[name="deduction"]').value = rate.deduction;
            form.querySelector('[name="netTransfer"]').value = rate.allowance - rate.deduction; // สุทธิโอน = เบี้ยเลี้ยง − หัก
        };

        const attachTripFormExtras = () => {
            const form = document.getElementById('trip-form');
            if (!form) return;
            form.addEventListener('submit', handleTripFormSubmit);
            // Picking a ทะเบียน auto-fills ประเภทรถ from the vehicle master.
            const plateInput = form.querySelector('[name="plate"]');
            plateInput?.addEventListener('change', () => {
                const v = logisticsService.getVehicles().find(x => (x.registrationNumber || '').trim() === plateInput.value.trim());
                if (v && v.vehicleType) form.querySelector('[name="vehicleType"]').value = v.vehicleType;
            });
            // Auto-calc allowance when ประเภทคิว or ประเภทรถ changes (silent).
            form.querySelector('[name="queueType"]')?.addEventListener('change', () => handleCalcAllowance(form, true));
            form.querySelector('[name="vehicleType"]')?.addEventListener('change', () => handleCalcAllowance(form, true));
        };
        const handleAddTrip = () => { openModal('เพิ่มคิวรถ', renderTripFormContent()); attachTripFormExtras(); };
        const handleEditTrip = (id) => {
            const trip = logisticsService.getTrips().find(t => t.id === id);
            if (!trip) return;
            openModal('แก้ไขคิวรถ', renderTripFormContent(trip));
            attachTripFormExtras();
        };
        const handleDeleteTrip = async (id) => {
            if (!confirm('ต้องการลบคิวรถนี้ใช่หรือไม่?')) return;
            const ym = appState.tripsMonth;
            const trip = logisticsService.getTrips().find(t => t.id === id);
            try {
                await logisticsService.deleteTrip(ym, id);
                if (trip) {
                    const { id: _id, ...data } = trip;
                    showUndoToast('ลบคิวรถแล้ว', () => logisticsService.restoreTrip(ym, id, data));
                }
            } catch (err) {
                console.error('Delete trip failed:', err);
                alert('ลบคิวรถไม่สำเร็จ: ' + (err?.message || err));
            }
        };
        const handleTripFormSubmit = async (e) => {
            e.preventDefault();
            const form = e.target;
            const fd = new FormData(form);
            const id = fd.get('id');
            const originalMonth = fd.get('originalMonth');
            const getDist = (unit) => {
                const container = form.querySelector(`#route-distribution-list-${unit.toLowerCase()}`);
                const data = {};
                container?.querySelectorAll('.distribution-row').forEach(row => {
                    const key = row.querySelector('[data-field="stationKey"]').value;
                    const amount = Number(row.querySelector('[data-field="amount"]').value) || 0;
                    if (key && amount > 0) data[key] = (data[key] || 0) + amount;
                });
                return data;
            };
            const date = fd.get('date');
            if (!date) { alert('กรุณาระบุวันที่'); return; }
            const gasStationsLiters = getDist('Liters');
            const gasStationsBaht = getDist('Baht');
            const trip = {
                date, yearMonth: date.slice(0, 7),
                direction: fd.get('direction') || 'ขาขึ้น',
                driver: fd.get('driver') || '', plate: fd.get('plate') || '', vehicleType: fd.get('vehicleType') || '',
                origin: fd.get('origin') || '', destination: fd.get('destination') || '',
                queueType: fd.get('queueType') || '', jobDetail: fd.get('jobDetail') || '', notes: fd.get('notes') || '',
                gasStationsLiters, gasStationsBaht,
                totalLiters: Object.values(gasStationsLiters).reduce((a, b) => a + b, 0),
                totalBaht: Object.values(gasStationsBaht).reduce((a, b) => a + b, 0),
                revenue: Number(fd.get('revenue')) || 0,
                netTransfer: Number(fd.get('netTransfer')) || 0,
                fuelCost: Number(fd.get('fuelCost')) || 0,
                driverAllowance: Number(fd.get('driverAllowance')) || 0,
                deduction: Number(fd.get('deduction')) || 0,
                distance: Number(fd.get('distance')) || 0,
            };
            try {
                if (id) {
                    // Keep the fields this form does not edit (source, importedFrom, sourceRow, crossing/waiting fees,
                    // and the template's trailerDrop / trailerPlate, which the fuel audit relies on).
                    const { id: _omit, ...existing } = appState.trips.find(t => t.id === id) || {};
                    Object.assign(trip, { ...existing, ...trip });
                    if (originalMonth && originalMonth !== trip.yearMonth) {
                        // Moving to another month: write to the new month FIRST (keeping its key), then remove
                        // the old one. If the write fails we throw before deleting, so the trip is never lost.
                        await logisticsService.updateTrip(trip.yearMonth, id, trip);
                        await logisticsService.deleteTrip(originalMonth, id);
                    } else {
                        await logisticsService.updateTrip(trip.yearMonth, id, trip);
                    }
                } else {
                    await logisticsService.addTrip(trip.yearMonth, trip);
                }
                closeModal();
                // If the trip lands in a different month than the one being viewed, jump there.
                if (trip.yearMonth !== appState.tripsMonth) logisticsService.loadTripsForMonth(trip.yearMonth);
            } catch (err) {
                console.error('Save trip failed:', err);
                alert('บันทึกคิวรถไม่สำเร็จ: ' + err.message);
            }
        };

        // --- View: Drivers (คนขับ) ---
        const renderDrivers = () => {
            return `
                <div class="space-y-6">
                    ${renderCard(`
                        <div class="flex flex-wrap justify-between items-center gap-4 mb-4">
                            <h2 class="text-2xl font-bold text-white flex items-center gap-3">${ICONS.IdCard} คนขับ</h2>
                            <div class="flex items-center gap-3">
                                <input type="text" id="driver-filter" value="${appState.driverFilter || ''}" placeholder="ค้นหา ชื่อ/เลขบัญชี" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white w-64" />
                                <button id="add-driver-btn" class="flex items-center gap-2 bg-primary hover:bg-secondary text-white font-bold py-2 px-4 rounded-lg shadow-md">${ICONS.Plus} เพิ่มคนขับ</button>
                            </div>
                        </div>
                        <div id="drivers-content">${renderDriversContent()}</div>
                    `)}
                </div>`;
        };
        const renderDriversContent = () => {
            const filter = (appState.driverFilter || '').toLowerCase();
            let drivers = logisticsService.getDrivers();
            if (filter) drivers = drivers.filter(d => (d.name || '').toLowerCase().includes(filter) || (d.bankAccount || '').toLowerCase().includes(filter));
            const rows = drivers.map(d => `
                <tr class="border-b border-slate-600 hover:bg-slate-700/20">
                    <td class="p-3">${d.name || ''}</td>
                    <td class="p-3 font-mono">${d.bankAccount || '<span class="text-gray-600">—</span>'}</td>
                    <td class="p-3 text-center"><div class="flex justify-center gap-2">
                        <button data-action="edit-driver" data-id="${d.id}" class="text-blue-400 hover:text-blue-300 p-2">${ICONS.Pencil}</button>
                        <button data-action="delete-driver" data-id="${d.id}" class="text-red-500 hover:text-red-400 p-2">${ICONS.Trash}</button>
                    </div></td>
                </tr>`).join('');
            return `
                <p class="text-sm text-gray-400 mb-2">ทั้งหมด ${drivers.length} คน</p>
                <div class="overflow-x-auto"><table class="w-full text-left">
                    <thead class="bg-slate-800 text-gray-400 uppercase text-sm"><tr><th class="p-3">ชื่อ-นามสกุล</th><th class="p-3">เลขบัญชีธนาคาร</th><th class="p-3 text-center w-24">จัดการ</th></tr></thead>
                    <tbody id="drivers-body">${rows}</tbody>
                </table>
                ${drivers.length === 0 ? '<p class="text-center py-8 text-gray-500">ยังไม่มีข้อมูลคนขับ — เพิ่มเอง หรือนำเข้าจากชีต "คนขับ"</p>' : ''}
                </div>`;
        };
        const attachDriversListeners = () => {
            document.getElementById('driver-filter')?.addEventListener('input', (e) => {
                appState.driverFilter = e.target.value;
                const c = document.getElementById('drivers-content'); if (c) c.innerHTML = renderDriversContent();
            });
            document.getElementById('add-driver-btn')?.addEventListener('click', handleAddDriver);
        };
        const renderDriverFormContent = (driver) => `
            <form id="driver-form" class="space-y-5">
                <input type="hidden" name="id" value="${driver?.id || ''}" />
                <div><label class="block text-sm text-gray-400 mb-1">ชื่อ-นามสกุล</label>
                    <input type="text" name="name" value="${driver?.name || ''}" required class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>
                <div><label class="block text-sm text-gray-400 mb-1">เลขบัญชีธนาคาร</label>
                    <input type="text" name="bankAccount" value="${driver?.bankAccount || ''}" class="w-full bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white font-mono" /></div>
                <footer class="flex justify-end gap-4 pt-2">
                    <button type="button" data-action="close-modal" class="bg-slate-600 hover:bg-slate-500 text-white font-bold py-2 px-4 rounded-lg">ยกเลิก</button>
                    <button type="submit" class="bg-primary hover:bg-secondary text-white font-bold py-2 px-6 rounded-lg">บันทึก</button>
                </footer>
            </form>`;
        const handleAddDriver = () => { openModal('เพิ่มคนขับ', renderDriverFormContent()); document.getElementById('driver-form').addEventListener('submit', handleDriverFormSubmit); };
        const handleEditDriver = (id) => { const d = logisticsService.getDrivers().find(x => x.id === id); if (!d) return; openModal('แก้ไขคนขับ', renderDriverFormContent(d)); document.getElementById('driver-form').addEventListener('submit', handleDriverFormSubmit); };
        const handleDeleteDriver = async (id) => {
            if (!confirm('ลบคนขับนี้ใช่หรือไม่?')) return;
            const d = logisticsService.getDrivers().find(x => x.id === id);
            try { await logisticsService.deleteDriver(id); if (d) { const { id: _i, ...data } = d; showUndoToast('ลบคนขับแล้ว', () => logisticsService.restoreDriver(id, data)); } }
            catch (e) { console.error('Delete driver failed:', e); }
        };
        const handleDriverFormSubmit = async (e) => {
            e.preventDefault();
            const fd = new FormData(e.target); const id = fd.get('id');
            const data = { name: (fd.get('name') || '').trim(), bankAccount: (fd.get('bankAccount') || '').trim() };
            if (!data.name) { alert('กรุณากรอกชื่อ'); return; }
            try { if (id) await logisticsService.updateDriver(id, data); else await logisticsService.addDriver(data); closeModal(); }
            catch (err) { alert('บันทึกไม่สำเร็จ: ' + err.message); }
        };

        // --- View: Payroll (สรุปการเงินคนขับ) — aggregated from the loaded month's trips ---
        const aggregatePayroll = () => {
            const trips = logisticsService.getTrips();
            const byDriver = {};
            trips.forEach(t => {
                const key = t.driver || '(ไม่ระบุ)';
                if (!byDriver[key]) byDriver[key] = { driver: key, count: 0, revenue: 0, allowance: 0, deduction: 0, net: 0 };
                const g = byDriver[key];
                g.count++; g.revenue += t.revenue || 0; g.allowance += t.driverAllowance || 0; g.deduction += t.deduction || 0; g.net += t.netTransfer || 0;
            });
            const bankOf = {}; logisticsService.getDrivers().forEach(d => { bankOf[(d.name || '').trim()] = d.bankAccount || ''; });
            const list = Object.values(byDriver).sort((a, b) => b.net - a.net);
            return { list, bankOf };
        };
        const renderPayroll = () => {
            const month = appState.tripsMonth || localDateStr().slice(0, 7);
            return `
                <div class="space-y-6">
                    ${renderCard(`
                        <div class="flex flex-wrap justify-between items-end gap-4 mb-4">
                            <h2 class="text-2xl font-bold text-white flex items-center gap-3">${ICONS.ChartBar} สรุปการเงินคนขับ (Payroll)</h2>
                            <div class="flex items-end gap-3">
                                <div><label class="block text-xs text-gray-400 mb-1">เดือน</label>
                                    <input type="month" id="payroll-month" value="${month}" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>
                                <button id="payroll-export-btn" class="flex items-center gap-2 bg-success hover:bg-green-600 text-white font-bold py-2 px-4 rounded-lg shadow-md">${ICONS.Download} Export</button>
                            </div>
                        </div>
                        <div id="payroll-content">${renderPayrollContent()}</div>
                    `)}
                </div>`;
        };
        const renderPayrollContent = () => {
            const { list, bankOf } = aggregatePayroll();
            const rows = list.map(g => `
                <tr class="border-b border-slate-700 hover:bg-slate-700/20">
                    <td class="p-2 whitespace-nowrap">${g.driver}</td>
                    <td class="p-2 font-mono text-gray-400 whitespace-nowrap">${bankOf[g.driver.trim()] || ''}</td>
                    <td class="p-2 text-right">${g.count}</td>
                    <td class="p-2 text-right">${g.revenue.toLocaleString()}</td>
                    <td class="p-2 text-right">${g.allowance.toLocaleString()}</td>
                    <td class="p-2 text-right text-red-400">${g.deduction.toLocaleString()}</td>
                    <td class="p-2 text-right font-bold text-green-400">${g.net.toLocaleString()}</td>
                </tr>`).join('');
            const tot = list.reduce((a, g) => ({ count: a.count + g.count, revenue: a.revenue + g.revenue, allowance: a.allowance + g.allowance, deduction: a.deduction + g.deduction, net: a.net + g.net }), { count: 0, revenue: 0, allowance: 0, deduction: 0, net: 0 });
            return `
                <p class="text-xs text-gray-500 mb-3">เดือน ${appState.tripsMonth || '(เลือกเดือน)'} · ${list.length} คนขับ · คำนวณจากคิวรถของเดือนนั้น</p>
                <div class="overflow-x-auto"><table class="w-full text-left text-sm">
                    <thead class="bg-slate-800 text-gray-400 uppercase sticky top-0 z-10"><tr><th class="p-2">คนขับ</th><th class="p-2">เลขบัญชี</th><th class="p-2 text-right">คิว</th><th class="p-2 text-right">รายรับ</th><th class="p-2 text-right">เบี้ยขับ</th><th class="p-2 text-right">หักเงิน</th><th class="p-2 text-right">สุทธิโอน</th></tr></thead>
                    <tbody>${rows}</tbody>
                    <tfoot><tr class="bg-slate-800 font-bold"><td class="p-2" colspan="2">รวม</td><td class="p-2 text-right">${tot.count}</td><td class="p-2 text-right">${tot.revenue.toLocaleString()}</td><td class="p-2 text-right">${tot.allowance.toLocaleString()}</td><td class="p-2 text-right text-red-400">${tot.deduction.toLocaleString()}</td><td class="p-2 text-right text-green-400">${tot.net.toLocaleString()}</td></tr></tfoot>
                </table>
                ${list.length === 0 ? '<p class="text-center py-8 text-gray-500">ยังไม่มีข้อมูล — เลือกเดือนที่นำเข้าคิวรถแล้ว</p>' : ''}
                </div>`;
        };
        const attachPayrollListeners = () => {
            const monthInput = document.getElementById('payroll-month');
            const target = appState.tripsMonth || (monthInput ? monthInput.value : '');
            if (target && appState.tripsMonth !== target) logisticsService.loadTripsForMonth(target);
            monthInput?.addEventListener('change', (e) => logisticsService.loadTripsForMonth(e.target.value));
            document.getElementById('payroll-export-btn')?.addEventListener('click', handlePayrollExport);
        };
        const handlePayrollExport = () => {
            const { list, bankOf } = aggregatePayroll();
            if (!list.length) { alert('ไม่มีข้อมูลให้ export'); return; }
            const headers = ['คนขับ', 'เลขบัญชี', 'จำนวนคิว', 'รายรับ', 'เบี้ยขับ', 'หักเงิน', 'สุทธิโอน'];
            const data = list.map(g => [g.driver, bankOf[g.driver.trim()] || '', g.count, g.revenue, g.allowance, g.deduction, g.net]);
            const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Payroll');
            XLSX.writeFile(wb, `Payroll_${appState.tripsMonth || localDateStr().slice(0, 7)}.xlsx`);
        };

        // --- View: Reports (รายงานสรุป) — summarizes the loaded month's trips ---
        const renderReports = () => {
            const month = appState.tripsMonth || localDateStr().slice(0, 7);
            return `
                <div class="space-y-6">
                    ${renderCard(`
                        <div class="flex flex-wrap justify-between items-end gap-4 mb-4">
                            <h2 class="text-2xl font-bold text-white flex items-center gap-3">${ICONS.ChartBar} รายงานสรุป</h2>
                            <div><label class="block text-xs text-gray-400 mb-1">เดือน</label>
                                <input type="month" id="reports-month" value="${month}" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>
                        </div>
                        <div id="reports-content">${renderReportsContent()}</div>
                    `)}
                </div>`;
        };
        const renderReportsContent = () => {
            const rs = appState.reportsState;
            const allTrips = logisticsService.getTrips();
            if (!allTrips.length) return `<p class="text-center py-8 text-gray-500">ไม่มีข้อมูล — เลือกเดือนที่นำเข้าคิวรถแล้ว</p>`;

            // plate → own/partner classification
            const driveOf = {}; logisticsService.getVehicles().forEach(v => { driveOf[(v.registrationNumber || '').trim()] = v.driveType || ''; });
            const driveClass = (t) => { const d = driveOf[(t.plate || '').trim()]; return d && d.includes('ร่วม') ? 'รถร่วม' : d ? 'รถบริษัท' : 'ไม่ทราบ'; };

            // E) Filters (options built from ALL trips so they never disappear)
            const dirOptions = [...new Set(allTrips.map(t => t.direction).filter(Boolean))];
            let trips = allTrips;
            if (rs.fDirection) trips = trips.filter(t => (t.direction || '') === rs.fDirection);
            if (rs.fDrive) trips = trips.filter(t => driveClass(t) === rs.fDrive);
            if (rs.fDriver) { const q = rs.fDriver.toLowerCase(); trips = trips.filter(t => (t.driver || '').toLowerCase().includes(q)); }

            const viewBtn = (val, label) => `<button data-report-view="${val}" class="px-3 py-1.5 rounded-md text-sm ${rs.view === val ? 'bg-accent text-white' : 'text-gray-300 hover:bg-slate-600'}">${label}</button>`;
            const filterBar = `
                <div class="flex flex-wrap items-end justify-between gap-3 mb-4">
                    <div class="flex flex-wrap items-end gap-3">
                        <div><label class="block text-xs text-gray-400 mb-1">ทิศทาง</label>
                            <select id="rep-f-direction" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white text-sm">
                                <option value="">ทั้งหมด</option>${dirOptions.map(d => `<option value="${d}" ${rs.fDirection === d ? 'selected' : ''}>${d}</option>`).join('')}
                            </select></div>
                        <div><label class="block text-xs text-gray-400 mb-1">ประเภทรถ</label>
                            <select id="rep-f-drive" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white text-sm">
                                <option value="">ทั้งหมด</option>
                                <option value="รถบริษัท" ${rs.fDrive === 'รถบริษัท' ? 'selected' : ''}>รถบริษัท</option>
                                <option value="รถร่วม" ${rs.fDrive === 'รถร่วม' ? 'selected' : ''}>รถร่วม</option>
                            </select></div>
                        <div><label class="block text-xs text-gray-400 mb-1">ค้นหาคนขับ</label>
                            <input id="rep-f-driver" value="${rs.fDriver}" placeholder="ชื่อคนขับ" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white text-sm w-40" /></div>
                    </div>
                    <div class="flex items-center gap-1 bg-slate-800 rounded-lg p-1">${viewBtn('both', 'ทั้งคู่')}${viewBtn('charts', 'กราฟ')}${viewBtn('tables', 'ตาราง')}</div>
                </div>`;

            if (!trips.length) return `${filterBar}<p class="text-center py-8 text-gray-500">ไม่มีคิวตรงตามตัวกรอง — ลองล้างตัวกรอง</p>`;

            const sum = (arr, f) => arr.reduce((s, t) => s + (f(t) || 0), 0);
            const totLiters = sum(trips, t => t.totalLiters), totBaht = sum(trips, t => t.totalBaht);
            const totRevenue = sum(trips, t => t.revenue), totNet = sum(trips, t => t.netTransfer);
            const fmt = (n) => (n || 0).toLocaleString();

            // group helper → [{key, count, net}] sorted by net desc
            const groupBy = (keyFn) => {
                const m = {};
                trips.forEach(t => { const k = keyFn(t) || '(ไม่ระบุ)'; if (!m[k]) m[k] = { key: k, count: 0, net: 0 }; m[k].count++; m[k].net += t.netTransfer || 0; });
                return Object.values(m).sort((a, b) => b.net - a.net);
            };
            const byDir = groupBy(t => t.direction);
            const byType = groupBy(t => t.vehicleType);
            const topDrivers = groupBy(t => t.driver).slice(0, 10);
            const ownPartner = { 'รถบริษัท': { count: 0, net: 0 }, 'รถร่วม': { count: 0, net: 0 }, 'ไม่ทราบ': { count: 0, net: 0 } };
            trips.forEach(t => { const k = driveClass(t); ownPartner[k].count++; ownPartner[k].net += t.netTransfer || 0; });

            const bars = (rows, maxKeyWidth = 'w-28') => {
                const max = Math.max(1, ...rows.map(r => r.net));
                return rows.map(r => `
                    <div class="flex items-center gap-3">
                        <span class="${maxKeyWidth} shrink-0 text-sm text-gray-300 truncate" title="${r.key}">${r.key}</span>
                        <div class="flex-1 bg-slate-800 rounded-full h-5 overflow-hidden"><div class="bg-accent h-full rounded-full" style="width:${Math.max(2, Math.round((r.net / max) * 100))}%"></div></div>
                        <span class="w-16 shrink-0 text-right text-xs text-gray-400">${r.count} คิว</span>
                        <span class="w-24 shrink-0 text-right text-sm font-semibold text-accent">${r.net.toLocaleString()}</span>
                    </div>`).join('');
            };

            const stationTable = reportTable('station', [
                { label: 'ปั๊ม', key: 'name', fmt: r => r.name },
                { label: 'ครั้ง', key: 'count' },
                { label: 'ลิตร', key: 'liters', cls: 'text-accent' },
                { label: 'บาท', key: 'baht' },
            ], computeReportStationRows(trips), { key: 'liters', dir: 'desc' });

            const dailyRows = computeReportDailyRows(trips);
            const dailyTable = reportTable('daily', [
                { label: 'วันที่', key: 'date', fmt: r => r.date },
                { label: 'คิว', key: 'count' },
                { label: 'ลิตร', key: 'liters', cls: 'text-accent' },
                { label: 'บาท', key: 'baht' },
                { label: 'รายรับ', key: 'revenue' },
                { label: 'สุทธิโอน', key: 'net', cls: 'text-green-400' },
            ], dailyRows, { key: 'date', dir: 'asc' },
                ['รวม (' + dailyRows.length + ' วัน)', fmt(trips.length), fmt(totLiters), fmt(totBaht), fmt(totRevenue), fmt(totNet)]);

            const routeTable = reportTable('route', [
                { label: 'เส้นทาง', key: 'route', fmt: r => r.route },
                { label: 'เที่ยว', key: 'count' },
                { label: 'ลิตรรวม', key: 'liters', cls: 'text-accent' },
                { label: 'ลิตร/เที่ยว', key: 'perTrip', sortVal: r => r.count ? r.liters / r.count : 0, fmt: r => (r.count ? r.liters / r.count : 0).toFixed(1) },
                { label: 'บาท', key: 'baht' },
            ], computeReportRouteRows(trips), { key: 'count', dir: 'desc' });

            const driverTable = reportTable('driver', [
                { label: 'คนขับ', key: 'driver', fmt: r => r.driver },
                { label: 'คิว', key: 'count' },
                { label: 'ลิตร', key: 'liters', cls: 'text-accent' },
                { label: 'รายรับ', key: 'revenue' },
                { label: 'สุทธิโอน', key: 'net', cls: 'text-green-400' },
            ], computeReportDriverRows(trips), { key: 'net', dir: 'desc' });

            const showCharts = rs.view !== 'tables';
            const showTables = rs.view !== 'charts';

            const chartsSection = `
                ${renderCard(`<h3 class="font-bold text-white mb-3 flex items-center gap-2">${ICONS.ChartBar} แนวโน้มน้ำมันรายวัน (ลิตร)</h3><div class="overflow-x-auto">${renderReportTrend(dailyRows)}</div>`)}
                <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 mt-6">
                    ${renderCard(`<h3 class="font-bold text-white mb-3">ตามทิศทาง</h3><div class="space-y-2">${bars(byDir, 'w-20')}</div>`)}
                    ${renderCard(`<h3 class="font-bold text-white mb-3">รถบริษัท vs รถร่วม</h3><div class="space-y-2">${bars(Object.entries(ownPartner).filter(([, v]) => v.count).map(([k, v]) => ({ key: k, count: v.count, net: v.net })), 'w-20')}</div>`)}
                    ${renderCard(`<h3 class="font-bold text-white mb-3">ตามประเภทรถ</h3><div class="space-y-2">${bars(byType, 'w-28')}</div>`)}
                    ${renderCard(`<h3 class="font-bold text-white mb-3">คนขับสุทธิโอนสูงสุด (Top 10)</h3><div class="space-y-2">${bars(topDrivers, 'w-32')}</div>`)}
                </div>`;

            const tablesSection = `
                <div class="mt-6">
                    ${renderCard(`
                        <div class="flex flex-wrap justify-between items-center gap-3 mb-3">
                            <h3 class="font-bold text-white flex items-center gap-2">${ICONS.ChartBar} ตารางสรุปรายวัน</h3>
                            <button id="reports-export-btn" class="flex items-center gap-2 bg-success hover:bg-green-600 text-white font-bold py-2 px-4 rounded-lg shadow-md">${ICONS.Download} Export Excel (ทุกตาราง)</button>
                        </div>${dailyTable}`)}
                </div>
                <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 mt-6">
                    ${renderCard(`<h3 class="font-bold text-white mb-3 flex items-center gap-2">${ICONS.ChartBar} รายเส้นทาง</h3>${routeTable}`)}
                    ${renderCard(`<h3 class="font-bold text-white mb-3 flex items-center gap-2">${ICONS.IdCard} รายคนขับ</h3>${driverTable}`)}
                    ${renderCard(`<h3 class="font-bold text-white mb-3 flex items-center gap-2">⛽ รายปั๊ม</h3>${stationTable}`)}
                </div>`;

            return `
                ${filterBar}
                <p class="text-xs text-gray-500 mb-3">เดือน ${appState.tripsMonth || '(เลือกเดือน)'} · แสดง ${trips.length.toLocaleString()} จาก ${allTrips.length.toLocaleString()} คิว · 💡 คลิกหัวคอลัมน์เพื่อเรียงลำดับ</p>
                <div class="grid grid-cols-2 md:grid-cols-5 gap-3 mb-6">
                    ${importStatCard('จำนวนคิว', trips.length)}
                    ${importStatCard('น้ำมัน (ลิตร)', totLiters)}
                    ${importStatCard('น้ำมัน (บาท)', totBaht)}
                    ${importStatCard('รายรับรวม', totRevenue)}
                    ${importStatCard('สุทธิโอนรวม', totNet)}
                </div>
                ${showCharts ? chartsSection : ''}
                ${showTables ? tablesSection : ''}
                <div class="mt-6">${renderReportCompare()}</div>`;
        };

        // D) Simple SVG line chart of daily liters (no chart lib).
        const renderReportTrend = (rows) => {
            if (!rows || rows.length < 2) return '<p class="text-gray-500 text-sm">ข้อมูลไม่พอวาดกราฟ (ต้องมีอย่างน้อย 2 วัน)</p>';
            const W = 760, H = 200, pad = 34;
            const max = Math.max(1, ...rows.map(r => r.liters));
            const stepX = (W - pad * 2) / (rows.length - 1);
            const pts = rows.map((r, i) => [pad + i * stepX, H - pad - (r.liters / max) * (H - pad * 2)]);
            const line = pts.map((p, i) => (i ? 'L' : 'M') + p[0].toFixed(1) + ' ' + p[1].toFixed(1)).join(' ');
            const area = `M${pad} ${H - pad} ` + pts.map(p => 'L' + p[0].toFixed(1) + ' ' + p[1].toFixed(1)).join(' ') + ` L${(pad + (rows.length - 1) * stepX).toFixed(1)} ${H - pad} Z`;
            const dots = pts.map((p, i) => `<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="2.5" fill="#f97316"><title>${rows[i].date}: ${rows[i].liters.toLocaleString()} ลิตร (${rows[i].count} คิว)</title></circle>`).join('');
            const xlabels = [0, Math.floor((rows.length - 1) / 2), rows.length - 1].map(i => `<text x="${(pad + i * stepX).toFixed(1)}" y="${H - 8}" fill="#94a3b8" font-size="11" text-anchor="middle">${(rows[i].date || '').slice(5)}</text>`).join('');
            return `<svg viewBox="0 0 ${W} ${H}" class="w-full" style="min-width:640px" preserveAspectRatio="xMidYMid meet">
                <text x="${pad}" y="16" fill="#94a3b8" font-size="11">สูงสุด ${max.toLocaleString()} ล.</text>
                <path d="${area}" fill="#f9731622" /><path d="${line}" fill="none" stroke="#f97316" stroke-width="2" />${dots}${xlabels}
            </svg>`;
        };

        // C) Cross-month comparison section (loads on demand via getMonthlySummaries).
        const renderReportCompare = () => {
            const c = appState.reportsState.compare;
            const controls = `
                <div class="flex flex-wrap items-end gap-3 mb-3">
                    <div><label class="block text-xs text-gray-400 mb-1">ตั้งแต่เดือน</label>
                        <input type="month" id="cmp-from" value="${c.from}" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white text-sm" /></div>
                    <div><label class="block text-xs text-gray-400 mb-1">ถึงเดือน</label>
                        <input type="month" id="cmp-to" value="${c.to}" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white text-sm" /></div>
                    <button id="cmp-load" class="bg-primary hover:bg-secondary text-white font-bold py-2 px-4 rounded-lg text-sm">เปรียบเทียบ</button>
                </div>`;
            let body;
            if (c.loading) body = '<p class="text-gray-400 text-sm">กำลังโหลด...</p>';
            else if (c.data && c.data.length) {
                const maxL = Math.max(1, ...c.data.map(d => d.liters));
                const barsHtml = c.data.map(d => `
                    <div class="flex items-center gap-3">
                        <span class="w-20 shrink-0 text-sm text-gray-300">${d.month}</span>
                        <div class="flex-1 bg-slate-800 rounded-full h-5 overflow-hidden"><div class="bg-accent h-full rounded-full" style="width:${Math.max(2, Math.round(d.liters / maxL * 100))}%"></div></div>
                        <span class="w-24 shrink-0 text-right text-sm text-accent font-mono">${d.liters.toLocaleString()} ล.</span>
                    </div>`).join('');
                const tbl = reportTable('cmp', [
                    { label: 'เดือน', key: 'month', fmt: r => r.month },
                    { label: 'คิว', key: 'count' },
                    { label: 'ลิตร', key: 'liters', cls: 'text-accent' },
                    { label: 'บาท', key: 'baht' },
                    { label: 'รายรับ', key: 'revenue' },
                    { label: 'สุทธิโอน', key: 'net', cls: 'text-green-400' },
                ], c.data, { key: 'month', dir: 'asc' });
                body = `<div class="space-y-2 mb-4">${barsHtml}</div>${tbl}`;
            } else if (c.data) body = '<p class="text-gray-500 text-sm">ไม่มีข้อมูลในช่วงที่เลือก</p>';
            else body = '<p class="text-gray-500 text-sm">เลือกช่วงเดือนแล้วกด "เปรียบเทียบ" เพื่อดูแนวโน้มข้ามเดือน</p>';
            return renderCard(`<h3 class="font-bold text-white mb-3 flex items-center gap-2">${ICONS.ChartBar} เปรียบเทียบรายเดือน</h3>${controls}<div id="reports-compare-body">${body}</div>`);
        };

        // Group a month's trips into per-day report rows (date × count/liters/baht/revenue/net), sorted by date.
        const computeReportDailyRows = (trips) => {
            const m = {};
            trips.forEach(t => {
                const d = t.date || '(ไม่ระบุ)';
                if (!m[d]) m[d] = { date: d, count: 0, liters: 0, baht: 0, revenue: 0, net: 0 };
                const r = m[d];
                r.count++; r.liters += t.totalLiters || 0; r.baht += t.totalBaht || 0;
                r.revenue += t.revenue || 0; r.net += t.netTransfer || 0;
            });
            return Object.values(m).sort((a, b) => a.date.localeCompare(b.date));
        };

        // Per-route rows (origin → destination × trips/liters/baht), sorted by trip count desc.
        const computeReportRouteRows = (trips) => {
            const m = {};
            trips.forEach(t => {
                const k = `${t.origin || '?'} → ${t.destination || '?'}`;
                if (!m[k]) m[k] = { route: k, count: 0, liters: 0, baht: 0 };
                const r = m[k]; r.count++; r.liters += t.totalLiters || 0; r.baht += t.totalBaht || 0;
            });
            return Object.values(m).sort((a, b) => b.count - a.count);
        };

        // Per-station rows (times refueled / liters / baht), sorted by liters desc.
        const computeReportStationRows = (trips) => {
            const m = {};
            const touch = (k) => { if (!m[k]) m[k] = { key: k, count: 0, liters: 0, baht: 0 }; return m[k]; };
            trips.forEach(t => {
                new Set([...Object.keys(t.gasStationsLiters || {}), ...Object.keys(t.gasStationsBaht || {})]).forEach(k => touch(k).count++);
                Object.entries(t.gasStationsLiters || {}).forEach(([k, v]) => { touch(k).liters += v; });
                Object.entries(t.gasStationsBaht || {}).forEach(([k, v]) => { touch(k).baht += v; });
            });
            return Object.values(m).map(r => ({ ...r, name: logisticsService.getFriendlyGasStationName(r.key) })).sort((a, b) => b.liters - a.liters);
        };

        // Per-driver rows (queues/liters/revenue/allowance/net), sorted by net desc.
        const computeReportDriverRows = (trips) => {
            const m = {};
            trips.forEach(t => {
                const k = t.driver || '(ไม่ระบุ)';
                if (!m[k]) m[k] = { driver: k, count: 0, liters: 0, revenue: 0, allowance: 0, net: 0 };
                const r = m[k]; r.count++; r.liters += t.totalLiters || 0;
                r.revenue += t.revenue || 0; r.allowance += t.driverAllowance || 0; r.net += t.netTransfer || 0;
            });
            return Object.values(m).sort((a, b) => b.net - a.net);
        };

        // Sort rows by the active sort state for a table (falls back to the provided default).
        const sortReportRows = (tableId, headers, rows, def) => {
            const s = appState.reportsState.sorts[tableId] || def;
            if (!s || !s.key) return rows;
            const h = headers.find(x => x.key === s.key);
            const val = (r) => (h && h.sortVal) ? h.sortVal(r) : r[s.key];
            const dir = s.dir === 'asc' ? 1 : -1;
            return [...rows].sort((a, b) => {
                const va = val(a), vb = val(b);
                if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
                return String(va).localeCompare(String(vb), 'th') * dir;
            });
        };

        // Generic SORTABLE report-table renderer.
        //   tableId: unique id used for sort state + click routing
        //   headers: [{label, key, fmt?(row), cls?, sortVal?(row)}] — first column is left-aligned text
        //   def: default sort {key, dir}; totalRow: optional array of pre-formatted footer cells
        const reportTable = (tableId, headers, rows, def, totalRow) => {
            const s = appState.reportsState.sorts[tableId] || def || {};
            const sorted = sortReportRows(tableId, headers, rows, def);
            const arrow = (h) => s.key === h.key ? (s.dir === 'asc' ? ' ▲' : ' ▼') : '';
            return `
            <div class="overflow-x-auto max-h-[440px] rounded-lg border border-slate-700">
                <table class="w-full text-sm text-left">
                    <thead class="bg-slate-800 text-gray-400 uppercase sticky top-0 z-10"><tr>
                        ${headers.map((h, i) => `<th data-sort-table="${tableId}" data-sort-key="${h.key}" class="p-2 ${i === 0 ? '' : 'text-right'} whitespace-nowrap cursor-pointer select-none hover:text-white">${h.label}${arrow(h)}</th>`).join('')}
                    </tr></thead>
                    <tbody>
                        ${sorted.length ? sorted.map(r => `<tr class="border-b border-slate-700/50 hover:bg-slate-700/20">
                            ${headers.map((h, i) => `<td class="p-2 ${i === 0 ? 'whitespace-nowrap' : 'text-right font-mono'} ${h.cls || ''}">${h.fmt ? h.fmt(r) : (r[h.key] || 0).toLocaleString()}</td>`).join('')}
                        </tr>`).join('') : `<tr><td colspan="${headers.length}" class="text-center py-6 text-gray-500">ไม่มีข้อมูล</td></tr>`}
                    </tbody>
                    ${totalRow ? `<tfoot class="bg-slate-800 font-bold text-white sticky bottom-0"><tr>${totalRow.map((c, i) => `<td class="p-2 ${i === 0 ? '' : 'text-right font-mono'}">${c}</td>`).join('')}</tr></tfoot>` : ''}
                </table>
            </div>`;
        };

        const handleReportsExport = () => {
            const trips = logisticsService.getTrips();
            if (!trips.length) { alert('ไม่มีข้อมูลให้ export'); return; }
            const total = (rows, f) => rows.reduce((s, r) => s + (r[f] || 0), 0);
            const wb = XLSX.utils.book_new();
            const addSheet = (name, header, rows, cols, totalRow) => {
                const aoa = [header, ...rows];
                if (totalRow) aoa.push(totalRow);
                const ws = XLSX.utils.aoa_to_sheet(aoa);
                if (cols) ws['!cols'] = cols;
                XLSX.utils.book_append_sheet(wb, ws, name);
            };

            // 1) รายวัน
            const daily = computeReportDailyRows(trips);
            addSheet('สรุปรายวัน',
                ['วันที่', 'จำนวนคิว', 'น้ำมัน (ลิตร)', 'น้ำมัน (บาท)', 'รายรับ', 'สุทธิโอน'],
                daily.map(r => [r.date, r.count, r.liters, r.baht, r.revenue, r.net]),
                [{ wch: 14 }, { wch: 10 }, { wch: 13 }, { wch: 13 }, { wch: 12 }, { wch: 12 }],
                ['รวม', trips.length, total(daily, 'liters'), total(daily, 'baht'), total(daily, 'revenue'), total(daily, 'net')]);

            // 2) รายเส้นทาง
            const routes = computeReportRouteRows(trips);
            addSheet('รายเส้นทาง',
                ['เส้นทาง', 'เที่ยว', 'ลิตรรวม', 'ลิตร/เที่ยว', 'บาท'],
                routes.map(r => [r.route, r.count, r.liters, +(r.count ? r.liters / r.count : 0).toFixed(1), r.baht]),
                [{ wch: 40 }, { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 12 }]);

            // 3) รายปั๊ม
            const stations = computeReportStationRows(trips);
            addSheet('รายปั๊ม',
                ['ปั๊ม', 'จำนวนครั้ง', 'ลิตร', 'บาท'],
                stations.map(r => [r.name, r.count, r.liters, r.baht]),
                [{ wch: 28 }, { wch: 12 }, { wch: 12 }, { wch: 12 }]);

            // 4) รายคนขับ
            const drivers = computeReportDriverRows(trips);
            addSheet('รายคนขับ',
                ['คนขับ', 'คิว', 'ลิตร', 'รายรับ', 'เบี้ยเลี้ยง', 'สุทธิโอน'],
                drivers.map(r => [r.driver, r.count, r.liters, r.revenue, r.allowance, r.net]),
                [{ wch: 22 }, { wch: 8 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 12 }]);

            XLSX.writeFile(wb, `รายงานสรุป_${appState.tripsMonth || localDateStr().slice(0, 7)}.xlsx`);
        };
        // Enumerate 'YYYY-MM' months inclusive between two (capped so a bad range can't loop forever).
        const monthRange = (from, to) => {
            const out = [];
            let [y, m] = from.split('-').map(Number);
            const [ty, tm] = to.split('-').map(Number);
            if (!y || !ty || y > ty || (y === ty && m > tm)) return [];
            while ((y < ty || (y === ty && m <= tm)) && out.length < 36) {
                out.push(`${y}-${String(m).padStart(2, '0')}`);
                m++; if (m > 12) { m = 1; y++; }
            }
            return out;
        };

        const handleReportsCompare = async () => {
            const from = document.getElementById('cmp-from')?.value;
            const to = document.getElementById('cmp-to')?.value;
            if (!from || !to) { alert('เลือกช่วงเดือน (ตั้งแต่/ถึง) ก่อน'); return; }
            const months = monthRange(from, to);
            if (!months.length) { alert('ช่วงเดือนไม่ถูกต้อง (เดือนเริ่มต้องไม่เกินเดือนสิ้นสุด)'); return; }
            const rs = appState.reportsState;
            rs.compare = { from, to, data: null, loading: true };
            rerenderReportsContent();
            try {
                const data = await logisticsService.getMonthlySummaries(months);
                rs.compare = { from, to, data, loading: false };
            } catch (err) {
                console.error('Compare load failed:', err);
                alert('โหลดข้อมูลเปรียบเทียบไม่สำเร็จ: ' + (err?.message || err));
                rs.compare = { from, to, data: [], loading: false };
            }
            rerenderReportsContent();
        };

        // Re-render only the reports content region + re-bind its inner controls (fast; no full app render).
        const rerenderReportsContent = () => {
            const c = document.getElementById('reports-content');
            if (!c) return;
            c.innerHTML = renderReportsContent();
            attachReportsContentListeners();
        };

        // Binds all controls INSIDE #reports-content (filters, view toggle, sort headers, export, compare).
        const attachReportsContentListeners = () => {
            const rs = appState.reportsState;
            document.getElementById('rep-f-direction')?.addEventListener('change', (e) => { rs.fDirection = e.target.value; rerenderReportsContent(); });
            document.getElementById('rep-f-drive')?.addEventListener('change', (e) => { rs.fDrive = e.target.value; rerenderReportsContent(); });
            document.getElementById('rep-f-driver')?.addEventListener('input', (e) => {
                rs.fDriver = e.target.value; rerenderReportsContent();
                const el = document.getElementById('rep-f-driver'); if (el) { el.focus(); el.setSelectionRange(el.value.length, el.value.length); }
            });
            document.querySelectorAll('[data-report-view]').forEach(b => b.addEventListener('click', () => { rs.view = b.dataset.reportView; rerenderReportsContent(); }));
            document.querySelectorAll('[data-sort-table]').forEach(th => th.addEventListener('click', () => {
                const id = th.dataset.sortTable, key = th.dataset.sortKey;
                const cur = rs.sorts[id];
                rs.sorts[id] = (cur && cur.key === key) ? { key, dir: cur.dir === 'asc' ? 'desc' : 'asc' } : { key, dir: 'desc' };
                rerenderReportsContent();
            }));
            document.getElementById('reports-export-btn')?.addEventListener('click', handleReportsExport);
            document.getElementById('cmp-load')?.addEventListener('click', handleReportsCompare);
        };

        // --- View: Fuel audit (ตรวจการเติม) — actual fills vs the official rate sheet ---
        // Audit results only change when trips, rates, vehicles or tolerance change — not on every search keystroke.
        let auditCache = {};
        // Earlier months' trips, read once (no live listener) for the liters carried over from them.
        // Cleared by refreshAuditPast() whenever the audit page is opened, the month / carry start changes,
        // or the user presses "โหลดยอดยกใหม่", so edits made in earlier months are picked up.
        let auditPastTrips = {};
        let auditPastFailed = new Set();
        let auditPastGen = 0; // ignores answers from a fetch started before the last refresh
        const auditPastLoading = new Set();
        let auditPastResults = { key: null, results: [] };
        const refreshAuditPast = () => { auditPastTrips = {}; auditPastFailed = new Set(); auditPastLoading.clear(); auditPastGen++; auditPastResults = { key: null, results: [] }; };
        const prevYm = (ym) => { let [y, m] = ym.split('-').map(Number); if (--m < 1) { m = 12; y--; } return `${y}-${String(m).padStart(2, '0')}`; };
        const MAX_CARRY_MONTHS = 36; // monthRange's cap
        const loadPastMonths = (months) => {
            const todo = months.filter(m => !(m in auditPastTrips) && !auditPastFailed.has(m) && !auditPastLoading.has(m));
            if (!todo.length) return;
            const gen = auditPastGen;
            todo.forEach(m => auditPastLoading.add(m));
            Promise.all(todo.map(m => logisticsService.fetchTripsMonth(m)
                .then(t => { if (gen === auditPastGen) auditPastTrips[m] = t; })
                .catch(err => { console.error(`Load trips ${m} for carry-over failed:`, err); if (gen === auditPastGen) auditPastFailed.add(m); })
                .finally(() => { if (gen === auditPastGen) auditPastLoading.delete(m); })))
                .then(() => { if (gen === auditPastGen && appState.view === 'audit') rerenderAuditContent(); });
        };
        // Column L downhill liters (downhillFuel/{ym}) — fetched once per month the audit shows, never the whole node.
        const dfuelLoaded = new Set(), dfuelLoading = new Set();
        const loadDownhillFuel = (months) => {
            const todo = months.filter(m => /^\d{4}-\d{2}$/.test(m || '') && !dfuelLoaded.has(m) && !dfuelLoading.has(m));
            if (!todo.length) return;
            todo.forEach(m => dfuelLoading.add(m));
            Promise.all(todo.map(m => logisticsService.fetchDownhillFuelMonth(m)
                .then(v => { dfuelLoaded.add(m); return v; })
                .catch(err => { console.error(`Load downhillFuel ${m} failed:`, err); return {}; })
                .finally(() => dfuelLoading.delete(m))))
                .then(parts => {
                    const merged = Object.assign({}, appState.downhillFuel, ...parts);
                    if (Object.keys(merged).length !== Object.keys(appState.downhillFuel).length) {
                        appState.downhillFuel = merged; // new object → the audit cache recomputes
                        if (appState.view === 'audit') rerenderAuditContent();
                    }
                });
        };
        const resetDownhillFuel = () => { dfuelLoaded.clear(); appState.downhillFuel = {}; };
        // Trip order inside a day: board row number when present, else the record id.
        const tripOrder = (a, b) => (a.trip.date || '').localeCompare(b.trip.date || '') || ((a.leg === 'down') - (b.leg === 'down')) || ((a.trip.sourceRow || 0) - (b.trip.sourceRow || 0)) || String(a.trip.id || '').localeCompare(String(b.trip.id || ''));

        const computeAudit = () => {
            const as = appState.auditState;
            const tol = Math.max(0, Number(as.tolerance) || 0);
            const month = appState.tripsMonth;
            const carryOn = /^\d{4}-\d{2}$/.test(as.carryStart || '') && /^\d{4}-\d{2}$/.test(month || '') && month >= as.carryStart;
            const pastAll = carryOn && as.carryStart < month ? monthRange(as.carryStart, prevYm(month)) : [];
            const carryTooLong = carryOn && as.carryStart < month && pastAll[pastAll.length - 1] !== prevYm(month);
            const pastMonths = carryTooLong ? [] : pastAll;
            loadPastMonths(pastMonths);
            const carryFailed = pastMonths.filter(m => auditPastFailed.has(m));
            const carryPending = !carryFailed.length && pastMonths.some(m => !(m in auditPastTrips));
            const carryReady = carryOn && !carryTooLong && !carryPending && !carryFailed.length;
            const pastKey = carryReady ? pastMonths.map(m => `${m}:${auditPastTrips[m].length}`).join(',') + `|${auditPastGen}` : 'none';
            loadDownhillFuel([...pastMonths, month]);
            const dfuel = appState.downhillFuel;
            const ox = as.originExtras;
            // Board short names → full names (settings/fuelAudit.driverAliases, confirmed by the user): the audit,
            // carry-over and driver summary treat both spellings as one driver. Stored trips are not changed.
            const dal = as.driverAliases;
            const aliasTrip = (t) => applyDriverAlias(t, dal);
            if (auditCache.trips !== appState.trips || auditCache.std !== appState.rateStandards || auditCache.veh !== appState.vehicles || auditCache.tol !== tol || auditCache.pastKey !== pastKey || auditCache.dfuel !== dfuel || auditCache.ox !== ox || auditCache.dal !== dal) {
                const ctx = buildAuditContext(appState.rateStandards, appState.vehicles, { originExtras: ox });
                // Downhill trips without their own liters take the column L fill (downhillFuel); the trip itself is untouched.
                const run = (trips) => (trips || []).map(t => {
                    const r = auditTrip(withDownhillFuel(aliasTrip(t), dfuel[t.id]), ctx, tol);
                    if (r && r.trip.fuelFrom === 'colL') r.flags = [...r.flags, `ลิตรจากช่อง "คิวรถล่าสุด" ไฟล์ Transport: ${r.trip.fuelText}`];
                    return r;
                }).filter(Boolean);
                const all = run(appState.trips).sort(tripOrder);
                let balances = {};
                if (carryReady) {
                    // Earlier months do not change with the live month — audit them once per (data, rates, vehicles, tolerance).
                    if (auditPastResults.key !== pastKey || auditPastResults.std !== appState.rateStandards || auditPastResults.veh !== appState.vehicles || auditPastResults.tol !== tol || auditPastResults.dfuel !== dfuel || auditPastResults.ox !== ox || auditPastResults.dal !== dal) {
                        auditPastResults = { key: pastKey, std: appState.rateStandards, veh: appState.vehicles, tol, dfuel, ox, dal, results: pastMonths.flatMap(m => run(auditPastTrips[m])) };
                    }
                    const ordered = [...auditPastResults.results.map(r => ({ ...r })), ...all].sort(tripOrder);
                    reconcileColumnL(ordered, new Set([...pastMonths, month]));
                    flagDoubleDownhill(ordered, [...pastMonths.flatMap(m => auditPastTrips[m] || []), ...appState.trips].map(aliasTrip));
                    balances = applyCarry(ordered);
                } else {
                    reconcileColumnL(all, new Set([month]));
                    flagDoubleDownhill(all, appState.trips.map(aliasTrip));
                }
                auditCache = { trips: appState.trips, std: appState.rateStandards, veh: appState.vehicles, tol, pastKey, dfuel, ox, dal, all, balances };
            }
            const { all, balances } = auditCache;
            const q = (as.search || '').trim().toLowerCase();
            const shown = all.filter(r => (!as.status || r.status === as.status) &&
                (!q || `${r.trip.driver} ${r.trip.driverBoard || ""} ${r.trip.plate} ${r.trip.destination} ${r.trip.origin}`.toLowerCase().includes(q)));
            const count = {}; all.forEach(r => { count[r.status] = (count[r.status] || 0) + 1; });
            const downCount = {}; all.forEach(r => { if (r.downStatus) downCount[r.downStatus] = (downCount[r.downStatus] || 0) + 1; });
            const downTrips = {}; all.forEach(r => { if (r.leg === 'down') downTrips[r.status] = (downTrips[r.status] || 0) + 1; });
            const downTripsTotal = Object.values(downTrips).reduce((a, b) => a + b, 0);
            const sumDiff = (st) => all.filter(r => r.status === st).reduce((a, r) => a + r.diff, 0);
            const byDriver = {};
            const newDriver = (name) => ({ driver: name, n: 0, checked: 0, ok: 0, over: 0, overL: 0, under: 0, underL: 0, wrong: 0, balance: carryReady && name in balances ? balances[name] : (carryReady ? 0 : null) });
            // Drivers still owing / owed liters from earlier months stay listed even with no trip this month.
            if (carryReady) for (const name of Object.keys(balances)) if (balances[name]) byDriver[name] = newDriver(name);
            for (const r of all) {
                const name = String(r.trip.driver || '').trim() || '(ไม่ระบุ)';
                const d = byDriver[name] ||= newDriver(name);
                d.n++;
                if (r.status !== 'no-rate') d.checked++;
                if (r.status === 'ok') d.ok++;
                if (r.status === 'over') { d.over++; d.overL += r.diff; }
                if (r.status === 'under') { d.under++; d.underL += r.diff; }
                if (r.status === 'wrong-pump') d.wrong++;
            }
            const drivers = Object.values(byDriver).sort((a, b) => b.overL - a.overL || b.over - a.over);
            return { all, shown, count, downCount, downTrips, downTripsTotal, overL: sumDiff('over'), underL: sumDiff('under'), drivers, tol, carryOn, carryReady, carryPending, carryFailed, carryTooLong, pastMonths };
        };

        const renderAudit = () => {
            const month = appState.tripsMonth || localDateStr().slice(0, 7);
            const meta = appState.rateStandardsMeta;
            return `
                <div class="space-y-6">
                    ${renderCard(`
                        <div class="flex flex-wrap justify-between items-end gap-4 mb-2">
                            <h2 class="text-2xl font-bold text-white flex items-center gap-3">${ICONS.Truck} ตรวจการเติมน้ำมัน</h2>
                            <div class="flex flex-wrap items-end gap-3">
                                <div><label class="block text-xs text-gray-400 mb-1">เดือน</label>
                                    <input type="month" id="audit-month" value="${month}" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>
                                <div><label class="block text-xs text-gray-400 mb-1">ยอมให้คลาดเคลื่อน (ลิตร)</label>
                                    <input type="number" min="0" step="1" id="audit-tolerance" value="${appState.auditState.tolerance}" class="w-28 bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>
                                <div><label class="block text-xs text-gray-400 mb-1" title="ลิตรที่คนขับเติมเกิน/ขาด จะยกไปหักในเที่ยวถัดไปที่มีอัตรา เริ่มนับจากเดือนนี้ (ว่าง = ไม่ยกยอด)">ยกยอดลิตรเริ่มเดือน</label>
                                    <input type="month" id="audit-carry-start" value="${esc(appState.auditState.carryStart)}" class="bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" /></div>
                                <button id="audit-export-btn" class="flex items-center gap-2 bg-emerald-700 hover:bg-emerald-800 text-white font-bold py-2 px-4 rounded-lg shadow-md whitespace-nowrap">${ICONS.Download} Export Excel</button>
                            </div>
                        </div>
                        <p class="text-xs text-gray-400">${meta
                            ? `อัตรามาตรฐาน: ${esc(meta.fileName)} · ${appState.rateStandards.length} อัตรา · บันทึกเมื่อ ${esc(meta.importedAt)}${meta.flagged ? ` · <span class="text-amber-300">${meta.flagged} แถวในไฟล์อัตรามีข้อสังเกต</span>` : ''}`
                            : '<span class="text-amber-300">⚠️ ยังไม่มีอัตรามาตรฐาน — นำเข้าไฟล์ "อัตราเติมน้ำมันประจำ ปั๊ม ต่างๆ" ที่แท็บ นำเข้า Excel ก่อน</span>'}</p>
                        <div id="audit-content" class="mt-4">${renderAuditContent()}</div>
                        <div id="audit-ox" class="mt-2">${renderOriginExtrasEditor()}</div>
                    `)}
                </div>`;
        };

        const renderAuditContent = () => {
            if (!appState.rateStandards.length) return '';
            if (!logisticsService.getTrips().length) return `<p class="text-center py-8 text-gray-500">ไม่มีคิวรถในเดือนนี้ — เลือกเดือนอื่น</p>`;
            const as = appState.auditState;
            const { all, shown, count, downCount, downTrips, downTripsTotal, overL, underL, drivers, tol, carryOn, carryReady, carryPending, carryFailed, carryTooLong, pastMonths } = computeAudit();
            const signed = (v) => (v > 0 ? '+' : '') + Number(v).toLocaleString();
            const refreshBtn = pastMonths.length ? ' <button id="audit-carry-refresh" class="text-sky-300 hover:text-sky-200 font-bold">↻ โหลดยอดยกใหม่</button>' : '';
            const carryNote = !carryOn
                ? 'ไม่ได้ยกยอดลิตร (ตั้ง "ยกยอดลิตรเริ่มเดือน" เพื่อเปิดใช้)'
                : carryTooLong ? `<span class="text-red-400">ยกยอดย้อนหลังได้ไม่เกิน ${MAX_CARRY_MONTHS} เดือน — เลื่อน "ยกยอดลิตรเริ่มเดือน" ให้ใกล้ขึ้น</span>`
                : carryFailed.length ? `<span class="text-red-400">โหลดคิวเดือน ${carryFailed.join(', ')} ไม่สำเร็จ — ยังคำนวณยอดยกไม่ได้</span>${refreshBtn}`
                : carryPending ? `กำลังโหลดคิว ${pastMonths.length} เดือนก่อนหน้าเพื่อคำนวณยอดยก…`
                : `ยกยอดลิตรตั้งแต่ ${esc(appState.auditState.carryStart)} (${pastMonths.length ? `รวม ${pastMonths.length} เดือนก่อนหน้า` : 'เริ่มเดือนนี้'}) · + = เติมเกินค้างไว้ หักจากเที่ยวถัดไป · − = เติมขาด ได้เพิ่มเที่ยวถัดไป · ไม่ยกยอดเที่ยวที่ไม่มีชื่อคนขับ${refreshBtn}`;
            const downCell = (r) => !r.downStatus ? '' :
                `<span class="${DOWN_STATUS[r.downStatus].cls}">${r.downStatus === 'missing' ? `อัตรา ${r.down} · ${DOWN_STATUS.missing.label}` : `${r.downActual}/${r.down} ${DOWN_STATUS[r.downStatus].label}`}</span>`;
            const carryCell = (r) => !carryReady ? '' : r.carryIn == null ? '<span class="text-gray-500">ไม่มีชื่อคนขับ</span>' :
                `<div class="whitespace-nowrap">ยกมา ${signed(r.carryIn)}</div>${r.allot != null && r.carryIn ? `<div class="text-xs text-gray-400 whitespace-nowrap">ควรเติม ${r.allot.toLocaleString()}</div>` : ''}<div class="whitespace-nowrap font-bold ${r.carryOut > 0 ? 'text-red-400' : r.carryOut < 0 ? 'text-amber-300' : 'text-gray-400'}">ยกไป ${signed(r.carryOut)}</div>`;
            const card = (status, label, value, sub = '') => `
                <button data-audit-status="${status}" class="text-left rounded-lg p-3 border ${as.status === status ? 'border-accent bg-slate-700' : 'border-slate-700 bg-slate-800 hover:bg-slate-700'}">
                    <div class="text-xs text-gray-400">${label}</div>
                    <div class="text-2xl font-bold text-white">${Number(value).toLocaleString()}</div>
                    ${sub ? `<div class="text-xs text-gray-400">${sub}</div>` : ''}
                </button>`;
            const badge = (st) => `<span class="px-2 py-0.5 rounded text-xs text-white whitespace-nowrap ${AUDIT_STATUS[st].cls}">${AUDIT_STATUS[st].label}</span>`;
            const LIMIT = 500;
            const rows = shown.slice(0, LIMIT).map(r => `
                <tr class="border-b border-slate-700 align-top">
                    <td class="p-2 whitespace-nowrap">${esc(r.trip.date)}</td>
                    <td class="p-2 whitespace-nowrap">${esc(r.trip.driver)}${r.trip.driverBoard ? `<div class="text-xs text-gray-500">กระดาน: ${esc(r.trip.driverBoard)}</div>` : ""}</td>
                    <td class="p-2 whitespace-nowrap font-mono text-xs">${esc(r.trip.plate)}</td>
                    <td class="p-2 whitespace-nowrap">${r.cls ? esc(AUDIT_CLASSES[r.cls]) : '-'}${r.rateCls && r.rateCls !== r.cls ? `<div class="text-xs text-sky-300">ตัดหาง → คิด${esc(AUDIT_CLASSES[r.rateCls])}</div>` : ''}${r.model ? ` <span class="text-xs text-indigo-300">${esc(r.model)}</span>` : ''}</td>
                    <td class="p-2">${r.leg === 'down' ? '<span class="px-1.5 py-0.5 rounded text-xs bg-teal-700 text-white">ขาล่อง</span> ' : ''}${esc(r.trip.origin)} → ${esc(r.trip.destination)}</td>
                    <td class="p-2">${fmtPumps(r.actual)}</td>
                    <td class="p-2">${r.status === 'no-rate' ? '-' : r.status === 'manual' ? (r.expectedTotal ? `รวม ${r.expectedTotal.toLocaleString()}` : '-') : fmtPumps(r.expected)}</td>
                    <td class="p-2 text-right font-bold ${r.diff > 0 ? 'text-red-400' : r.diff < 0 ? 'text-amber-300' : 'text-gray-400'}">${r.status === 'no-rate' || (r.status === 'manual' && !r.expectedTotal) ? '' : signed(r.diff)}</td>
                    <td class="p-2">${badge(r.status)}</td>
                    <td class="p-2 text-xs">${downCell(r)}</td>
                    <td class="p-2 text-xs">${carryCell(r)}</td>
                    <td class="p-2 text-xs text-gray-300">${[r.reason, ...r.flags].filter(Boolean).map(esc).join('<br>')}</td>
                </tr>`).join('');
            const driverRows = drivers.map(d => `
                <tr class="border-b border-slate-700">
                    <td class="p-2">${esc(d.driver)}</td><td class="p-2 text-right">${d.n}</td><td class="p-2 text-right">${d.checked}</td><td class="p-2 text-right">${d.ok}</td>
                    <td class="p-2 text-right text-red-400">${d.over}</td><td class="p-2 text-right text-red-400 font-bold">${d.overL ? '+' + d.overL.toLocaleString() : ''}</td>
                    <td class="p-2 text-right text-amber-300">${d.under}</td><td class="p-2 text-right text-amber-300">${d.underL ? d.underL.toLocaleString() : ''}</td>
                    <td class="p-2 text-right text-purple-300">${d.wrong}</td>
                    <td class="p-2 text-right font-bold ${d.balance > 0 ? 'text-red-400' : d.balance < 0 ? 'text-amber-300' : 'text-gray-400'}">${d.balance == null ? '-' : signed(d.balance)}</td>
                </tr>`).join('');
            const stdRows = appState.rateStandards.map(s => `
                <tr class="border-b border-slate-700 ${s.flags.length ? 'bg-amber-900/20' : ''}">
                    <td class="p-2 whitespace-nowrap">${esc(AUDIT_CLASSES[s.cls])}${s.variant === 'sino' ? ' (sino)' : ''}</td>
                    <td class="p-2 whitespace-nowrap">${esc(s.origin)} → ${esc(s.destinationLabel)}</td>
                    <td class="p-2">${fmtRateUp(s)}</td>
                    <td class="p-2 text-right">${s.down || '-'}${s.downVictor !== s.down ? ` (victor ${s.downVictor})` : ''}</td>
                    <td class="p-2 text-xs">${esc(s.allowanceText)}</td>
                    <td class="p-2 text-xs text-amber-300">${(s.flags || []).map(esc).join('<br>')}${s.duplicateOf ? '<br>(ไม่ใช้ตรวจ — ใช้แถวแรก)' : ''}</td>
                </tr>`).join('');
            return `
                <div class="grid grid-cols-2 md:grid-cols-4 xl:grid-cols-7 gap-3 mb-3">
                    ${card('', 'ทั้งหมด (เที่ยวที่มีการเติม)', all.length, downTripsTotal ? `ขาล่อง ${downTripsTotal} เที่ยว` : '')}
                    ${card('ok', 'ตรงตามอัตรา', count.ok || 0)}
                    ${card('over', 'เติมเกิน', count.over || 0, overL ? `รวม +${overL.toLocaleString()} ลิตร` : '')}
                    ${card('under', 'เติมขาด', count.under || 0, underL ? `รวม ${underL.toLocaleString()} ลิตร` : '')}
                    ${card('wrong-pump', 'ลิตรรวมถูก แต่ผิดปั๊ม', count['wrong-pump'] || 0)}
                    ${card('manual', 'ต้องตรวจเอง (ถอดหาง กพ. / เติมเพิ่มขาล่อง)', count.manual || 0)}
                    ${card('no-rate', 'ตรวจไม่ได้ (ไม่มีอัตรา)', count['no-rate'] || 0)}
                </div>
                <p class="text-xs text-gray-400 mb-1">ขาล่องไพศาลที่เขียนไว้ในเที่ยวขาขึ้น: ตรง ${downCount.ok || 0} · เกิน ${downCount.over || 0} · ขาด ${downCount.under || 0} · ไม่ได้บันทึก ${downCount.missing || 0} เที่ยว · ${downTripsTotal ? `เที่ยวขาล่อง (คิวขาล่อง) ที่กรอกการเติม: ตรง ${downTrips.ok || 0} · เกิน ${downTrips.over || 0} · ขาด ${downTrips.under || 0} · ผิดปั๊ม ${downTrips['wrong-pump'] || 0} · ตรวจไม่ได้ ${downTrips['no-rate'] || 0}` : 'เที่ยวขาล่องในคิวขาล่องยังไม่มีลิตร — ตรวจได้เมื่อทีมกรอกช่อง ปั๊มไพศาล'}</p>
                <p class="text-xs text-gray-400 mb-4">${carryNote}</p>
                <div class="flex flex-wrap items-center gap-3 mb-3">
                    <input type="text" id="audit-search" value="${esc(as.search)}" placeholder="ค้นหา คนขับ / ทะเบียน / ปลายทาง" class="w-full md:w-80 bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white" />
                    <span class="text-xs text-gray-400">แสดง ${Math.min(shown.length, LIMIT).toLocaleString()} จาก ${shown.length.toLocaleString()} เที่ยว${shown.length > LIMIT ? ' (Export Excel เพื่อดูครบ)' : ''} · คลาดเคลื่อนได้ ${tol} ลิตร · ไม่รวมรถร่วม</span>
                </div>
                <div class="overflow-x-auto rounded-lg border border-slate-700">
                    <table class="w-full text-sm text-left">
                        <thead class="bg-slate-800 text-gray-400 sticky top-0"><tr>
                            <th class="p-2">วันที่</th><th class="p-2">คนขับ</th><th class="p-2">ทะเบียน</th><th class="p-2">ประเภทรถ</th><th class="p-2">ต้นทาง → ปลายทาง</th>
                            <th class="p-2">เติมจริง (ลิตร)</th><th class="p-2">ตามอัตรา (ลิตร)</th><th class="p-2 text-right">ต่าง</th><th class="p-2">ผล</th><th class="p-2">ขาล่องไพศาล (บันทึก/อัตรา)</th><th class="p-2">ยอดยก (ลิตร)</th><th class="p-2">หมายเหตุ</th>
                        </tr></thead>
                        <tbody>${rows}</tbody>
                    </table>
                    ${shown.length === 0 ? '<p class="text-center py-6 text-gray-500">ไม่มีรายการตามเงื่อนไข</p>' : ''}
                </div>
                <h3 class="font-bold text-white mt-6 mb-2">สรุปตามคนขับ</h3>
                <div class="overflow-x-auto rounded-lg border border-slate-700">
                    <table class="w-full text-sm">
                        <thead class="bg-slate-800 text-gray-400"><tr>
                            <th class="p-2 text-left">คนขับ</th><th class="p-2 text-right">เที่ยว</th><th class="p-2 text-right">ตรวจได้</th><th class="p-2 text-right">ตรง</th>
                            <th class="p-2 text-right">เกิน (เที่ยว)</th><th class="p-2 text-right">เกิน (ลิตร)</th><th class="p-2 text-right">ขาด (เที่ยว)</th><th class="p-2 text-right">ขาด (ลิตร)</th><th class="p-2 text-right">ผิดปั๊ม</th><th class="p-2 text-right">ลิตรค้างยกไป</th>
                        </tr></thead>
                        <tbody>${driverRows}</tbody>
                    </table>
                </div>
                <div class="mt-6">
                    <button id="audit-toggle-standards" class="text-sky-300 hover:text-sky-200 text-sm font-bold">${as.showStandards ? '▼' : '▶'} ตารางอัตรามาตรฐาน (${appState.rateStandards.length} อัตรา)</button>
                    ${as.showStandards ? `
                    <div class="overflow-x-auto rounded-lg border border-slate-700 mt-2">
                        <table class="w-full text-sm text-left">
                            <thead class="bg-slate-800 text-gray-400"><tr><th class="p-2">ประเภทรถ</th><th class="p-2">เส้นทาง</th><th class="p-2">ขาขึ้น (ลิตร)</th><th class="p-2 text-right">ขาล่องไพศาล</th><th class="p-2">เบี้ยขับ</th><th class="p-2">ข้อสังเกต</th></tr></thead>
                            <tbody>${stdRows}</tbody>
                        </table>
                    </div>` : ''}
                </div>`;
        };

        // Editor for "extra fill by pickup place" rules (settings/fuelAudit.originExtras), shared by every user.
        // It edits a DRAFT (as.oxDraft, text fields kept as typed): the audit keeps using the saved rules until
        // "บันทึกกฎ" succeeds, and every keystroke is copied into the draft so a re-render never loses typing.
        // Lives outside #audit-content so filter/search re-renders do not rebuild it.
        const OX_PUMPS = Object.keys(GAS_STATION_FRIENDLY_NAMES).filter(k => k !== 'daopaisan_nakhonsawan'); // ไพศาล = downhill logic
        const oxToDraft = (rules) => (rules || []).map(r => ({ label: r.label || '', keywords: (r.keywords || []).join(', '), exclude: (r.exclude || []).join(', '), pump: r.pump || OX_PUMPS[0], maxLiters: Number(r.maxLiters) || 0 }));
        const renderOriginExtrasEditor = () => {
            const as = appState.auditState;
            const saved = as.originExtras || [];
            const draft = as.oxDraft || oxToDraft(saved);
            const dirty = !!as.oxDraft && JSON.stringify(normalizeOriginExtras(as.oxDraft)) !== JSON.stringify(saved);
            const pumpOpts = (sel) => {
                const known = OX_PUMPS.map(k => `<option value="${k}" ${k === sel ? 'selected' : ''}>${esc(GAS_STATION_FRIENDLY_NAMES[k])}</option>`).join('');
                // A saved pump that is no longer in the list stays selected and visible instead of silently changing.
                return sel && !OX_PUMPS.includes(sel) ? `<option value="${esc(sel)}" selected>(ไม่รู้จัก) ${esc(sel)}</option>${known}` : known;
            };
            const field = 'w-full bg-slate-800 border border-slate-600 rounded-md px-2 py-1 text-white';
            const row = (r, i) => `
                <div class="grid grid-cols-1 md:grid-cols-12 gap-2 items-end border-b border-slate-700 pb-3" data-ox-row="${i}">
                    <div class="md:col-span-2"><label class="block text-xs text-gray-400 mb-1">ชื่อกฎ</label><input data-ox="label" value="${esc(r.label)}" class="${field}" /></div>
                    <div class="md:col-span-3"><label class="block text-xs text-gray-400 mb-1">ต้นทางที่มีคำว่า (คั่นด้วย ,)</label><input data-ox="keywords" value="${esc(r.keywords)}" class="${field}" /></div>
                    <div class="md:col-span-2"><label class="block text-xs text-gray-400 mb-1">ยกเว้นคำว่า</label><input data-ox="exclude" value="${esc(r.exclude)}" class="${field}" /></div>
                    <div class="md:col-span-2"><label class="block text-xs text-gray-400 mb-1">ปั๊ม</label><select data-ox="pump" class="${field}">${pumpOpts(r.pump)}</select></div>
                    <div class="md:col-span-2"><label class="block text-xs text-gray-400 mb-1">ได้เพิ่มไม่เกิน (ลิตร)</label><input data-ox="maxLiters" type="number" min="0" step="1" value="${Number(r.maxLiters) || 0}" class="${field}" /></div>
                    <div class="md:col-span-1"><button data-ox-del="${i}" class="text-red-400 hover:text-red-300 text-sm">ลบ</button></div>
                </div>`;
            return `
                <button id="audit-toggle-ox" class="text-sky-300 hover:text-sky-200 text-sm font-bold">${as.showOriginExtras ? '▼' : '▶'} กฎเติมเพิ่มตามต้นทาง (${saved.length} กฎ${dirty ? ' · มีการแก้ที่ยังไม่บันทึก' : ''})</button>
                ${as.showOriginExtras ? `
                <div class="mt-2 bg-slate-800/60 border border-slate-700 rounded-lg p-3 space-y-3 text-sm">
                    <p class="text-gray-400">เที่ยวขาขึ้นที่ต้นทางมีคำที่กำหนด เติมปั๊มนี้ได้<b class="text-white">เพิ่มจากอัตรา ไม่เกิน</b> จำนวนลิตรที่ตั้ง (ไม่เติมก็ไม่นับว่าขาด) · ใช้ร่วมกันทุกคน · ผลตรวจเปลี่ยนเมื่อกด "บันทึกกฎ" แล้วเท่านั้น</p>
                    ${draft.map(row).join('') || '<p class="text-gray-500">ยังไม่มีกฎ</p>'}
                    <div class="flex flex-wrap gap-2">
                        <button id="audit-ox-add" class="bg-slate-600 hover:bg-slate-500 text-white font-bold py-1 px-3 rounded-lg">+ เพิ่มกฎ</button>
                        <button id="audit-ox-save" class="bg-primary hover:bg-secondary text-white font-bold py-1 px-3 rounded-lg">บันทึกกฎ</button>
                        ${as.oxDraft ? '<button id="audit-ox-cancel" class="bg-slate-700 hover:bg-slate-600 text-white py-1 px-3 rounded-lg">ยกเลิกการแก้</button>' : ''}
                        <span id="audit-ox-msg" class="self-center">${as.oxMsg || ''}</span>
                    </div>
                </div>` : ''}`;
        };
        const rerenderOriginExtras = () => {
            const el = document.getElementById('audit-ox');
            if (el) { el.innerHTML = renderOriginExtrasEditor(); attachOriginExtrasListeners(); }
        };
        const readOriginExtrasForm = () => [...document.querySelectorAll('[data-ox-row]')].map(el => {
            const v = (f) => el.querySelector(`[data-ox="${f}"]`)?.value ?? '';
            return { label: v('label'), keywords: v('keywords'), exclude: v('exclude'), pump: v('pump'), maxLiters: Number(v('maxLiters')) || 0 };
        });
        const attachOriginExtrasListeners = () => {
            const as = appState.auditState;
            document.getElementById('audit-toggle-ox')?.addEventListener('click', () => { as.showOriginExtras = !as.showOriginExtras; rerenderOriginExtras(); });
            // Copy every edit into the draft at once, so any re-render shows what was typed.
            document.querySelectorAll('[data-ox-row] [data-ox]').forEach(inp => {
                const sync = () => { as.oxDraft = readOriginExtrasForm(); as.oxMsg = ''; };
                inp.addEventListener('input', sync);
                inp.addEventListener('change', sync);
            });
            document.getElementById('audit-ox-add')?.addEventListener('click', () => {
                as.oxDraft = [...readOriginExtrasForm(), { label: '', keywords: '', exclude: '', pump: 'tekyong_chonburi', maxLiters: 0 }];
                rerenderOriginExtras();
            });
            document.querySelectorAll('[data-ox-del]').forEach(b => b.addEventListener('click', () => {
                const rows = readOriginExtrasForm();
                rows.splice(Number(b.dataset.oxDel), 1);
                as.oxDraft = rows;
                rerenderOriginExtras();
            }));
            document.getElementById('audit-ox-cancel')?.addEventListener('click', () => { as.oxDraft = null; as.oxMsg = ''; rerenderOriginExtras(); });
            document.getElementById('audit-ox-save')?.addEventListener('click', async () => {
                const raw = readOriginExtrasForm();
                const rules = normalizeOriginExtras(raw);
                if (rules.length !== raw.length && !confirm(`มี ${raw.length - rules.length} กฎที่ไม่ครบ (ต้องมีคำต้นทาง ปั๊ม และลิตรมากกว่า 0) จะถูกตัดออก — บันทึกต่อไหม?`)) return;
                try {
                    // Firebase drops an empty array, which would bring the default back — store false for "no rules".
                    await logisticsService.saveAuditSettings({ originExtras: rules.length ? rules : false });
                    as.originExtras = rules;
                    as.oxDraft = null;
                    as.oxMsg = `<span class="text-green-400">✓ บันทึกกฎแล้ว (${rules.length} กฎ)</span>`;
                    rerenderOriginExtras();
                    rerenderAuditContent();
                } catch (err) {
                    console.error('Save origin extras failed:', err);
                    as.oxMsg = `<span class="text-red-400">บันทึกไม่สำเร็จ: ${esc(err.message)}</span>`;
                    rerenderOriginExtras();
                }
            });
        };

        const rerenderAuditContent = () => {
            const el = document.getElementById('audit-content');
            if (el) { el.innerHTML = renderAuditContent(); attachAuditContentListeners(); }
        };
        const attachAuditContentListeners = () => {
            const as = appState.auditState;
            document.querySelectorAll('[data-audit-status]').forEach(b => b.addEventListener('click', () => { as.status = b.dataset.auditStatus; rerenderAuditContent(); }));
            const search = document.getElementById('audit-search');
            search?.addEventListener('input', (e) => {
                as.search = e.target.value;
                const pos = e.target.selectionStart;
                rerenderAuditContent();
                const again = document.getElementById('audit-search');
                if (again) { again.focus(); again.setSelectionRange(pos, pos); }
            });
            document.getElementById('audit-toggle-standards')?.addEventListener('click', () => { as.showStandards = !as.showStandards; rerenderAuditContent(); });
            document.getElementById('audit-carry-refresh')?.addEventListener('click', () => { refreshAuditPast(); rerenderAuditContent(); });
        };
        const attachAuditListeners = () => {
            const monthInput = document.getElementById('audit-month');
            const target = appState.tripsMonth || (monthInput ? monthInput.value : '');
            if (target && appState.tripsMonth !== target) logisticsService.loadTripsForMonth(target);
            monthInput?.addEventListener('change', (e) => { refreshAuditPast(); logisticsService.loadTripsForMonth(e.target.value); });
            // Shared settings: saved to Firebase so every user audits with the same tolerance / carry start.
            document.getElementById('audit-tolerance')?.addEventListener('change', (e) => {
                appState.auditState.tolerance = Math.max(0, Number(e.target.value) || 0);
                rerenderAuditContent();
                logisticsService.saveAuditSettings({ tolerance: appState.auditState.tolerance }).catch(err => console.error('Save audit settings failed:', err));
            });
            document.getElementById('audit-carry-start')?.addEventListener('change', (e) => {
                appState.auditState.carryStart = /^\d{4}-\d{2}$/.test(e.target.value) ? e.target.value : '';
                refreshAuditPast();
                rerenderAuditContent();
                logisticsService.saveAuditSettings({ carryStart: appState.auditState.carryStart }).catch(err => console.error('Save audit settings failed:', err));
            });
            document.getElementById('audit-export-btn')?.addEventListener('click', handleAuditExport);
            attachAuditContentListeners();
            attachOriginExtrasListeners();
        };

        const handleAuditExport = () => {
            if (!appState.rateStandards.length) { alert('ยังไม่มีอัตรามาตรฐาน'); return; }
            const { all, drivers, tol, carryOn, carryReady } = computeAudit();
            if (!all.length) { alert('ไม่มีเที่ยวให้ export ในเดือนนี้'); return; }
            // Never hand out a file with blank carry columns that look final.
            if (carryOn && !carryReady) { alert('ยอดยกลิตรยังคำนวณไม่เสร็จ (กำลังโหลดหรือโหลดเดือนก่อนหน้าไม่สำเร็จ) — รอสักครู่หรือกด "โหลดยอดยกใหม่" แล้วลอง Export อีกครั้ง'); return; }
            const pumps = [...new Set(all.flatMap(r => [...Object.keys(r.actual), ...Object.keys(r.expected || {})]))];
            const wb = XLSX.utils.book_new();
            const sheet = (name, headers, data, widths) => {
                const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
                ws['!cols'] = widths.map(w => ({ wch: w }));
                ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: data.length, c: headers.length - 1 } }) };
                XLSX.utils.book_append_sheet(wb, ws, name);
            };
            const n = (v) => (Number(v) ? Number(v) : '');
            sheet('ผลตรวจ',
                ['วันที่', 'ทิศทาง', 'คนขับ', 'ชื่อในกระดานคิว', 'ทะเบียน', 'ประเภทรถ', 'รุ่น', 'ต้นทาง', 'ปลายทาง', 'ผล', 'เติมจริงรวม (ลิตร)', 'ตามอัตรารวม (ลิตร)', 'ต่าง (ลิตร)',
                    ...pumps.flatMap(p => [`${pumpName(p)} จริง`, `${pumpName(p)} อัตรา`]), 'เติมเป็นเงิน (บาท)',
                    'ขาล่องไพศาล อัตรา', 'ขาล่องไพศาล บันทึก', 'ขาล่อง ผล', 'ลิตรยกมา', 'ควรเติม (หลังหักยกมา)', 'ลิตรยกไป', 'หมายเหตุ'],
                all.map(r => [r.trip.date, r.leg === 'down' ? 'ขาล่อง' : 'ขาขึ้น', r.trip.driver, r.trip.driverBoard || '', r.trip.plate, r.cls ? AUDIT_CLASSES[r.cls] + (r.rateCls && r.rateCls !== r.cls ? ` (ตัดหาง คิด${AUDIT_CLASSES[r.rateCls]})` : '') : '', r.model, r.trip.origin, r.trip.destination,
                    AUDIT_STATUS[r.status].label, n(r.actualTotal), r.status === 'no-rate' ? '' : n(r.expectedTotal), r.status === 'no-rate' ? '' : r.diff,
                    ...pumps.flatMap(p => [n(r.actual[p]), r.status === 'no-rate' ? '' : n((r.expected || {})[p])]), n(r.bahtTotal),
                    n(r.down), r.downStatus && r.downStatus !== 'missing' ? r.downActual : '', r.downStatus ? DOWN_STATUS[r.downStatus].label : '',
                    carryReady && r.carryIn != null ? r.carryIn : '', carryReady && r.allot != null ? r.allot : '', carryReady && r.carryOut != null ? r.carryOut : '',
                    [r.reason, ...r.flags].filter(Boolean).join(' | ')]),
                [11, 8, 22, 14, 18, 8, 20, 18, 12, 12, 12, 10, ...pumps.flatMap(() => [12, 12]), 12, 12, 12, 12, 10, 14, 10, 60]);
            sheet('สรุปคนขับ',
                ['คนขับ', 'เที่ยว', 'ตรวจได้', 'ตรง', 'เกิน (เที่ยว)', 'เกิน (ลิตร)', 'ขาด (เที่ยว)', 'ขาด (ลิตร)', 'ผิดปั๊ม', 'ลิตรค้างยกไป (+ เกิน / − ขาด)'],
                drivers.map(d => [d.driver, d.n, d.checked, d.ok, d.over, d.overL, d.under, d.underL, d.wrong, d.balance == null ? '' : d.balance]),
                [24, 8, 8, 8, 10, 10, 10, 10, 8, 16]);
            sheet('อัตรามาตรฐาน',
                ['ชีต', 'แถว', 'ประเภทรถ', 'รุ่น', 'ต้นทาง', 'ปลายทาง', 'ขาขึ้น (ลิตร)', 'รวมขาขึ้น', 'ขาล่องไพศาล', 'ขาล่อง victor', 'เบี้ยขับ', 'ข้อสังเกต'],
                appState.rateStandards.map(s => [s.sheet, s.row, AUDIT_CLASSES[s.cls], s.variant === 'sino' ? 'sino' : '', s.origin, s.destinationLabel,
                    Object.entries(s.up).map(([k, v]) => `${pumpName(k)} ${v}`).join(' + '), s.upTotal, n(s.down), n(s.downVictor), s.allowanceText, (s.flags || []).join(' | ')]),
                [14, 6, 18, 8, 12, 24, 40, 10, 12, 12, 30, 50]);
            XLSX.writeFile(wb, `ตรวจการเติมน้ำมัน_${appState.tripsMonth || localDateStr().slice(0, 7)}_คลาดเคลื่อน${tol}ลิตร.xlsx`);
        };

        const attachReportsListeners = () => {
            const monthInput = document.getElementById('reports-month');
            const target = appState.tripsMonth || (monthInput ? monthInput.value : '');
            if (target && appState.tripsMonth !== target) logisticsService.loadTripsForMonth(target);
            monthInput?.addEventListener('change', (e) => logisticsService.loadTripsForMonth(e.target.value));
            attachReportsContentListeners();
        };

        const renderVehicleManagementContent = () => {
            const filter = appState.vehicleManagementFilter.toLowerCase();
            let vehicles = logisticsService.getVehicles();
            if (filter) vehicles = vehicles.filter(v =>
                (v.registrationNumber || '').toLowerCase().includes(filter) ||
                (v.vehicleType || '').toLowerCase().includes(filter) ||
                (v.company || '').toLowerCase().includes(filter) ||
                (v.driveType || '').toLowerCase().includes(filter)
            );
            const partnerCount = vehicles.filter(v => (v.driveType || '').includes('ร่วม')).length;

            const tableRows = vehicles.map(vehicle => {
                const isPartner = (vehicle.driveType || '').includes('ร่วม');
                return `
                <tr key="${vehicle.id}" class="border-b border-slate-600 hover:bg-slate-700/20">
                    <td class="p-3 font-mono">${vehicle.registrationNumber}</td>
                    <td class="p-3">${vehicle.vehicleType || ''}${vehicle.model ? ` <span class="ml-1 px-2 py-0.5 rounded text-xs bg-indigo-700 text-white">${vehicle.model}</span>` : ''}</td>
                    <td class="p-3 text-center">${vehicle.driveType ? `<span class="px-2 py-0.5 rounded text-xs ${isPartner ? 'bg-amber-700' : 'bg-sky-700'} text-white">${vehicle.driveType}</span>` : ''}</td>
                    <td class="p-3">${vehicle.company || ''}</td>
                    <td class="p-3 text-center">
                        <div class="flex justify-center gap-2">
                            <button data-action="edit-vehicle" data-id="${vehicle.id}" class="text-blue-400 hover:text-blue-300 p-2 rounded-full transition-colors">${ICONS.Pencil}</button>
                            <button data-action="delete-vehicle" data-id="${vehicle.id}" class="text-red-500 hover:text-red-400 p-2 rounded-full transition-colors">${ICONS.Trash}</button>
                        </div>
                    </td>
                </tr>`;
            }).join('');

            return renderCard(`
                <p class="text-sm text-gray-400 mb-2">ทั้งหมด ${vehicles.length} คัน · รถร่วม ${partnerCount} คัน</p>
                <div class="overflow-x-auto">
                    <table class="w-full text-left">
                        <thead class="bg-slate-800 text-sm text-gray-400 uppercase">
                            <tr>
                                <th class="p-3">ทะเบียนรถ</th><th class="p-3">ประเภทรถ</th><th class="p-3 text-center">ชนิด</th><th class="p-3">บริษัท</th><th class="p-3 text-center">จัดการ</th>
                            </tr>
                        </thead>
                        <tbody id="vehicles-table-body">${tableRows}</tbody>
                    </table>
                    ${vehicles.length === 0 ? '<p class="text-center py-8 text-gray-500">ไม่พบข้อมูลทะเบียนรถ</p>' : ''}
                </div>
            `);
        };

        // --- View: Vehicle Management ---
        const renderVehicleManagement = () => {
            return `
                <div class="space-y-8">
                    ${renderCard(`
                        <div class="flex flex-wrap justify-between items-center gap-4">
                            <h2 class="text-2xl font-bold text-white flex items-center gap-3">${ICONS.IdCard}จัดการข้อมูลทะเบียนรถ</h2>
                            <div class="flex items-center gap-4">
                                <input type="text" id="vehicle-management-filter" value="${appState.vehicleManagementFilter}" placeholder="ค้นหาทะเบียน/ประเภท..." class="w-full md:w-64 bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" />
                                 <button id="refresh-vehicles-btn" title="รีเฟรชข้อมูล" class="bg-sky-600 hover:bg-sky-700 text-white font-bold p-2 rounded-lg shadow-md transition-transform transform hover:scale-105">
                                    ${ICONS.Refresh}
                                </button>
                                <button id="add-new-vehicle-btn" class="flex items-center gap-2 bg-primary hover:bg-secondary text-white font-bold py-2 px-4 rounded-lg shadow-md transition-transform transform hover:scale-105">
                                    ${ICONS.Plus} เพิ่มทะเบียนใหม่
                                </button>
                            </div>
                        </div>
                    `)}
                    <div id="vehicle-management-content">
                        ${renderVehicleManagementContent()}
                    </div>
                </div>
            `;
        };

        // --- Modals ---
        const renderModal = () => `<div id="modal-container" class="hidden"></div>`;
        const renderAuthModal = () => `<div id="auth-modal-container" class="hidden"></div>`;

        const openModal = (title, content) => {
            const modalContainer = document.getElementById('modal-container');
            modalContainer.innerHTML = `
                <div class="fixed inset-0 bg-black bg-opacity-70 z-40 flex justify-center items-center">
                    <div class="bg-neutral rounded-xl shadow-2xl w-full max-w-2xl max-h-[90vh] flex flex-col">
                        <header class="flex justify-between items-center p-4 border-b border-slate-600">
                            <h2 class="text-xl font-bold text-white">${title}</h2>
                            <button data-action="close-modal" class="text-gray-400 hover:text-white transition-colors">
                                <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12" /></svg>
                            </button>
                        </header>
                        <main class="p-6 overflow-y-auto">${content}</main>
                    </div>
                </div>
            `;
            modalContainer.classList.remove('hidden');
        };

        const closeModal = () => {
             document.getElementById('modal-container').classList.add('hidden');
             document.getElementById('auth-modal-container').classList.add('hidden');
        };
        
        const openAuthModal = () => {
             const modalContainer = document.getElementById('auth-modal-container');
             modalContainer.innerHTML = `
                <div class="fixed inset-0 bg-black bg-opacity-70 z-40 flex justify-center items-center">
                    <div class="bg-neutral rounded-xl shadow-2xl w-full max-w-md">
                        <header class="flex justify-between items-center p-4 border-b border-slate-600">
                            <h2 class="text-xl font-bold text-white">เข้าสู่หน้าจัดการข้อมูล</h2>
                            <button data-action="close-modal" class="text-gray-400 hover:text-white transition-colors">
                                <svg xmlns="http://www.w3.org/2000/svg" class="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12" /></svg>
                            </button>
                        </header>
                        <main class="p-6">
                             <form id="auth-form" class="space-y-4">
                                <div>
                                    <label for="password" class="block text-sm font-medium text-gray-400 mb-1">กรุณากรอกรหัสผ่านเพื่อดำเนินการต่อ</label>
                                    <input id="password" type="password" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" autofocus />
                                </div>
                                <p id="auth-error" class="text-sm text-red-400 hidden"></p>
                                <div class="flex justify-end pt-2">
                                    <button type="submit" class="bg-primary hover:bg-secondary text-white font-bold py-2 px-6 rounded-lg shadow-md transition-transform transform hover:scale-105">ยืนยัน</button>
                                </div>
                            </form>
                        </main>
                    </div>
                </div>
             `;
             modalContainer.classList.remove('hidden');
             const authForm = document.getElementById('auth-form');
             if (authForm) {
                 authForm.addEventListener('submit', handleAuthSubmit);
             }
        }

        const renderRouteFormContent = (route) => {
            const distributionItemsLiters = route?.gasStationsLiters ? Object.entries(route.gasStationsLiters).map(([stationKey, amount]) => ({stationKey, amount})) : [];
            const distributionItemsBaht = route?.gasStationsBaht ? Object.entries(route.gasStationsBaht).map(([stationKey, amount]) => ({stationKey, amount})) : [];
            
             return `
                <form id="route-form" class="space-y-6">
                    <input type="hidden" name="id" value="${route?.id || ''}" />
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div>
                            <label class="block text-sm font-medium text-gray-400 mb-1">ประเภทรถ</label>
                            <input type="text" name="vehicleType" value="${route?.vehicleType || ''}" placeholder="เช่น, 18 ล้อ" required class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" />
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-gray-400 mb-1">ต้นทาง</label>
                            <input type="text" name="origin" value="${route?.origin || ''}" placeholder="เช่น, กทม" required class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" />
                        </div>
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-gray-400 mb-1">ปลายทาง</label>
                        <input type="text" name="destination" value="${route?.destination || ''}" placeholder="เช่น, เชียงใหม่" required class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" />
                    </div>
                    <div class="grid grid-cols-1 sm:grid-cols-3 gap-4">
                        <div>
                            <label class="block text-sm font-medium text-gray-400 mb-1">ทิศทาง</label>
                            <select name="direction" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent">
                                <option value="ขาขึ้น" ${(route?.direction || 'ขาขึ้น') === 'ขาขึ้น' ? 'selected' : ''}>ขาขึ้น</option>
                                <option value="ขาล่อง" ${route?.direction === 'ขาล่อง' ? 'selected' : ''}>ขาล่อง</option>
                            </select>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-gray-400 mb-1">ค่าข้าม</label>
                            <input type="text" name="crossingFee" value="${route?.crossingFee || ''}" placeholder="เช่น, ไม่ข้าม / ข้าม1" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" />
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-gray-400 mb-1">ค่าค้าง</label>
                            <input type="text" name="waitingFee" value="${route?.waitingFee || ''}" placeholder="เช่น, ไม่ค้าง / ค้าง200" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" />
                        </div>
                    </div>

                    <div class="space-y-4 rounded-lg border border-slate-600 p-4">
                        <h3 class="font-semibold text-lg text-white">การกระจายน้ำมัน (ลิตร)</h3>
                        <div id="route-distribution-list-liters">
                             ${distributionItemsLiters.map((item, index) => renderRouteDistributionRow(item, index, 'Liters')).join('')}
                        </div>
                        <button type="button" data-action="add-route-distribution" data-unit="Liters" class="w-full border-2 border-dashed border-slate-500 hover:bg-slate-700/50 text-gray-300 font-semibold py-2 px-4 rounded-lg transition">+ เพิ่มจุดเติมน้ำมัน (ลิตร)</button>
                    </div>
                    
                    <div class="space-y-4 rounded-lg border border-slate-600 p-4">
                        <h3 class="font-semibold text-lg text-white">การกระจายน้ำมัน (บาท)</h3>
                        <div id="route-distribution-list-baht">
                             ${distributionItemsBaht.map((item, index) => renderRouteDistributionRow(item, index, 'Baht')).join('')}
                        </div>
                        <button type="button" data-action="add-route-distribution" data-unit="Baht" class="w-full border-2 border-dashed border-slate-500 hover:bg-slate-700/50 text-gray-300 font-semibold py-2 px-4 rounded-lg transition">+ เพิ่มจุดเติมน้ำมัน (บาท)</button>
                    </div>


                    <footer class="flex justify-end gap-4 pt-4">
                        <button type="button" data-action="close-modal" class="bg-slate-600 hover:bg-slate-500 text-white font-bold py-2 px-4 rounded-lg transition">ยกเลิก</button>
                        <button type="submit" class="bg-primary hover:bg-secondary text-white font-bold py-2 px-6 rounded-lg transition">บันทึกข้อมูล</button>
                    </footer>
                </form>
             `;
        };
        
        const renderVehicleFormContent = (vehicle) => {
            return `
                <form id="vehicle-form" class="space-y-6">
                    <input type="hidden" name="id" value="${vehicle?.id || ''}" />
                    <div>
                        <label class="block text-sm font-medium text-gray-400 mb-1">ทะเบียนรถ</label>
                        <input type="text" name="registrationNumber" value="${vehicle?.registrationNumber || ''}" placeholder="เช่น, 70-1234" required class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" />
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-gray-400 mb-1">ประเภทรถ</label>
                        <input type="text" name="vehicleType" value="${vehicle?.vehicleType || ''}" placeholder="เช่น, 18 ล้อ HINO" required class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" />
                    </div>
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div>
                            <label class="block text-sm font-medium text-gray-400 mb-1">ชนิด</label>
                            <select name="driveType" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent">
                                <option value="รถบริษัท" ${(vehicle?.driveType || 'รถบริษัท') === 'รถบริษัท' ? 'selected' : ''}>รถบริษัท</option>
                                <option value="รถร่วม" ${vehicle?.driveType === 'รถร่วม' ? 'selected' : ''}>รถร่วม</option>
                            </select>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-gray-400 mb-1">บริษัทรถ</label>
                            <input type="text" name="company" value="${vehicle?.company || ''}" placeholder="เช่น, นีโอสยาม" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" />
                        </div>
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-gray-400 mb-1">รุ่นรถ (ใช้เลือกอัตราเติมน้ำมัน)</label>
                        <select name="model" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent">
                            <option value="" ${!vehicle?.model ? 'selected' : ''}>ทั่วไป</option>
                            <option value="sino" ${vehicle?.model === 'sino' ? 'selected' : ''}>Sino (อัตรา "รถ sino")</option>
                            <option value="victor" ${vehicle?.model === 'victor' ? 'selected' : ''}>Victor (ขาล่องไพศาล 190 ลิตร)</option>
                        </select>
                    </div>
                    <footer class="flex justify-end gap-4 pt-4">
                        <button type="button" data-action="close-modal" class="bg-slate-600 hover:bg-slate-500 text-white font-bold py-2 px-4 rounded-lg transition">ยกเลิก</button>
                        <button type="submit" class="bg-primary hover:bg-secondary text-white font-bold py-2 px-6 rounded-lg transition">บันทึกข้อมูล</button>
                    </footer>
                </form>
            `;
        };

        const renderFuelRecordEditFormContent = (record) => {
            const distributionLiters = record.fuelDistributionLiters ? Object.entries(record.fuelDistributionLiters).map(([stationKey, amount]) => ({ stationKey, amount })) : [];
            const distributionBaht = record.fuelDistributionBaht ? Object.entries(record.fuelDistributionBaht).map(([stationKey, amount]) => ({ stationKey, amount })) : [];

             return `
                <form id="fuel-record-edit-form" class="space-y-4">
                    <input type="hidden" name="id" value="${record.id}" />
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div>
                            <label class="block text-sm font-medium text-gray-400 mb-1">วันที่</label>
                            <input type="date" name="date" value="${record.date}" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white" />
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-gray-400 mb-1">ทะเบียนรถ</label>
                            <input type="text" name="vehicleRegistration" value="${record.vehicleRegistration}" placeholder="เช่น, กข-1234" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white" />
                        </div>
                    </div>
                     <div>
                        <label class="block text-sm font-medium text-gray-400 mb-1">ประเภทรถ</label>
                        <input type="text" name="vehicleType" value="${record.vehicleType}" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white" />
                    </div>
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                         <div>
                            <label class="block text-sm font-medium text-gray-400 mb-1">ต้นทาง</label>
                            <input type="text" name="origin" value="${record.origin}" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white" />
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-gray-400 mb-1">ปลายทาง</label>
                            <input type="text" name="destination" value="${record.destination}" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white" />
                        </div>
                    </div>
                    
                    <!-- Liters Distribution -->
                    <div class="space-y-4 rounded-lg border border-slate-600 p-4">
                        <h3 class="font-semibold text-lg text-white">การกระจายน้ำมัน (ลิตร)</h3>
                        <div id="edit-distribution-list-liters">
                            ${distributionLiters.map((item, index) => renderEditDistributionRow(item, index, 'Liters')).join('')}
                        </div>
                        <button type="button" data-action="add-edit-distribution" data-unit="Liters" class="w-full border-2 border-dashed border-slate-500 hover:bg-slate-700/50 text-gray-300 font-semibold py-2 px-4 rounded-lg transition">+ เพิ่มจุดเติมน้ำมัน (ลิตร)</button>
                    </div>

                     <!-- Baht Distribution -->
                    <div class="space-y-4 rounded-lg border border-slate-600 p-4">
                        <h3 class="font-semibold text-lg text-white">การกระจายน้ำมัน (บาท)</h3>
                        <div id="edit-distribution-list-baht">
                            ${distributionBaht.map((item, index) => renderEditDistributionRow(item, index, 'Baht')).join('')}
                        </div>
                        <button type="button" data-action="add-edit-distribution" data-unit="Baht" class="w-full border-2 border-dashed border-slate-500 hover:bg-slate-700/50 text-gray-300 font-semibold py-2 px-4 rounded-lg transition">+ เพิ่มจุดเติมน้ำมัน (บาท)</button>
                    </div>
                    
                    <div>
                        <label class="block text-sm font-medium text-gray-400 mb-1">หมายเหตุ</label>
                        <textarea name="notes" placeholder="หมายเหตุเพิ่มเติมเกี่ยวกับการเดินทาง" rows="3" class="w-full bg-slate-800 border border-slate-600 rounded-md shadow-sm px-3 py-2 text-white">${record.notes || ''}</textarea>
                    </div>

                    <footer class="flex justify-end gap-4 pt-4">
                        <button type="button" data-action="close-modal" class="bg-slate-600 hover:bg-slate-500 text-white font-bold py-2 px-4 rounded-lg transition">ยกเลิก</button>
                        <button type="submit" class="bg-primary hover:bg-secondary text-white font-bold py-2 px-6 rounded-lg transition">บันทึกการเปลี่ยนแปลง</button>
                    </footer>
                </form>
             `;
        };
        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        // 5. EVENT LISTENERS & HANDLERS
        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
        let isMainListenerAttached = false;

        const attachEventListeners = () => {
            const container = document.getElementById('app-container');
            if (!container) return;

            // Use event delegation for dynamic elements, but only attach it once.
            if (!isMainListenerAttached) {
                container.addEventListener('click', handleAppClick);
                isMainListenerAttached = true;
            }

            // Attach listeners for specific elements within the current view
            if (appState.view === 'form') {
                attachFormListeners();
            } else if (appState.view === 'dashboard') {
                attachDashboardListeners();
            } else if (appState.view === 'data') {
                attachDataManagementListeners();
            } else if (appState.view === 'vehicles') {
                attachVehicleManagementListeners();
            } else if (appState.view === 'import') {
                attachImportListeners();
            } else if (appState.view === 'queue') {
                attachQueueListeners();
            } else if (appState.view === 'drivers') {
                attachDriversListeners();
            } else if (appState.view === 'payroll') {
                attachPayrollListeners();
            } else if (appState.view === 'reports') {
                attachReportsListeners();
            } else if (appState.view === 'audit') {
                attachAuditListeners();
            }
        };

        const handleAppClick = (e) => {
            const target = e.target.closest('[data-action], .nav-button');
            if (!target) return;

            const action = target.dataset.action;
            const id = target.dataset.id;
            
            // --- Global Actions ---
            if (target.classList.contains('nav-button')) {
                const view = target.dataset.view;
                if (['data', 'vehicles', 'import', 'queue', 'drivers', 'payroll', 'reports', 'audit'].includes(view) && !appState.isAuthenticated) {
                    openAuthModal();
                    // Store the intended view to navigate to after successful authentication
                    appState.pendingView = view; 
                } else {
                    if (view === 'audit') refreshAuditPast(); // pick up edits made in earlier months since last time
                    appState.view = view;
                    renderApp();
                }
            }
            if (action === 'close-modal') closeModal();

            // --- Authentication Actions ---
             if (action === 'auth-confirm') {
                 // Logic handled by form submission
             }

            // --- Route Actions ---
            if (action === 'add-new-route') handleAddNewRoute();
            if (action === 'edit-route') handleEditRoute(id);
            if (action === 'delete-route') handleDeleteRoute(id);
            if (action === 'add-route-distribution') handleAddRouteDistribution(target);

            // --- Trip / Queue Actions ---
            if (action === 'edit-trip') handleEditTrip(id);
            if (action === 'delete-trip') handleDeleteTrip(id);
            if (action === 'calc-allowance') handleCalcAllowance(target);
            if (action === 'calc-allowance-form') handleCalcAllowanceForm();
            if (action === 'apply-route-to-trip') handleApplyRouteToTrip(target);

            // --- Driver Actions ---
            if (action === 'edit-driver') handleEditDriver(id);
            if (action === 'delete-driver') handleDeleteDriver(id);

            // --- Fuel Record Actions ---
            if (action === 'apply-route') handleApplyRouteToForm();
            if (action === 'edit-fuel-record') handleEditFuelRecord(id);
            if (action === 'delete-fuel-record') handleDeleteFuelRecord(id);
            if (action === 'add-edit-distribution') handleAddEditDistribution(target);

             // --- Vehicle Actions ---
            if (action === 'add-new-vehicle') handleAddNewVehicle();
            if (action === 'edit-vehicle') handleEditVehicle(id);
            if (action === 'delete-vehicle') handleDeleteVehicle(id);

            // --- Common Form Actions ---
            if (action === 'remove-distribution') {
                const row = e.target.closest('.distribution-row');
                if (row) row.remove();

                // If in main form, update totals
                if (document.getElementById('fuel-record-form')) {
                    updateTotalAmount('liters');
                    updateTotalAmount('baht');
                }
            }
        };
        
        const attachFormListeners = () => {
            const form = document.getElementById('fuel-record-form');
            if (form) form.addEventListener('submit', handleTripEntrySubmit);

            // Autocomplete listeners
            attachAutocompleteOutsideClick();
            const driverNames = [...new Set(logisticsService.getDrivers().map(d => d.name).filter(Boolean))];
            const runTypes = (() => { const rt = [...new Set(logisticsService.getRateTable().map(r => r.runType).filter(Boolean))]; return rt.length ? rt : ['หนักล่อง', 'ตีเปล่า', 'คิวพิเศษ']; })();
            attachAutocompleteListeners('form-vehicle-registration', logisticsService.getVehicleRegistrations());
            attachAutocompleteListeners('form-vehicle-type', logisticsService.getUniqueValues('vehicleType'));
            attachAutocompleteListeners('form-driver', driverNames);
            attachAutocompleteListeners('form-queue-type', runTypes);
            attachAutocompleteListeners('form-origin', logisticsService.getUniqueValues('origin'));
            attachAutocompleteListeners('form-destination', logisticsService.getUniqueValues('destination'));

            // V-lookup trigger
            ['form-vehicle-type', 'form-origin', 'form-destination'].forEach(id => {
                const input = document.getElementById(id);
                if(input) input.addEventListener('change', updateVlookupCard);
            });
            // Auto-calc เบี้ยเลี้ยง/หัก/สุทธิโอน when ประเภทคิว or ประเภทรถ changes (silent, no popup).
            ['form-queue-type', 'form-vehicle-type'].forEach(id => {
                document.getElementById(id)?.addEventListener('change', () => handleCalcAllowanceForm(true));
            });

            // Auto-fill ประเภทรถ from the selected ทะเบียนรถ (each vehicle has a known type).
            const regInput = document.getElementById('form-vehicle-registration');
            if (regInput) {
                regInput.addEventListener('change', () => {
                    const reg = regInput.value.trim();
                    if (!reg) return;
                    const vehicle = logisticsService.getVehicles().find(v => (v.registrationNumber || '').trim() === reg);
                    const typeInput = document.getElementById('form-vehicle-type');
                    if (vehicle && vehicle.vehicleType && typeInput) {
                        typeInput.value = vehicle.vehicleType;
                        updateVlookupCard();
                    }
                });
            }
            
            // Distribution buttons
            const addLitersBtn = document.getElementById('add-distribution-btn-liters');
            const addBahtBtn = document.getElementById('add-distribution-btn-baht');
            if(addLitersBtn) addLitersBtn.addEventListener('click', () => updateDistributionList('liters'));
            if(addBahtBtn) addBahtBtn.addEventListener('click', () => updateDistributionList('baht'));
            
            // Reset button
            const resetBtn = document.getElementById('reset-fuel-form-btn');
            if(resetBtn) resetBtn.addEventListener('click', resetFuelForm);

            // Recalculate totals on change
            document.getElementById('distribution-list-liters')?.addEventListener('input', () => updateTotalAmount('liters'));
            document.getElementById('distribution-list-baht')?.addEventListener('input', () => updateTotalAmount('baht'));

            // Initial V-lookup call
            updateVlookupCard();
        };

        const attachDashboardListeners = () => {
            const monthInput = document.getElementById('dashboard-month');
            const target = appState.tripsMonth || (monthInput ? monthInput.value : '');
            if (target && appState.tripsMonth !== target) logisticsService.loadTripsForMonth(target);
            monthInput?.addEventListener('change', (e) => {
                const ym = e.target.value;
                // If a specific-day filter is set but falls outside the newly chosen month, clear it
                // so we don't filter this month's trips by a date it can never contain (→ 0 results).
                if (appState.dashboardFilters.day && appState.dashboardFilters.day.slice(0, 7) !== ym) {
                    appState.dashboardFilters.day = '';
                    const di = document.getElementById('dashboard-day'); if (di) di.value = '';
                }
                logisticsService.loadTripsForMonth(ym);
            });

            const rerender = () => { const c = document.getElementById('dashboard-content'); if (c) c.innerHTML = renderDashboardContent(); };
            document.getElementById('dashboard-day')?.addEventListener('change', (e) => {
                const val = e.target.value;
                appState.dashboardFilters.day = val;
                // A day picked outside the currently loaded month must trigger loading that month's
                // trips first — otherwise the exact-date filter matches nothing. loadTripsForMonth
                // re-renders the dashboard once data arrives, and the day filter still applies.
                if (val) {
                    const ym = val.slice(0, 7);
                    if (ym !== appState.tripsMonth) {
                        const mi = document.getElementById('dashboard-month'); if (mi) mi.value = ym;
                        logisticsService.loadTripsForMonth(ym);
                        return;
                    }
                }
                rerender();
            });
            document.getElementById('dashboard-plate')?.addEventListener('input', (e) => { appState.dashboardFilters.vehicleRegistration = e.target.value; rerender(); });
            document.getElementById('dashboard-clear-day')?.addEventListener('click', () => {
                appState.dashboardFilters.day = '';
                const di = document.getElementById('dashboard-day'); if (di) di.value = '';
                rerender();
            });
            document.getElementById('export-excel-btn')?.addEventListener('click', handleExportToExcel);
        };

        const attachDataManagementListeners = () => {
            const filterInput = document.getElementById('data-management-filter');
            if (filterInput) {
                filterInput.addEventListener('input', (e) => {
                    appState.dataManagementFilter = e.target.value;
                    const content = document.getElementById('data-management-content');
                    if(content) content.innerHTML = renderDataManagementContent();
                });
            }
             const refreshBtn = document.getElementById('refresh-routes-btn');
            if(refreshBtn) {
                refreshBtn.addEventListener('click', () => {
                    const content = document.getElementById('data-management-content');
                    if(content) content.innerHTML = renderDataManagementContent();
                });
            }
            const backupBtn = document.getElementById('backup-data-btn');
            if(backupBtn) {
                backupBtn.addEventListener('click', handleBackupData);
            }
            document.getElementById('routes-export-btn')?.addEventListener('click', handleRoutesExport);
        };

        // Export the route table exactly as shown (same search filter, sort and pump columns).
        const handleRoutesExport = () => {
            const routes = getVisibleRoutes();
            if (!routes.length) { alert('ไม่มีเส้นทางให้ export'); return; }
            const label = (key, unit) => `${GAS_STATION_FRIENDLY_NAMES[key] || key} (${unit === 'Liters' ? 'ลิตร' : 'บาท'})`;
            const cols = ROUTE_STATION_COLUMNS.flatMap(col => col.units.map(unit => ({ key: col.key, unit, label: label(col.key, unit) })));
            // A route can hold a pump/unit pair the on-screen table has no column for (the edit form allows any
            // pump in either unit) — add those as extra columns so no amount is dropped from the file.
            const seen = new Set(cols.map(c => `${c.unit}|${c.key}`));
            for (const r of routes) {
                for (const [unit, obj] of [['Liters', r.gasStationsLiters], ['Baht', r.gasStationsBaht]]) {
                    for (const [key, v] of Object.entries(obj || {})) {
                        if (v !== '' && v != null && !seen.has(`${unit}|${key}`)) { seen.add(`${unit}|${key}`); cols.push({ key, unit, label: label(key, unit) }); }
                    }
                }
            }
            const num = (v) => { const n = Number(v); return v === '' || v == null || !isFinite(n) || n === 0 ? '' : n; };
            // ค่าข้าม/ค่าค้าง hold text ("ไม่ข้าม") or an amount — keep amounts numeric so Excel can sum them.
            const textOrNum = (v) => { const s = String(v ?? '').trim(); return s !== '' && isFinite(Number(s)) ? Number(s) : s; };
            const headers = ['ประเภทรถ', 'ต้นทาง', 'ปลายทาง', 'ทิศทาง', 'ค่าข้าม', 'ค่าค้าง', 'รวมน้ำมัน (ลิตร)', 'รวมน้ำมัน (บาท)', ...cols.map(c => c.label)];
            const data = routes.map(r => [
                r.vehicleType || '', r.origin || '', r.destination || '', r.direction || '',
                textOrNum(r.crossingFee), textOrNum(r.waitingFee), num(r.totalLiters), num(r.totalBaht),
                ...cols.map(c => num(c.unit === 'Liters' ? r.gasStationsLiters?.[c.key] : r.gasStationsBaht?.[c.key])),
            ]);
            const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
            ws['!cols'] = headers.map((h, i) => ({ wch: i < 3 ? 22 : Math.max(10, Math.min(h.length + 2, 26)) }));
            ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: data.length, c: headers.length - 1 } }) };
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'เส้นทาง');
            // Any search (even spaces) filters the rows, so mark the file name; keep it file-name safe and short.
            const search = appState.dataManagementFilter;
            const safe = search.replace(/[\\/:*?"<>|]/g, '').trim().slice(0, 30);
            const suffix = search ? `_ค้นหา-${safe || 'กรองแล้ว'}` : '';
            XLSX.writeFile(wb, `เส้นทาง_${localDateStr()}${suffix}.xlsx`);
        };

         const attachVehicleManagementListeners = () => {
            const filterInput = document.getElementById('vehicle-management-filter');
            if (filterInput) {
                filterInput.addEventListener('input', (e) => {
                    appState.vehicleManagementFilter = e.target.value;
                    const content = document.getElementById('vehicle-management-content');
                    if(content) content.innerHTML = renderVehicleManagementContent();
                });
            }
             const refreshBtn = document.getElementById('refresh-vehicles-btn');
            if(refreshBtn) {
                refreshBtn.addEventListener('click', () => {
                    const content = document.getElementById('vehicle-management-content');
                    if(content) content.innerHTML = renderVehicleManagementContent();
                });
            }
        };
        
        const attachAutocompleteListeners = (id, options) => {
            const input = document.getElementById(id);
            const suggestionsContainer = document.getElementById(`${id}-suggestions`);
            if (!input || !suggestionsContainer) return;

            let currentOptions = [];
            let activeIndex = -1;

            const renderList = () => {
                suggestionsContainer.innerHTML = currentOptions
                    .map((opt, i) => `<div data-index="${i}" class="p-2 cursor-pointer ${i === activeIndex ? 'bg-slate-700' : 'hover:bg-slate-700'}">${opt}</div>`)
                    .join('');
            };

            const scrollActiveIntoView = () => {
                const el = suggestionsContainer.querySelector(`[data-index="${activeIndex}"]`);
                if (el) el.scrollIntoView({ block: 'nearest' });
            };

            const choose = (opt) => {
                input.value = opt;
                suggestionsContainer.classList.add('hidden');
                activeIndex = -1;
                input.dispatchEvent(new Event('change')); // trigger vlookup / auto-fill
            };

            const showSuggestions = () => {
                const value = input.value.toLowerCase();
                // Empty input → show ALL options (so clicking/focusing the field opens the list);
                // otherwise filter by what was typed.
                currentOptions = value.length > 0
                    ? options.filter(opt => opt.toLowerCase().includes(value))
                    : options.slice();
                activeIndex = -1;
                if (currentOptions.length > 0) {
                    renderList();
                    suggestionsContainer.classList.remove('hidden');
                } else {
                    suggestionsContainer.classList.add('hidden');
                }
            };

            input.addEventListener('input', showSuggestions);
            input.addEventListener('focus', showSuggestions);

            input.addEventListener('keydown', (e) => {
                const visible = !suggestionsContainer.classList.contains('hidden') && currentOptions.length > 0;
                if (!visible) return;
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    activeIndex = (activeIndex + 1) % currentOptions.length;
                    renderList(); scrollActiveIntoView();
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    activeIndex = (activeIndex - 1 + currentOptions.length) % currentOptions.length;
                    renderList(); scrollActiveIntoView();
                } else if (e.key === 'Enter') {
                    if (activeIndex >= 0) { e.preventDefault(); choose(currentOptions[activeIndex]); }
                } else if (e.key === 'Escape') {
                    suggestionsContainer.classList.add('hidden');
                }
            });

            suggestionsContainer.addEventListener('click', (e) => {
                const item = e.target.closest('[data-index]');
                if (item) choose(currentOptions[Number(item.dataset.index)]);
            });
        };

        // Close any open autocomplete dropdown when clicking outside its container.
        // Attached ONCE at module load — previously this was registered inside
        // attachAutocompleteListeners (6× per form render), so switching views repeatedly
        // accumulated dozens of stale document-level handlers (memory/perf leak).
        let autocompleteOutsideClickAttached = false;
        const attachAutocompleteOutsideClick = () => {
            if (autocompleteOutsideClickAttached) return;
            autocompleteOutsideClickAttached = true;
            document.addEventListener('click', (e) => {
                if (!e.target.closest('.autocomplete-container')) {
                    document.querySelectorAll('[id$="-suggestions"]').forEach(el => el.classList.add('hidden'));
                }
            });
        };
        
        const updateVlookupCard = () => {
            const cardContent = document.getElementById('vlookup-card-content');
            if (!cardContent) return;

            const vehicleType = document.getElementById('form-vehicle-type')?.value;
            const origin = document.getElementById('form-origin')?.value;
            const destination = document.getElementById('form-destination')?.value;

            const route = logisticsService.findMatchingRoute(vehicleType, origin, destination);

            if (!route) {
                cardContent.innerHTML = `
                    <h3 class="text-xl font-bold text-white mb-4">ข้อมูลเส้นทาง</h3>
                    <p class="text-gray-400">ไม่พบข้อมูลเส้นทางที่ตรงกัน</p>
                    <p class="text-sm text-gray-500 mt-2">กรุณากรอกข้อมูลประเภทรถ, ต้นทาง, และปลายทางให้ครบถ้วน</p>
                `;
                return;
            }
            
            const stationList = (stations) => {
                if (!stations || Object.keys(stations).length === 0) return '<p class="text-sm text-gray-500">ไม่มี</p>';
                return Object.entries(stations)
                    .map(([key, value]) => `
                        <div class="flex justify-between items-center text-sm py-1">
                           <span class="text-gray-400">${logisticsService.getFriendlyGasStationName(key)}:</span>
                           <span class="font-semibold">${value.toLocaleString()}</span>
                        </div>
                    `).join('');
            };

            cardContent.innerHTML = `
                <h3 class="text-xl font-bold text-white mb-4">ข้อมูลเส้นทาง (อัตโนมัติ)</h3>
                <div class="space-y-2">
                    <div class="flex justify-between items-center">
                        <span class="font-semibold text-gray-300">ประเภทรถ:</span>
                        <span>${route.vehicleType}</span>
                    </div>
                    <div class="flex justify-between items-center">
                        <span class="font-semibold text-gray-300">เส้นทาง:</span>
                        <span class="text-right">${route.origin} → ${route.destination}</span>
                    </div>
                    <hr class="border-slate-600 my-2">
                    <div class="flex justify-between items-center text-lg">
                        <span class="font-bold text-accent">รวม (ลิตร):</span>
                        <span class="font-bold text-accent">${(route.totalLiters || 0).toLocaleString()}</span>
                    </div>
                     ${stationList(route.gasStationsLiters)}
                     <hr class="border-slate-600 my-2">
                    <div class="flex justify-between items-center text-lg">
                        <span class="font-bold text-accent">รวม (บาท):</span>
                        <span class="font-bold text-accent">${(route.totalBaht || 0).toLocaleString()}</span>
                    </div>
                    ${stationList(route.gasStationsBaht)}
                    <button type="button" data-action="apply-route" class="w-full mt-4 bg-success hover:bg-green-600 text-white font-bold py-2 px-4 rounded-lg shadow-md transition">ใช้ค่าจากเส้นทางนี้</button>
                    <p class="text-xs text-gray-500 mt-1 text-center">เติมจุดเติมน้ำมัน (ลิตร+บาท) ให้อัตโนมัติ</p>
                </div>
            `;
        };

        // Populate the fuel form's distribution rows from the matched route (item 6).
        const handleApplyRouteToForm = () => {
            const vehicleType = document.getElementById('form-vehicle-type')?.value;
            const origin = document.getElementById('form-origin')?.value;
            const destination = document.getElementById('form-destination')?.value;
            const route = logisticsService.findMatchingRoute(vehicleType, origin, destination);
            if (!route) return;

            const fill = (listId, stations, unit) => {
                const container = document.getElementById(listId);
                if (!container) return;
                container.innerHTML = '';
                Object.entries(stations || {}).forEach(([stationKey, amount], index) => {
                    container.insertAdjacentHTML('beforeend', renderRouteDistributionRow({ stationKey, amount }, index, unit));
                });
            };
            fill('distribution-list-liters', route.gasStationsLiters, 'Liters');
            fill('distribution-list-baht', route.gasStationsBaht, 'Baht');
            updateTotalAmount('liters');
            updateTotalAmount('baht');
        };
        
        const updateDistributionList = (unit) => {
            const unitUpper = unit.charAt(0).toUpperCase() + unit.slice(1); // Liters or Baht
            const container = document.getElementById(`distribution-list-${unit}`);
            if (!container) return;
            
            const index = container.children.length;
            const newItemHtml = renderRouteDistributionRow({}, index, unitUpper);
            container.insertAdjacentHTML('beforeend', newItemHtml);
        };
        
        const updateTotalAmount = (unit) => {
            const container = document.getElementById(`distribution-list-${unit}`);
            const totalInput = document.getElementById(`form-total-${unit}`);
            if (!container || !totalInput) return;

            const rows = container.querySelectorAll('.distribution-row');
            let total = 0;
            rows.forEach(row => {
                const amountInput = row.querySelector('input[type="number"]');
                const stationSelect = row.querySelector('select');
                const amount = Number(amountInput?.value) || 0;
                const station = stationSelect?.value || '';
                total += amount;
                // Highlight rows that won't be saved: station without amount, or amount without station.
                const incomplete = (station && amount <= 0) || (!station && amount > 0);
                [amountInput, stationSelect].forEach(el => {
                    if (!el) return;
                    el.classList.toggle('ring-2', incomplete);
                    el.classList.toggle('ring-red-500', incomplete);
                });
            });
            totalInput.value = total > 0 ? total : '';
        };

        const resetFuelForm = () => {
            const form = document.getElementById('fuel-record-form');
            if (form) {
                form.reset();
                // Also reset date to today and clear dynamic lists
                document.getElementById('form-date').value = localDateStr();
                document.getElementById('distribution-list-liters').innerHTML = '';
                document.getElementById('distribution-list-baht').innerHTML = '';
                document.getElementById('form-total-liters').value = '';
                document.getElementById('form-total-baht').value = '';
                updateVlookupCard();
                const status = document.getElementById('form-status');
                if (status) status.innerHTML = '';
            }
        };
        
        // 🧮 for the page form (reads/writes by id). silent = no alert (used for auto-calc on change).
        const handleCalcAllowanceForm = (silent = false) => {
            const queueType = document.getElementById('form-queue-type')?.value || '';
            const vehicleType = document.getElementById('form-vehicle-type')?.value || '';
            const rate = logisticsService.findRate(queueType, vehicleType);
            if (!rate) {
                if (!silent) alert(`ไม่พบอัตราสำหรับ ประเภทคิว "${queueType || '(ว่าง)'}" + ประเภทรถ "${vehicleType || '(ว่าง)'}"\n\nประเภทวิ่งในตาราง: หนักล่อง, ตีเปล่า, คิวพิเศษ (ต้องนำเข้าชีต "ฐานคำนวน" ก่อน)`);
                return;
            }
            document.getElementById('form-allowance').value = rate.allowance;
            document.getElementById('form-deduction').value = rate.deduction;
            document.getElementById('form-net').value = rate.allowance - rate.deduction;
        };

        // Save the page trip-entry form → a trip (trips/{month}).
        const handleTripEntrySubmit = async (e) => {
            e.preventDefault();
            const statusDiv = document.getElementById('form-status');
            const getDist = (unit) => {
                const container = document.getElementById(`distribution-list-${unit}`);
                const data = {};
                container?.querySelectorAll('.distribution-row').forEach(row => {
                    const key = row.querySelector('[data-field="stationKey"]')?.value;
                    const amt = Number(row.querySelector('[data-field="amount"]')?.value) || 0;
                    if (key && amt > 0) data[key] = (data[key] || 0) + amt;
                });
                return data;
            };
            const val = (id) => (document.getElementById(id)?.value || '').trim();
            const num = (id) => Number(document.getElementById(id)?.value) || 0;
            const date = val('form-date');
            const plate = val('form-vehicle-registration');
            if (!date || !plate) { statusDiv.innerHTML = `<p class="text-red-400">กรุณากรอกอย่างน้อย วันที่ + ทะเบียนรถ</p>`; return; }
            const gasStationsLiters = getDist('liters'), gasStationsBaht = getDist('baht');
            const trip = {
                date, yearMonth: date.slice(0, 7),
                direction: val('form-direction') || 'ขาขึ้น',
                plate, vehicleType: val('form-vehicle-type'), driver: val('form-driver'),
                queueType: val('form-queue-type'), origin: val('form-origin'), destination: val('form-destination'),
                jobDetail: val('form-job-detail'), notes: val('form-notes'),
                gasStationsLiters, gasStationsBaht,
                totalLiters: Object.values(gasStationsLiters).reduce((a, b) => a + b, 0),
                totalBaht: Object.values(gasStationsBaht).reduce((a, b) => a + b, 0),
                revenue: num('form-revenue'), netTransfer: num('form-net'), fuelCost: num('form-fuelcost'),
                driverAllowance: num('form-allowance'), deduction: num('form-deduction'), distance: num('form-distance'),
            };
            const submitBtn = e.target.querySelector('button[type="submit"]');
            const orig = submitBtn ? submitBtn.textContent : '';
            if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = 'กำลังบันทึก...'; submitBtn.classList.add('opacity-60'); }
            try {
                await logisticsService.addTrip(trip.yearMonth, trip);
                resetFuelForm();
                statusDiv.innerHTML = `<p class="text-green-400">บันทึกคิวรถสำเร็จ! (${trip.date})</p>`;
                setTimeout(() => statusDiv.innerHTML = '', 3000);
            } catch (error) {
                console.error('Error saving trip:', error);
                statusDiv.innerHTML = `<p class="text-red-400">เกิดข้อผิดพลาด: ${error.message}</p>`;
            } finally {
                if (submitBtn) { submitBtn.disabled = false; submitBtn.textContent = orig; submitBtn.classList.remove('opacity-60'); }
            }
        };

        const handleAuthSubmit = (e) => {
            e.preventDefault();
            const passwordInput = document.getElementById('password');
            const errorP = document.getElementById('auth-error');
            // Simple password check, in a real app this would be more secure.
            if (passwordInput.value === "1234") {
                appState.isAuthenticated = true;
                closeModal();
                if (appState.pendingView) {
                    appState.view = appState.pendingView;
                    appState.pendingView = null;
                    renderApp();
                }
            } else {
                errorP.textContent = "รหัสผ่านไม่ถูกต้อง";
                errorP.classList.remove('hidden');
                passwordInput.focus();
            }
        };

        const handleAddNewRoute = () => {
            openModal('เพิ่มเส้นทางใหม่', renderRouteFormContent());
            document.getElementById('route-form').addEventListener('submit', handleRouteFormSubmit);
        };

        const handleEditRoute = (id) => {
            const route = logisticsService.getRoutes().find(r => r.id === id);
            if (!route) return;
            openModal('แก้ไขข้อมูลเส้นทาง', renderRouteFormContent(route));
            document.getElementById('route-form').addEventListener('submit', handleRouteFormSubmit);
        };
        
        const handleDeleteRoute = async (id) => {
            if (confirm('คุณต้องการลบข้อมูลเส้นทางนี้ใช่หรือไม่?')) {
                const route = logisticsService.getRoutes().find(r => r.id === id);
                try {
                    await logisticsService.deleteRoute(id);
                    // No need to re-render, Firebase listener will do it.
                    if (route) {
                        const { id: _id, ...data } = route;
                        showUndoToast('ลบเส้นทางแล้ว', () => logisticsService.restoreRoute(id, data));
                    }
                } catch (error) {
                    console.error("Failed to delete route:", error);
                }
            }
        };

        const handleRouteFormSubmit = async (e) => {
            e.preventDefault();
            const form = e.target;
            const formData = new FormData(form);
            const id = formData.get('id');

            const getDistributionData = (unit) => {
                const container = form.querySelector(`#route-distribution-list-${unit.toLowerCase()}`);
                if (!container) return {};
                const data = {};
                 container.querySelectorAll('.distribution-row').forEach(row => {
                    const stationKey = row.querySelector('[data-field="stationKey"]').value;
                    const amount = Number(row.querySelector('[data-field="amount"]').value) || 0;
                    if (stationKey && amount > 0) {
                        data[stationKey] = (data[stationKey] || 0) + amount;
                    }
                });
                return data;
            };

            const data = {
                vehicleType: formData.get('vehicleType'),
                origin: formData.get('origin'),
                destination: formData.get('destination'),
                direction: formData.get('direction') || 'ขาขึ้น',
                crossingFee: formData.get('crossingFee') || '',
                waitingFee: formData.get('waitingFee') || '',
                gasStationsLiters: getDistributionData('Liters'),
                gasStationsBaht: getDistributionData('Baht'),
            };
            
            data.totalLiters = Object.values(data.gasStationsLiters).reduce((sum, val) => sum + val, 0);
            data.totalBaht = Object.values(data.gasStationsBaht).reduce((sum, val) => sum + val, 0);

            // Required fields
            if (!(data.vehicleType || '').trim() || !(data.origin || '').trim() || !(data.destination || '').trim()) {
                alert('กรุณากรอก ประเภทรถ, ต้นทาง และปลายทาง ให้ครบถ้วน');
                return;
            }

            // Duplicate route guard: same ประเภทรถ + ต้นทาง + ปลายทาง (ignore the route being edited itself)
            const duplicate = logisticsService.getRoutes().find(r =>
                r.id !== id &&
                (r.vehicleType || '').trim() === data.vehicleType.trim() &&
                (r.origin || '').trim() === data.origin.trim() &&
                (r.destination || '').trim() === data.destination.trim()
            );
            if (duplicate) {
                alert(`มีเส้นทางนี้อยู่แล้ว:\n\n${data.vehicleType}\n${data.origin} → ${data.destination}\n\nกรุณาแก้ไขเส้นทางเดิม หรือเปลี่ยนข้อมูลให้ไม่ซ้ำกัน`);
                return;
            }

            try {
                if (id) {
                    await logisticsService.updateRoute({ id, ...data });
                } else {
                    await logisticsService.addRoute(data);
                }
                closeModal();
                // No need to re-render, Firebase listener handles it.
            } catch (error) {
                console.error("Error saving route:", error);
                alert("เกิดข้อผิดพลาดในการบันทึกข้อมูลเส้นทาง: " + error.message);
            }
        };
        
        const handleAddRouteDistribution = (button) => {
            const unit = button.dataset.unit; // 'Liters' or 'Baht'
            const containerId = `route-distribution-list-${unit.toLowerCase()}`;
            const container = document.getElementById(containerId);
            if (!container) return;
            
            const index = container.children.length;
            const newItemHtml = renderRouteDistributionRow({}, index, unit);
            container.insertAdjacentHTML('beforeend', newItemHtml);
        };
        
        const handleAddNewVehicle = () => {
            openModal('เพิ่มทะเบียนรถใหม่', renderVehicleFormContent());
            document.getElementById('vehicle-form').addEventListener('submit', handleVehicleFormSubmit);
        };
        
        const handleEditVehicle = (id) => {
            const vehicle = logisticsService.getVehicles().find(v => v.id === id);
            if (!vehicle) return;
            openModal('แก้ไขข้อมูลทะเบียนรถ', renderVehicleFormContent(vehicle));
            document.getElementById('vehicle-form').addEventListener('submit', handleVehicleFormSubmit);
        };

        const handleDeleteVehicle = async (id) => {
            if (confirm('คุณต้องการลบข้อมูลทะเบียนรถนี้ใช่หรือไม่?')) {
                const vehicle = logisticsService.getVehicles().find(v => v.id === id);
                try {
                    await logisticsService.deleteVehicle(id);
                    if (vehicle) {
                        const { id: _id, ...data } = vehicle;
                        showUndoToast('ลบทะเบียนรถแล้ว', () => logisticsService.restoreVehicle(id, data));
                    }
                } catch (error) {
                    console.error("Failed to delete vehicle:", error);
                }
            }
        };

        const handleVehicleFormSubmit = async (e) => {
            e.preventDefault();
            const form = e.target;
            const formData = new FormData(form);
            const id = formData.get('id');
            const data = {
                registrationNumber: formData.get('registrationNumber'),
                vehicleType: formData.get('vehicleType'),
                driveType: formData.get('driveType') || '',
                company: formData.get('company') || '',
                model: formData.get('model') || '',
            };

            try {
                if (id) {
                    await logisticsService.updateVehicle({ id, ...data });
                } else {
                    await logisticsService.addVehicle(data);
                }
                closeModal();
            } catch (error) {
                console.error("Error saving vehicle:", error);
                alert("เกิดข้อผิดพลาดในการบันทึกข้อมูล: " + error.message);
            }
        };
        
        const handleEditFuelRecord = (id) => {
            const record = logisticsService.getFuelRecords().find(r => r.id === id);
            if (!record) return;
            openModal('แก้ไขรายการน้ำมัน', renderFuelRecordEditFormContent(record));
            document.getElementById('fuel-record-edit-form').addEventListener('submit', handleFuelRecordEditFormSubmit);
        };
        
        

const handleDeleteFuelRecord = async (id) => {
    if (confirm('คุณต้องการลบรายการน้ำมันนี้ใช่หรือไม่?')) {
        const record = logisticsService.getFuelRecords().find(r => r.id === id);
        try {
            await logisticsService.deleteFuelRecord(id);
            if (record) {
                const { id: _id, ...data } = record;
                showUndoToast('ลบรายการน้ำมันแล้ว', () => logisticsService.restoreFuelRecord(id, data));
            }
        } catch (error) {
            console.error("Failed to delete fuel record:", error);
            alert("เกิดข้อผิดพลาดในการลบข้อมูล: " + error.message);
        }
    }
};
const handleFuelRecordEditFormSubmit = async (e) => {
            e.preventDefault();
            const form = e.target;
            const formData = new FormData(form);
            const id = formData.get('id');

            const getDistributionData = (unit) => {
                const container = form.querySelector(`#edit-distribution-list-${unit.toLowerCase()}`);
                if (!container) return {};
                const data = {};
                container.querySelectorAll('.distribution-row').forEach(row => {
                    const stationKey = row.querySelector('[data-field="stationKey"]')?.value;
                    const amount = Number(row.querySelector('[data-field="amount"]')?.value) || 0;
                    if (stationKey && amount > 0) {
                        data[stationKey] = (data[stationKey] || 0) + amount;
                    }
                });
                return data;
            };

            const data = {
                date: formData.get('date'),
                vehicleRegistration: formData.get('vehicleRegistration'),
                vehicleType: formData.get('vehicleType'),
                origin: formData.get('origin'),
                destination: formData.get('destination'),
                notes: formData.get('notes'),
                fuelDistributionLiters: getDistributionData('Liters'),
                fuelDistributionBaht: getDistributionData('Baht'),
            };
            
            data.totalLiters = Object.values(data.fuelDistributionLiters).reduce((sum, val) => sum + val, 0);
            data.totalBaht = Object.values(data.fuelDistributionBaht).reduce((sum, val) => sum + val, 0);

            try {
                await logisticsService.updateFuelRecord({ id, ...data });
                closeModal();
            } catch (error) {
                console.error("Error updating fuel record:", error);
                alert("เกิดข้อผิดพลาดในการอัปเดตข้อมูล: " + error.message);
            }
        };
        
        const handleAddEditDistribution = (button) => {
            const unit = button.dataset.unit; // 'Liters' or 'Baht'
            const containerId = `edit-distribution-list-${unit.toLowerCase()}`;
            const container = document.getElementById(containerId);
            if (!container) return;
            
            const index = container.children.length;
            const newItemHtml = renderEditDistributionRow({}, index, unit);
            container.insertAdjacentHTML('beforeend', newItemHtml);
        };
        
        const handleBackupData = () => {
            // Full-fidelity backup of the entire dataset as a JSON file. Can be kept as a safety
            // copy and re-imported manually if data is ever lost.
            const backup = {
                app: 'NeoSiamLogistics & Transport Fuel',
                version: 1,
                exportedAt: new Date().toISOString(),
                routes: logisticsService.getRoutes(),
                fuelRecords: logisticsService.getFuelRecords(),
                vehicles: logisticsService.getVehicles(),
            };
            const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `NeoSiamLogistics_backup_${localDateStr()}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        };

        const handleExportToExcel = () => {
            const filters = appState.dashboardFilters;
            const day = filters.day || '';
            const plateFilter = (filters.vehicleRegistration || '').toUpperCase();
            let trips = logisticsService.getTrips();
            if (day) trips = trips.filter(t => t.date === day);
            if (plateFilter) trips = trips.filter(t => (t.plate || '').toUpperCase().includes(plateFilter));
            if (!trips.length) { alert('ไม่มีข้อมูลให้ export'); return; }

            const stationKeys = logisticsService.getAllGasStationKeys();
            const friendlyNames = logisticsService.getAllGasStationFriendlyNames();
            const headers = [
                'วันที่', 'ทิศทาง', 'ทะเบียน', 'ประเภทรถ', 'คนขับ', 'ต้นทาง', 'ปลายทาง', 'รวม(ลิตร)', 'รวม(บาท)',
                ...stationKeys.map(key => `${friendlyNames[key]} (ลิตร)`),
                ...stationKeys.map(key => `${friendlyNames[key]} (บาท)`),
                'หมายเหตุ',
            ];
            const data = trips.map(t => [
                t.date, t.direction || '', t.plate || '', t.vehicleType || '', t.driver || '', t.origin || '', t.destination || '',
                t.totalLiters || '', t.totalBaht || '',
                ...stationKeys.map(key => t.gasStationsLiters?.[key] || ''),
                ...stationKeys.map(key => t.gasStationsBaht?.[key] || ''),
                t.notes || '',
            ]);
            const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'FuelByPump');
            XLSX.writeFile(wb, `น้ำมันรายปั๊ม_${day || appState.tripsMonth || localDateStr().slice(0, 7)}.xlsx`);
        };
        
        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        // 6. APP INITIALIZATION
        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
        // Wire the UI callbacks into the data layer, then start the app.
        bindUI({ render: renderApp, hideLoading: hideLoadingOverlay });
        logisticsService.initialize();
