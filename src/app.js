// UI layer: rendering, event handlers, and app bootstrap.
import "./index.css";
import * as XLSX from "xlsx";
import { appState, logisticsService, bindUI } from "./services.js";
import { ICONS, GAS_STATION_FRIENDLY_NAMES, ROUTE_STATION_COLUMNS } from "./constants.js";
import { parseFuelBaseRows } from "./importer.js";
import { detectSheetType, parseUphillQueue, parseDownhillQueue, parsePartnerQueue, parseDailyRows } from "./importerTrips.js";
import { parseDrivers, parseRateTable, parseVehicles } from "./importerMasters.js";

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
                    <nav class="container mx-auto px-6 py-3 flex justify-between items-center">
                        <div class="flex items-center gap-3">
                            <div class="bg-gray-200 p-2 rounded-full">
                               <svg xmlns="http://www.w3.org/2000/svg" class="w-8 h-8" viewBox="0 0 24 24" fill="#0369a1"><path d="M12 22a7 7 0 0 0 7-7c0-2-1-3.9-3-5.5s-3.5-4-4-6.5c-.5 2.5-2 4.9-4 6.5C6 11.1 5 13 5 15a7 7 0 0 0 7 7z"></path></svg>
                            </div>
                            <h1 class="text-xl sm:text-2xl font-bold text-white tracking-wider">NeoSiamLogistics & Transport Fuel</h1>
                        </div>
                        <div class="flex items-center gap-2 sm:gap-4">
                            ${navButton('form', ICONS.Truck, 'บันทึกน้ำมัน')}
                            ${navButton('dashboard', ICONS.ChartBar, 'แดชบอร์ด')}
                            ${navButton('queue', ICONS.Truck, 'คิวรถ')}
                            ${navButton('payroll', ICONS.ChartBar, 'การเงิน')}
                            ${navButton('reports', ICONS.ChartBar, 'รายงาน')}
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

        const renderDataManagementContent = () => {
            const filter = appState.dataManagementFilter.toLowerCase();
            const routes = logisticsService.getRoutes();
            const filteredRoutes = filter ? routes.filter(r => 
                (r.vehicleType || '').toLowerCase().includes(filter) ||
                (r.origin || '').toLowerCase().includes(filter) ||
                (r.destination || '').toLowerCase().includes(filter)
            ) : routes;

            // Sort routes for consistent display
            filteredRoutes.sort((a, b) => {
                const vehicleTypeCompare = (a.vehicleType || '').localeCompare(b.vehicleType || '', 'th');
                if (vehicleTypeCompare !== 0) return vehicleTypeCompare;

                const originCompare = (a.origin || '').localeCompare(b.origin || '', 'th');
                if (originCompare !== 0) return originCompare;

                return (a.destination || '').localeCompare(b.destination || '', 'th');
            });

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
                            <div class="flex items-center gap-4">
                                <input type="text" id="data-management-filter" value="${appState.dataManagementFilter}" placeholder="ค้นหา..." class="w-full md:w-64 bg-slate-800 border border-slate-600 rounded-md px-3 py-2 text-white focus:outline-none focus:ring-2 focus:ring-accent" />
                                <button id="refresh-routes-btn" title="รีเฟรชข้อมูล" class="bg-sky-600 hover:bg-sky-700 text-white font-bold p-2 rounded-lg shadow-md transition-transform transform hover:scale-105">
                                    ${ICONS.Refresh}
                                </button>
                                <button id="backup-data-btn" title="สำรองข้อมูลทั้งหมด (เส้นทาง + บันทึกน้ำมัน + ทะเบียนรถ) เป็นไฟล์ JSON" class="flex items-center gap-2 bg-success hover:bg-green-600 text-white font-bold py-2 px-4 rounded-lg shadow-md transition-transform transform hover:scale-105">
                                    ${ICONS.Download} สำรองข้อมูล
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

        const renderImport = () => {
            return `
                <div class="max-w-5xl mx-auto">
                    ${renderCard(`
                        <h2 class="text-2xl font-bold text-white mb-2 flex items-center gap-3">${ICONS.Download} นำเข้าเส้นทางจาก Excel</h2>
                        <p class="text-gray-400 mb-6">อัปโหลดไฟล์ Excel แล้วเลือกชีต — ระบบตรวจชนิดให้อัตโนมัติ:<br>• <span class="text-accent">ฐานคำนวนน้ำมันขาขึ้น</span> → เส้นทาง &nbsp;•&nbsp; <span class="text-accent">คิวขาขึ้น/ขาล่อง, คิวรถร่วม</span> → คิวรถ &nbsp;•&nbsp; <span class="text-accent">ทะเบียน, คนขับ, ฐานคำนวน</span> → รถ/คนขับ/เบี้ยเลี้ยง</p>
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
                const def = importWorkbook.SheetNames.find(n => n.includes('ฐานคำนวนน้ำมันขาขึ้น')) || importWorkbook.SheetNames.find(n => n.includes('ฐานคำนวน'));
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
            const kind = detectSheetType(rows);
            importParsedRoutes = null;
            importParsedTrips = null;

            if (kind === 'fuelbase') {
                const result = parseFuelBaseRows(rows);
                importParsedRoutes = result.routes;
                preview.innerHTML = renderImportPreview(result);
                document.getElementById('import-replace-btn')?.addEventListener('click', () => handleImportCommit('replace'));
                document.getElementById('import-merge-btn')?.addEventListener('click', () => handleImportCommit('merge'));
            } else if (kind === 'uphill' || kind === 'downhill' || kind === 'partnerqueue') {
                const result = kind === 'uphill' ? parseUphillQueue(rows) : kind === 'downhill' ? parseDownhillQueue(rows) : parsePartnerQueue(rows);
                importParsedTrips = result.trips;
                preview.innerHTML = renderTripsImportPreview(result, kind);
                document.getElementById('import-trips-replace-btn')?.addEventListener('click', () => handleImportTripsCommit('replace'));
                document.getElementById('import-trips-merge-btn')?.addEventListener('click', () => handleImportTripsCommit('merge'));
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
                ? `แทนที่คิวรถทั้งหมดด้วย ${n.toLocaleString()} คิวจาก Excel?\n\n⚠️ คิวรถเดิมทุกเดือนจะถูกลบ (ควรสำรองข้อมูลก่อน)`
                : `เพิ่ม ${n.toLocaleString()} คิวเข้าไป (ไม่ลบของเดิม)?`;
            if (!confirm(msg)) return;
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
                    if (mode === 'replace') await logisticsService.replaceAllVehicles(items);
                    else await logisticsService.addVehiclesBatch(items);
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

        const attachImportListeners = () => {
            document.getElementById('import-file')?.addEventListener('change', handleImportFile);
            document.getElementById('import-parse-btn')?.addEventListener('click', handleImportParse);
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
                    if (originalMonth && originalMonth !== trip.yearMonth) {
                        // Moving to another month: add to the new month FIRST, then remove the old
                        // one. If the add fails we throw before deleting, so the trip is never lost.
                        await logisticsService.addTrip(trip.yearMonth, trip);
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
            const trips = logisticsService.getTrips();
            if (!trips.length) return `<p class="text-center py-8 text-gray-500">ไม่มีข้อมูล — เลือกเดือนที่นำเข้าคิวรถแล้ว</p>`;

            const sum = (arr, f) => arr.reduce((s, t) => s + (f(t) || 0), 0);
            const totLiters = sum(trips, t => t.totalLiters), totBaht = sum(trips, t => t.totalBaht);
            const totRevenue = sum(trips, t => t.revenue), totNet = sum(trips, t => t.netTransfer);

            // group helper → [{key, count, net}] sorted by net desc
            const groupBy = (keyFn) => {
                const m = {};
                trips.forEach(t => { const k = keyFn(t) || '(ไม่ระบุ)'; if (!m[k]) m[k] = { key: k, count: 0, net: 0 }; m[k].count++; m[k].net += t.netTransfer || 0; });
                return Object.values(m).sort((a, b) => b.net - a.net);
            };
            const byDir = groupBy(t => t.direction);
            const byType = groupBy(t => t.vehicleType);
            const topDrivers = groupBy(t => t.driver).slice(0, 10);

            // Own vs partner (join plate → vehicle.driveType)
            const driveOf = {}; logisticsService.getVehicles().forEach(v => { driveOf[(v.registrationNumber || '').trim()] = v.driveType || ''; });
            const ownPartner = { 'รถบริษัท': { count: 0, net: 0 }, 'รถร่วม': { count: 0, net: 0 }, 'ไม่ทราบ': { count: 0, net: 0 } };
            trips.forEach(t => { const d = driveOf[(t.plate || '').trim()]; const k = d && d.includes('ร่วม') ? 'รถร่วม' : d ? 'รถบริษัท' : 'ไม่ทราบ'; ownPartner[k].count++; ownPartner[k].net += t.netTransfer || 0; });

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

            // Fuel by station (liters)
            const stationTotals = {};
            trips.forEach(t => Object.entries(t.gasStationsLiters || {}).forEach(([k, v]) => { stationTotals[k] = (stationTotals[k] || 0) + v; }));
            const stationRows = Object.entries(stationTotals).sort((a, b) => b[1] - a[1])
                .map(([k, v]) => `<div class="flex justify-between text-sm py-1 border-b border-slate-700/50"><span class="text-gray-300">${logisticsService.getFriendlyGasStationName(k)}</span><span class="font-mono text-accent">${v.toLocaleString()} ล.</span></div>`).join('');

            return `
                <p class="text-xs text-gray-500 mb-3">เดือน ${appState.tripsMonth || '(เลือกเดือน)'} · ${trips.length.toLocaleString()} คิว</p>
                <div class="grid grid-cols-2 md:grid-cols-5 gap-3 mb-6">
                    ${importStatCard('จำนวนคิว', trips.length)}
                    ${importStatCard('น้ำมัน (ลิตร)', totLiters)}
                    ${importStatCard('น้ำมัน (บาท)', totBaht)}
                    ${importStatCard('รายรับรวม', totRevenue)}
                    ${importStatCard('สุทธิโอนรวม', totNet)}
                </div>
                <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    ${renderCard(`<h3 class="font-bold text-white mb-3">ตามทิศทาง</h3><div class="space-y-2">${bars(byDir, 'w-20')}</div>`)}
                    ${renderCard(`<h3 class="font-bold text-white mb-3">รถบริษัท vs รถร่วม</h3><div class="space-y-2">${bars(Object.entries(ownPartner).filter(([, v]) => v.count).map(([k, v]) => ({ key: k, count: v.count, net: v.net })), 'w-20')}</div>`)}
                    ${renderCard(`<h3 class="font-bold text-white mb-3">ตามประเภทรถ</h3><div class="space-y-2">${bars(byType, 'w-28')}</div>`)}
                    ${renderCard(`<h3 class="font-bold text-white mb-3">คนขับสุทธิโอนสูงสุด (Top 10)</h3><div class="space-y-2">${bars(topDrivers, 'w-32')}</div>`)}
                    ${renderCard(`<h3 class="font-bold text-white mb-3">น้ำมันรายปั๊ม (ลิตร)</h3>${stationRows || '<p class="text-gray-500 text-sm">ไม่มีข้อมูล</p>'}`)}
                </div>`;
        };
        const attachReportsListeners = () => {
            const monthInput = document.getElementById('reports-month');
            const target = appState.tripsMonth || (monthInput ? monthInput.value : '');
            if (target && appState.tripsMonth !== target) logisticsService.loadTripsForMonth(target);
            monthInput?.addEventListener('change', (e) => logisticsService.loadTripsForMonth(e.target.value));
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
                    <td class="p-3">${vehicle.vehicleType || ''}</td>
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
                if (['data', 'vehicles', 'import', 'queue', 'drivers', 'payroll', 'reports'].includes(view) && !appState.isAuthenticated) {
                    openAuthModal();
                    // Store the intended view to navigate to after successful authentication
                    appState.pendingView = view; 
                } else {
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
