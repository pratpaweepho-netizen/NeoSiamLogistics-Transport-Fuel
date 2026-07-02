// Constants & configuration (gas stations, Firebase config, icons).
export const GAS_STATION_KEYS = [
            'nissan_singburi',
            'jittamas_bangpain',
            'srinakorn_klongklung',
            'daopaisan_nakhonsawan',
            'ekawitat_nonthaburi',
            'daourai_sai3',
            'caltex_sahapat',
            'tekyong_chonburi',
            'rungpattana_wangnoi',
        ];

export const GAS_STATION_FRIENDLY_NAMES = {
            'nissan_singburi': 'นิสสัน (สิงห์บุรี)',
            'jittamas_bangpain': 'จิตตมาศ (บางประอิน)',
            'srinakorn_klongklung': 'ศรีนคร (คลองขลุง)',
            'daopaisan_nakhonsawan': 'ดาวไพศาล (นครสวรรค์)',
            'ekawitat_nonthaburi': 'เอกวิทัศน์ (นนทบุรี)',
            'daourai_sai3': 'ดาวอุไร (สาย 3)',
            'caltex_sahapat': 'คาลเท็กซ์ เครือสหพัฒน์',
            'tekyong_chonburi': 'เต็กย้งทวีโชค (ชลบุรี)',
            'rungpattana_wangnoi': 'รุ่งพัฒนาดี (วังน้อย)',
        };

// Column layout for the route-management table (data-driven so new stations/units appear automatically).
// units: which measures each station tracks — 'Liters' and/or 'Baht'.
export const ROUTE_STATION_COLUMNS = [
            { key: 'nissan_singburi',     units: ['Liters'] },
            { key: 'jittamas_bangpain',   units: ['Liters', 'Baht'] },
            { key: 'srinakorn_klongklung', units: ['Liters'] },
            { key: 'daopaisan_nakhonsawan', units: ['Liters'] },
            { key: 'ekawitat_nonthaburi', units: ['Liters'] },
            { key: 'daourai_sai3',        units: ['Liters', 'Baht'] },
            { key: 'caltex_sahapat',      units: ['Liters', 'Baht'] },
            { key: 'tekyong_chonburi',    units: ['Liters', 'Baht'] },
            { key: 'rungpattana_wangnoi', units: ['Liters'] },
        ];

export const firebaseConfig = {
          apiKey: "AIzaSyDnlm_o9uJQ-uYI4Nv0w-Apmz1d3JQxCQw",
          authDomain: "fleet-fuel-management-system.firebaseapp.com",
          databaseURL: "https://fleet-fuel-management-system-default-rtdb.asia-southeast1.firebasedatabase.app",
          projectId: "fleet-fuel-management-system",
          storageBucket: "fleet-fuel-management-system.appspot.com",
          messagingSenderId: "870197166069",
          appId: "1:870197166069:web:e97f0a6fadb034ae912abc"
        };
        
export const ICONS = {
            Truck: `<svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" viewBox="0 0 24 24" fill="#eab308" stroke="#eab308" stroke-width="1" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22a7 7 0 0 0 7-7c0-2-1-3.9-3-5.5s-3.5-4-4-6.5c-.5 2.5-2 4.9-4 6.5C6 11.1 5 13 5 15a7 7 0 0 0 7 7z"></path></svg>`,
            ChartBar: `<svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"></path><path d="M18 17V9"></path><path d="M13 17V5"></path><path d="M8 17v-3"></path></svg>`,
            Database: `<svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="5" rx="9" ry="3"></ellipse><path d="M3 5V19A9 3 0 0 0 21 19V5"></path><path d="M3 12A9 3 0 0 0 21 12"></path></svg>`,
            IdCard: `<svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 10.5V6a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2v12c0 1.1.9 2 2 2h12.5"></path><path d="M16 19h6"></path><path d="M19 16v6"></path><path d="M6 11h4"></path><path d="M6 15h2"></path><path d="M10 15h2"></path></svg>`,
            DocumentAdd: `<svg xmlns="http://www.w3.org/2000/svg" class="w-7 h-7 text-accent" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 13h6m-3-3v6m5 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>`,
            Trash: `<svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" /></svg>`,
            Pencil: `<svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15.232 5.232l3.536 3.536m-2.036-5.036a2.5 2.5 0 113.536 3.536L6.5 21.036H3v-3.5L15.232 5.232z" /></svg>`,
            Plus: `<svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4" /></svg>`,
            Download: `<svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" /></svg>`,
            Refresh: `<svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h5m-5-5a9 9 0 0115.54 5.54M20 20v-5h-5m5 5a9 9 0 01-15.54-5.54" /></svg>`,
            Clear: `<svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 14l2-2m0 0l2-2m-2 2l-2-2m2 2l2 2m7-2a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>`,
        };
