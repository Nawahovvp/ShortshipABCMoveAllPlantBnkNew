import { plantInfo } from './config.js';

export function formatShortCurrency(n) {
    if (n >= 1e9) return (n / 1e9).toFixed(2) + " พันล้าน";
    if (n >= 1e6) return (n / 1e6).toFixed(2) + " M";
    if (n >= 1e3) return (n / 1e3).toFixed(2) + " K";
    return n.toFixed(2);
}

export function toNumber(v) { return parseFloat((v || "0").toString().replace(/,/g, '')) || 0; }
export function formatNumber(n) { return Number(n).toLocaleString('th-TH'); }
export function formatCurrency(n) { return Number(n).toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

export function getPlantName(code) {
    return plantInfo[code]?.name || "";
}

export function getPlantLabel(code) {
    const name = getPlantName(code);
    return name ? `${code} ${name}` : code;
}
export function getPlantType(code) {
    const label = getPlantLabel(code);
    return label.includes("SA") ? "SA" : "Company";
}