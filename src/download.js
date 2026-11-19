// One way to hand the user a file (pattern from the subcontractor-truck project, 19 Nov 2026):
// the link is attached to the page before click() (Firefox ignores a detached link) and the object URL is freed
// only after 60 s — freeing it sooner can make a big download (backups are 10–20 MB) fail without any error.

const REVOKE_DELAY_MS = 60_000;

export function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    try {
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    } finally {
        setTimeout(() => URL.revokeObjectURL(url), REVOKE_DELAY_MS);
    }
}

export const downloadJson = (value, fileName) => downloadBlob(new Blob([JSON.stringify(value)], { type: 'application/json' }), fileName);
