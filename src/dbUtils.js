// Shared helpers for writing imports to Firebase without overwriting anything (pure; the write function is passed in).

/**
 * Write each item with a create-only function (a Firebase transaction that aborts when the key exists), in batches.
 * createFn(item) → Promise<boolean> (true = written, false = the key already existed). One failure does not stop the
 * rest; saving again later only fills in what is missing.
 * Returns { added, skipped, failed, error, addedItems }.
 */
export async function createAllIfAbsent(items, createFn, concurrency = 10) {
    let skipped = 0, failed = 0, error = null;
    const addedItems = [];
    for (let i = 0; i < items.length; i += concurrency) {
        const batch = items.slice(i, i + concurrency);
        const res = await Promise.allSettled(batch.map(createFn));
        res.forEach((r, j) => {
            if (r.status === 'rejected') { failed++; error = error || r.reason; }
            else if (r.value) addedItems.push(batch[j]);
            else skipped++;
        });
    }
    return { added: addedItems.length, skipped, failed, error, addedItems };
}
