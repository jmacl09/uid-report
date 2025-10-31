// Derive "Line" number from StartHardwareSku and C0 port.
// This implements the mapping algorithm for known Cisco SKUs based on provided mapping data.

export function deriveLineForC0(startHardwareSku: string, c0Port: string): number | undefined {
    if (!startHardwareSku || !c0Port) return undefined;

    const sku = String(startHardwareSku).toLowerCase();
    const port = String(c0Port).trim();

    // Helper to parse port segments like "0/3/1" or "1/15"
    const segs = port.split("/").map((s) => (s === "" ? NaN : Number(s)));
    const hasAllNums = segs.every((n) => Number.isFinite(n));
    if (!hasAllNums) return undefined;

    // SKU families we'll handle: 4351, 3945, 2921
    const is4351 = sku.includes("4351");
    const is3945 = sku.includes("3945");
    const is2921 = sku.includes("2921");

    // Cisco 4351 rules
    // Patterns from mapping:
    //  - 0/1/x => 2..25  (2 + x)
    //  - 0/2/x => 26..49 (26 + x)
    //  - 0/3/x => 50..73 (50 + x)
    //  - 1/0/x => 98..121 (98 + x)
    //  - 2/0/x => 194..217 (194 + x)
    if (is4351 && segs.length === 3) {
        const [shelf, slot, portNum] = segs as [number, number, number];
        if (shelf === 0 && slot === 1 && portNum >= 0 && portNum <= 23) return 2 + portNum;
        if (shelf === 0 && slot === 2 && portNum >= 0 && portNum <= 23) return 26 + portNum;
        if (shelf === 0 && slot === 3 && portNum >= 0 && portNum <= 23) return 50 + portNum;
        if (shelf === 1 && slot === 0 && portNum >= 0 && portNum <= 23) return 98 + portNum;
        if (shelf === 2 && slot === 0 && portNum >= 0 && portNum <= 23) return 194 + portNum;
        return undefined;
    }

    // Cisco 3945 rules
    //  - 1/0..1/31 => 67..98 (67 + idx)
    //  - 2/0..2/31 => 131..162 (131 + idx)
    //  - 3/0..3/31 => 195..226 (195 + idx)
    //  - 4/0..4/31 => 259..290 (259 + idx)
    if (is3945 && segs.length === 2) {
        const [slot, idx] = segs as [number, number];
        if (slot === 1 && idx >= 0 && idx <= 31) return 67 + idx;
        if (slot === 2 && idx >= 0 && idx <= 31) return 131 + idx;
        if (slot === 3 && idx >= 0 && idx <= 31) return 195 + idx;
        if (slot === 4 && idx >= 0 && idx <= 31) return 259 + idx;
        return undefined;
    }

    // Cisco 2921 rules
    //  - 0/0/0..15 => 3..18 (3 + idx)
    //  - 0/1/0..15 => 19..34 (19 + idx)
    //  - 0/2/0..15 => 35..50 (35 + idx)
    //  - 0/3/0..15 => 51..66 (51 + idx)
    //  - 1/0..1/31 => 67..98 (67 + idx)
    if (is2921) {
        if (segs.length === 3) {
            const [shelf, slot, idx] = segs as [number, number, number];
            if (shelf === 0 && slot === 0 && idx >= 0 && idx <= 15) return 3 + idx;
            if (shelf === 0 && slot === 1 && idx >= 0 && idx <= 15) return 19 + idx;
            if (shelf === 0 && slot === 2 && idx >= 0 && idx <= 15) return 35 + idx;
            if (shelf === 0 && slot === 3 && idx >= 0 && idx <= 15) return 51 + idx;
            return undefined;
        }
        if (segs.length === 2) {
            const [slot, idx] = segs as [number, number];
            if (slot === 1 && idx >= 0 && idx <= 31) return 67 + idx;
            return undefined;
        }
    }

    return undefined;
}

export default deriveLineForC0;