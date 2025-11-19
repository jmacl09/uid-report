export const computeScopeStage = (vals: {
  facilityA?: string;
  facilityZ?: string;
  diversity?: string | undefined;
  spliceA?: string | undefined;
  spliceZ?: string | undefined;
}): string => {
  const hasA = !!(vals.facilityA && vals.facilityA.trim());
  const hasZ = !!(vals.facilityZ && vals.facilityZ.trim());
  const hasDiv = !!(vals.diversity && vals.diversity.toString().trim());
  const hasSpA = !!(vals.spliceA && vals.spliceA.toString().trim());
  const hasSpZ = !!(vals.spliceZ && vals.spliceZ.toString().trim());
  // Both A and Z present -> use the new facility-pair stage mapping
  if (hasA && hasZ) {
    // Priority: both diversity + both splices -> 14
    if (hasDiv && hasSpA && hasSpZ) return "14";
    // Both splices -> 15
    if (hasSpA && hasSpZ) return "15";
    // Both present with diversity -> 13
    if (hasDiv) return "13";
    // Only both facility codes -> 12
    return "12";
  }

  if (hasA && !hasZ) {
    if (hasDiv && hasSpA) return "2";
    if (hasDiv) return "3";
    if (hasSpA) return "4";
    return "1";
  }
  if (hasZ && !hasA) {
    if (hasDiv && hasSpZ) return "5";
    if (hasDiv) return "6";
    if (hasSpZ) return "7";
    return "8";
  }
  return "1";
};
