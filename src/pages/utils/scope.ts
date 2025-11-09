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
