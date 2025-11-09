import { computeScopeStage } from "../utils/scope";

describe("computeScopeStage", () => {
  it("A only => 1", () => {
    expect(computeScopeStage({ facilityA: "ZRH21" })).toBe("1");
  });
  it("A + Diversity + Splice A => 2", () => {
    expect(computeScopeStage({ facilityA: "ZRH21", diversity: "East", spliceA: "AM111" })).toBe("2");
  });
  it("A + Diversity => 3", () => {
    expect(computeScopeStage({ facilityA: "ZRH21", diversity: "East" })).toBe("3");
  });
  it("A + Splice A => 4", () => {
    expect(computeScopeStage({ facilityA: "ZRH21", spliceA: "AM111" })).toBe("4");
  });
  it("Z + Diversity + Splice Z => 5", () => {
    expect(computeScopeStage({ facilityZ: "ZRH20", diversity: "East", spliceZ: "AJ1508" })).toBe("5");
  });
  it("Z + Diversity => 6", () => {
    expect(computeScopeStage({ facilityZ: "ZRH20", diversity: "East" })).toBe("6");
  });
  it("Z + Splice Z => 7", () => {
    expect(computeScopeStage({ facilityZ: "ZRH20", spliceZ: "AJ1508" })).toBe("7");
  });
  it("Z only => 8", () => {
    expect(computeScopeStage({ facilityZ: "ZRH20" })).toBe("8");
  });
});
