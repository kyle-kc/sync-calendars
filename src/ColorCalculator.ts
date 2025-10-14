type LAB = readonly [l: number, a: number, b: number];

export class ColorCalculator {
  private static readonly COLOR_ID_TO_HEX: Record<string, string> = {
    "1": "#a4bdfc", // Lavender
    "2": "#7ae7bf", // Sage
    "3": "#dbadff", // Grape
    "4": "#ff887c", // Flamingo
    "5": "#fbd75b", // Banana
    "6": "#ffb878", // Tangerine
    "7": "#46d6db", // Peacock
    "8": "#e1e1e1", // Graphite
    "9": "#5484ed", // Blueberry
    "10": "#51b749", // Basil
    "11": "#dc2127", // Tomato
  };

  private readonly hexCodeToClosestColorIdCache = new Map<string, string>();

  getClosestColorId(targetColorHex: string): string {
    if (!this.hexCodeToClosestColorIdCache.has(targetColorHex)) {
      const closestColorId = Object.entries(
        ColorCalculator.COLOR_ID_TO_HEX,
      ).reduce<{
        id: string | undefined;
        distance: number;
      }>(
        (closest, [colorId, hexCode]) => {
          const distance = ColorCalculator.calculateColorDistance(
            targetColorHex,
            hexCode,
          );
          return distance < closest.distance
            ? { id: colorId, distance }
            : closest;
        },
        { id: undefined, distance: Infinity },
      ).id;

      if (closestColorId) {
        this.hexCodeToClosestColorIdCache.set(targetColorHex, closestColorId);
      }
    }
    return this.hexCodeToClosestColorIdCache.get(targetColorHex) || "1";
  }

  private static hexToLab(hex: string): LAB {
    // Parse hex to RGB
    const number = parseInt(hex.slice(1), 16);
    const [r, g, b] = [
      (number >> 16) & 255,
      (number >> 8) & 255,
      number & 255,
    ].map((value) => {
      const normalizedChannel = value / 255;
      return normalizedChannel > 0.04045
        ? ((normalizedChannel + 0.055) / 1.055) ** 2.4
        : normalizedChannel / 12.92;
    });

    // Convert RGB → XYZ (D65)
    const [x, y, z] = [
      r * 0.4124564 + g * 0.3575761 + b * 0.1804375,
      r * 0.2126729 + g * 0.7151522 + b * 0.072175,
      r * 0.0193339 + g * 0.119192 + b * 0.9503041,
    ].map((value) => value * 100);

    // Normalize for LAB
    const [xr, yr, zr] = [x / 95.047, y / 100.0, z / 108.883].map((value) =>
      value > 0.008856 ? Math.cbrt(value) : 7.787 * value + 16 / 116,
    );

    // Return LAB
    return [116 * yr - 16, 500 * (xr - yr), 200 * (yr - zr)] as const;
  }

  private static calculateColorDistance(
    colorHex1: string,
    colorHex2: string,
  ): number {
    return Math.hypot(
      ColorCalculator.hexToLab(colorHex1)[0] -
        ColorCalculator.hexToLab(colorHex2)[0],
      ColorCalculator.hexToLab(colorHex1)[1] -
        ColorCalculator.hexToLab(colorHex2)[1],
      ColorCalculator.hexToLab(colorHex1)[2] -
        ColorCalculator.hexToLab(colorHex2)[2],
    );
  }
}
