import { describe, expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

const pkg = JSON.parse(readFileSync(join(import.meta.dir, "../package.json"), "utf8"));

describe("zero runtime dependencies", () => {
  test("package.json declares no dependencies or peerDependencies", () => {
    expect(pkg.dependencies ?? {}).toEqual({});
    expect(pkg.peerDependencies ?? {}).toEqual({});
  });
  test("no src/ file imports a non-relative module", () => {
    const dir = join(import.meta.dir, "../src");
    for (const f of readdirSync(dir)) {
      if (!f.endsWith(".ts")) continue;
      const src = readFileSync(join(dir, f), "utf8");
      const imports = [...src.matchAll(/from\s+"([^"]+)"/g)].map((m) => m[1]);
      for (const spec of imports) {
        expect(spec!.startsWith("./") || spec!.startsWith("../")).toBe(true);
      }
    }
  });
});
