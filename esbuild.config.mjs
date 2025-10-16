import esbuild from "esbuild";
import { rmSync, copyFileSync } from "fs";

rmSync("dist", { recursive: true, force: true });

esbuild
  .build({
    entryPoints: ["src/main.ts"],
    bundle: true,
    platform: "neutral",
    target: "es2019",
    format: "iife",
    outdir: "dist",
    plugins: [
      {
        name: "copy-appsscript.json",
        setup(build) {
          build.onEnd(() => {
            copyFileSync("appsscript.json", "dist/appsscript.json");
          });
        },
      },
    ],
  })
  .catch(() => process.exit(1));
