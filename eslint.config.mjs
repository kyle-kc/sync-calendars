import js from "@eslint/js";
import typescript from "@typescript-eslint/eslint-plugin";
import typescriptParser from "@typescript-eslint/parser";
import prettier from "eslint-config-prettier";

export default [
  js.configs.recommended,
  prettier,
  {
    files: ["src/**/*.ts"],
    languageOptions: {
      parser: typescriptParser,
      ecmaVersion: "latest",
      sourceType: "module",
      globals: {
        Calendar: "readonly",
        GoogleAppsScript: "readonly",
        LockService: "readonly",
        Logger: "readonly",
        PropertiesService: "readonly",
        ScriptApp: "readonly",
        Utilities: "readonly",
      },
    },
    plugins: {
      "@typescript-eslint": typescript,
    },
    rules: {
      ...typescript.configs.recommended.rules,
    },
  },
  {
    ignores: ["dist/"],
  },
];
