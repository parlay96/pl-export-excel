import path from "path";
import { babel } from "@rollup/plugin-babel";
import nodeResolve from "@rollup/plugin-node-resolve";
import commonjs from "@rollup/plugin-commonjs";
import esbuild, { minify as minifyPlugin } from "rollup-plugin-esbuild";
import pkg from "./package.json" assert { type: "json" };

const PKG_BRAND_NAME = "PlExportExcel";
const banner = `/*! ${PKG_BRAND_NAME} v${pkg.version} */\n`;
function formatBundleFilename(name, minify, ext) {
  return `${name}${minify ? ".min" : ""}.${ext}`;
}

export default {
  input: "./src/index.ts",
  output: {
    file: path.resolve("dist", formatBundleFilename("index.full", true, "js")),
    format: "umd",
    exports: "named",
    name: PKG_BRAND_NAME,
    sourcemap: true,
    banner
  },
  plugins: [
    nodeResolve({
      extensions: [".js", ".ts"]
    }),
    commonjs(),
    babel({
      babelHelpers: "bundled",
      presets: ["@babel/preset-env"],
      extensions: [".js", ".ts"],
      exclude: "**/node_modules/**"
    }),
    esbuild({
      exclude: [],
      sourceMap: true,
      target: "es2018",
      define: {
        "process.env.NODE_ENV": JSON.stringify("production")
      },
      treeShaking: true,
      legalComments: "eof"
    }),
    minifyPlugin({
      target: "es2018",
      sourceMap: true
    })
  ],
  treeshake: true
};
