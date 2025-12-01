import glob from "fast-glob";
import { babel } from "@rollup/plugin-babel";
import typescript from "@rollup/plugin-typescript";
import nodeResolve from "@rollup/plugin-node-resolve";
import commonjs from "@rollup/plugin-commonjs";
import esbuild from "rollup-plugin-esbuild";
import { terser } from "rollup-plugin-terser";
import pkg from "./package.json" assert { type: "json" };

const files = glob.sync("**/*.{ts,tsx}", {
  cwd: "./src",
  absolute: true,
  onlyFiles: true
});

const externalPackages = [...Object.keys(pkg.dependencies || {}), ...Object.keys(pkg.peerDependencies || {})];

export default {
  input: (() => {
    const excludes = ["node_modules", "test", "dist", "style", "types"];
    return files.filter((path) => !excludes.some((exclude) => path.includes(exclude)));
  })(),
  output: {
    dir: "dist",
    format: "esm",
    sourcemap: true,
    preserveModules: true,
    preserveModulesRoot: "src"
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
      // map compressed or converted code back to the original source code
      sourceMap: true,
      target: "es2018"
    }),
    typescript(),
    terser({
      mangle: false, // Disable variable name obfuscation
      output: {
        beautify: true, // Beautify output
        comments: false // Delete all comments
      }
    })
  ],
  // Creating regular expressions for the package to ensure that the subpath is also treated as external
  external: [...externalPackages.map((packageName) => new RegExp(`^${packageName}(/.*)?`)), "xlsx-style"],
  treeshake: false
};
