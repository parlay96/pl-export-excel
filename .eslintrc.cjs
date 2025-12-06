module.exports = {
  root: true,
  env: {
    node: true
  },
  extends: ["eslint:recommended", "plugin:@typescript-eslint/recommended"],
  parserOptions: {
    ecmaVersion: 2020
  },
  ignorePatterns: ["node_modules/", "dist/"],
  parser: "@typescript-eslint/parser",
  plugins: [],
  rules: {
    // JavaScript/ESLint 推荐的规则
    "no-console": "off", // 不允许使用 console.log 等
    "no-unused-vars": "warn", // 不允许存在未使用的变量
    "no-undef": "off", // 不允许使用未定义的变量

    "no-mixed-spaces-and-tabs": "off",
    "@typescript-eslint/ban-ts-comment": "off",
    quotes: "off",
    "no-unsafe-optional-chaining": "off",
    "prefer-const": "off",
    // TypeScript/ESLint 推荐的规则
    "@typescript-eslint/ban-types": "off",
    "@typescript-eslint/no-unused-vars": "warn", // 不允许存在未使用的 TypeScript 变量
    "@typescript-eslint/explicit-module-boundary-types": "off", // 允许不显式指定导出函数的返回类型
    "@typescript-eslint/no-explicit-any": "off", // 允许使用 any 类型
    "@typescript-eslint/no-explicit-any": ["off"]
  }
};
