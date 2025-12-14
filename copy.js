// copy.js
const fs = require("fs-extra");
const path = require("path");

async function copyFiles() {
  try {
    // 确保 dist 目录存在
    await fs.ensureDir("dist");

    // 从 node_modules 中拷贝 xlsx_style.min.js 到 dist 目录
    const sourcePath = path.join("./scripts", "xlsx_style.min.js");
    const targetPath = path.join("dist", "xlsx_style.min.js");

    await fs.copy(sourcePath, targetPath);
    console.log("xlsx_style.min.js copied to dist directory");
  } catch (err) {
    console.error("Error copying file:", err);
  }
}

copyFiles();
