const pptxgen = require("pptxgenjs");
const path = require("path");
const { execSync } = require("child_process");

function requireLocalOrGlobal(modulePath) {
  try {
    return require(modulePath);
  } catch (err) {
    const globalRoot = execSync("npm root -g", { encoding: "utf8" }).trim();
    return require(path.join(globalRoot, modulePath));
  }
}

const { mathjax } = requireLocalOrGlobal("mathjax-full/js/mathjax.js");
const { TeX } = requireLocalOrGlobal("mathjax-full/js/input/tex.js");
const { SVG } = requireLocalOrGlobal("mathjax-full/js/output/svg.js");
const { liteAdaptor } = requireLocalOrGlobal("mathjax-full/js/adaptors/liteAdaptor.js");
const { RegisterHTMLHandler } = requireLocalOrGlobal("mathjax-full/js/handlers/html.js");
const { AllPackages } = requireLocalOrGlobal("mathjax-full/js/input/tex/AllPackages.js");
const sharp = requireLocalOrGlobal("sharp");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Codex";
pptx.company = "OpenAI";
pptx.subject = "MVP_VP coordinate transforms";
pptx.title = "二维平移与旋转：变换矩阵的物理意义";
pptx.lang = "zh-CN";
pptx.theme = {
  headFontFace: "Microsoft YaHei",
  bodyFontFace: "Microsoft YaHei",
  lang: "zh-CN",
};
pptx.defineLayout({ name: "CUSTOM_WIDE", width: 13.333, height: 7.5 });
pptx.layout = "CUSTOM_WIDE";

let slide;

const C = {
  ink: "111827",
  muted: "64748B",
  grid: "CBD5E1",
  panel: "FFFFFF",
  teal: "0F766E",
  tealSoft: "CCFBF1",
  rose: "BE123C",
  roseSoft: "FFE4E6",
  amber: "B45309",
  amberSoft: "FEF3C7",
  blue: "2563EB",
};

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
const texInput = new TeX({ packages: AllPackages });
const svgOutput = new SVG({ fontCache: "none" });
const mathDocument = mathjax.document("", { InputJax: texInput, OutputJax: svgOutput });

function addText(text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    margin: opts.margin ?? 0.04,
    fontFace: opts.fontFace ?? "Microsoft YaHei",
    fontSize: opts.fontSize ?? 16,
    color: opts.color ?? C.ink,
    bold: opts.bold ?? false,
    italic: opts.italic ?? false,
    align: opts.align ?? "left",
    valign: opts.valign ?? "top",
    breakLine: opts.breakLine,
    fit: opts.fit,
    rotate: opts.rotate,
  });
}

function addMono(text, x, y, w, h, opts = {}) {
  addText(text, x, y, w, h, {
    fontFace: "Consolas",
    fontSize: opts.fontSize ?? 14,
    color: opts.color ?? C.ink,
    bold: opts.bold ?? false,
    align: opts.align ?? "center",
    valign: "mid",
    margin: opts.margin ?? 0.02,
    fit: "shrink",
  });
}

// Global LaTeX formula input environment.
// TeX source is rendered by MathJax to SVG, then inserted into PPTX as an image.
const tex = String.raw;

async function formulaPngData(formula, color = C.ink) {
  const node = mathDocument.convert(formula, { display: true });
  const html = adaptor.outerHTML(node);
  let svg = html.match(/<svg[\s\S]*<\/svg>/)?.[0];
  if (!svg) {
    throw new Error(`MathJax did not produce SVG for formula: ${formula}`);
  }
  svg = svg.replace("<svg ", `<svg color="#${color}" `);
  const png = await sharp(Buffer.from(svg), { density: 600 }).png().toBuffer();
  return `data:image/png;base64,${png.toString("base64")}`;
}

async function addFormula(formula, x, y, w, h, opts = {}) {
  slide.addImage({
    data: await formulaPngData(formula, opts.color ?? C.ink),
    x, y, w, h,
  });
}

function addBox(x, y, w, h, fill, line = "E2E8F0") {
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: fill },
    line: { color: line, width: 1 },
  });
}

function addLine(x, y, w, h, color = C.ink, width = 1.6, endArrowType = "triangle") {
  slide.addShape(pptx.ShapeType.line, {
    x, y, w, h,
    line: { color, width, endArrowType },
  });
}

function addDot(cx, cy, color, label, lx = 0.08, ly = -0.22) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x: cx - 0.045, y: cy - 0.045, w: 0.09, h: 0.09,
    fill: { color }, line: { color },
  });
  addText(label, cx + lx, cy + ly, 0.85, 0.24, {
    fontSize: 10.5, color, bold: true, margin: 0,
  });
}

function addAxis(originX, originY, scale, rotationDeg, labelPrefix, color, dashed = false) {
  const angle = rotationDeg * Math.PI / 180;
  const xdx = Math.cos(angle) * scale;
  const xdy = -Math.sin(angle) * scale;
  const ydx = Math.cos(angle + Math.PI / 2) * scale;
  const ydy = -Math.sin(angle + Math.PI / 2) * scale;
  const line = dashed ? { color, width: 1.4, dash: "dash", endArrowType: "triangle" } : { color, width: 1.7, endArrowType: "triangle" };

  slide.addShape(pptx.ShapeType.line, {
    x: originX, y: originY, w: xdx, h: xdy,
    line,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: originX, y: originY, w: ydx, h: ydy,
    line,
  });
  addText(`${labelPrefix}x`, originX + xdx + 0.02, originY + xdy - 0.12, 0.35, 0.2, {
    fontSize: 10, color, bold: true, margin: 0,
  });
  addText(`${labelPrefix}y`, originX + ydx - 0.16, originY + ydy - 0.18, 0.35, 0.2, {
    fontSize: 10, color, bold: true, margin: 0,
  });
}

function addMiniGrid(x, y, w, h) {
  for (let i = 1; i < 4; i++) {
    addLine(x + (w * i) / 4, y, 0, h, C.grid, 0.45, "none");
    addLine(x, y + (h * i) / 4, w, 0, C.grid, 0.45, "none");
  }
}

async function main() {
slide = pptx.addSlide();
slide.background = { color: "F8FAFC" };

// Header
addText("二维平移与旋转：变换矩阵的物理意义", 0.55, 0.34, 7.6, 0.46, {
  fontSize: 25, bold: true, color: C.ink, margin: 0,
});
addText("核心问题：矩阵不是“算式黑盒”，它描述了坐标系 A 在坐标系 B 中的位置与方向。", 0.58, 0.82, 8.2, 0.28, {
  fontSize: 11.5, color: C.muted, margin: 0,
});
await addFormula(tex`P_A \rightarrow P_B`, 10.9, 0.32, 1.82, 0.36, { color: C.teal });
addText("A 坐标表示   B 坐标表示", 10.68, 0.72, 2.18, 0.22, {
  fontSize: 8.8, color: C.muted, align: "center", margin: 0,
});

// Unified matrix band
addBox(0.55, 1.16, 12.22, 1.13, "FFFFFF", "CBD5E1");
addText("统一写法", 0.83, 1.35, 0.9, 0.28, {
  fontSize: 14.5, bold: true, color: C.teal, margin: 0,
});
await addFormula(tex`\mathrm{A2B} =
\begin{bmatrix}
R & T \\
0 & 1
\end{bmatrix}`, 1.75, 1.25, 2.15, 0.68, {
  fontSize: 17,
});
await addFormula(tex`P_{B} = \mathrm{A2B} \cdot P_{A}`, 4.02, 1.42, 2.1, 0.34, {
  fontSize: 16, color: C.rose, bold: true,
});
await addFormula(tex`R`, 6.12, 1.34, 0.12, 0.12, { color: C.ink });
addText("：坐标系 A 的坐标轴在坐标系 B 中的方向", 6.26, 1.31, 3.1, 0.27, {
  fontSize: 11.8, bold: true, color: C.ink, margin: 0, fit: "shrink",
});
await addFormula(tex`T`, 6.12, 1.71, 0.12, 0.12, { color: C.ink });
addText("：坐标系 A 的原点", 6.26, 1.68, 1.18, 0.27, {
  fontSize: 11.8, bold: true, color: C.ink, margin: 0, fit: "shrink",
});
await addFormula(tex`O_A`, 7.45, 1.71, 0.22, 0.12, { color: C.ink });
addText("在坐标系 B 中的位置", 7.72, 1.68, 1.55, 0.27, {
  fontSize: 11.8, bold: true, color: C.ink, margin: 0, fit: "shrink",
});
addText("矩阵的列向量在回答两个问题：\n1. 新坐标轴指向哪里？\n2. 新原点被放在哪里？", 9.55, 1.26, 2.75, 0.72, {
  fontSize: 11.8, color: C.muted, margin: 0.02, fit: "shrink",
});

// Panels
addBox(0.55, 2.55, 5.98, 4.42, "FFFFFF", "CBD5E1");
addBox(6.8, 2.55, 5.98, 4.42, "FFFFFF", "CBD5E1");
addBox(0.55, 2.55, 5.98, 0.46, C.tealSoft, "99F6E4");
addBox(6.8, 2.55, 5.98, 0.46, C.roseSoft, "FDA4AF");
addText("例 1：二维平移", 0.82, 2.66, 1.6, 0.24, {
  fontSize: 14, bold: true, color: C.teal, margin: 0,
});
await addFormula(tex`T=(t_x,t_y)`, 2.48, 2.66, 0.95, 0.22, { color: C.teal });
addText("例 2：二维旋转", 7.07, 2.66, 1.6, 0.24, {
  fontSize: 14, bold: true, color: C.rose, margin: 0,
});
await addFormula(tex`\theta=45^\circ`, 8.68, 2.66, 0.8, 0.22, { color: C.rose });

// Translation content
await addFormula(tex`\begin{bmatrix}
1 & 0 & t_x \\
0 & 1 & t_y \\
0 & 0 & 1
\end{bmatrix}`, 0.9, 3.27, 1.7, 1.0, {
  fontSize: 16,
});
addText("物理意义", 2.86, 3.17, 1.0, 0.23, {
  fontSize: 12.5, bold: true, color: C.teal, margin: 0,
});
await addFormula(tex`R=I`, 2.86, 3.47, 0.34, 0.13, { color: C.ink });
addText("：A 的坐标轴方向与 B 相同", 3.23, 3.46, 2.4, 0.2, {
  fontSize: 10.8, color: C.ink, margin: 0, fit: "shrink",
});
await addFormula(tex`T=(t_x,t_y)`, 2.86, 3.74, 0.58, 0.13, { color: C.ink });
addText("：把 A 的原点", 3.46, 3.73, 0.88, 0.2, {
  fontSize: 10.5, color: C.ink, margin: 0, fit: "shrink",
});
await addFormula(tex`O_A`, 4.32, 3.74, 0.18, 0.13, { color: C.ink });
addText("放到 B 坐标", 4.52, 3.73, 0.74, 0.2, {
  fontSize: 10.5, color: C.ink, margin: 0, fit: "shrink",
});
await addFormula(tex`(t_x,t_y)`, 5.24, 3.74, 0.38, 0.13, { color: C.ink });
addText("任意点都加同一个偏移量：", 2.86, 4.0, 1.5, 0.2, {
  fontSize: 10.2, color: C.ink, margin: 0, fit: "shrink",
});
await addFormula(tex`x_B=x_A+t_x,\quad y_B=y_A+t_y`, 4.25, 3.99, 1.7, 0.18, { color: C.ink });
addBox(0.9, 4.62, 2.35, 1.72, "F8FAFC", "E2E8F0");
addMiniGrid(1.05, 4.78, 2.05, 1.37);
const txOriginX = 1.28;
const txOriginY = 5.93;
addAxis(txOriginX, txOriginY, 1.15, 0, "B", C.ink);
addDot(txOriginX, txOriginY, C.ink, "", -0.34, 0.03);
await addFormula(tex`O_B`, txOriginX - 0.34, txOriginY + 0.04, 0.22, 0.13, { color: C.ink });
const ocX = 2.52;
const ocY = 5.06;
addAxis(ocX, ocY, 0.78, 0, "A", C.teal);
addDot(ocX, ocY, C.teal, "", 0.06, -0.24);
await addFormula(tex`O_A=(t_x,t_y)`, ocX + 0.03, ocY - 0.22, 0.72, 0.16, { color: C.teal });
addLine(txOriginX + 0.07, txOriginY - 0.07, ocX - txOriginX - 0.18, ocY - txOriginY + 0.18, C.amber, 1.3, "triangle");
await addFormula(tex`T`, 1.83, 5.36, 0.18, 0.15, { color: C.amber });
addBox(3.58, 4.62, 2.45, 1.72, C.amberSoft, "FCD34D");
addText("坐标转化的例子", 3.78, 4.81, 1.25, 0.22, { fontSize: 12, bold: true, color: C.amber, margin: 0 });
await addFormula(tex`\begin{gathered}
P_A=(2,1),\quad T=(3,2)\\[-2pt]
\begin{bmatrix}
x_B\\
y_B\\
1
\end{bmatrix}
=
\begin{bmatrix}
1&0&3\\
0&1&2\\
0&0&1
\end{bmatrix}
\begin{bmatrix}
2\\
1\\
1
\end{bmatrix}\\[-2pt]
\begin{aligned}
x_B&=1\cdot2+0\cdot1+3\cdot1=5\\
y_B&=0\cdot2+1\cdot1+2\cdot1=3\\
P_B&=(5,3)
\end{aligned}
\end{gathered}`, 3.7, 4.96, 2.05, 1.08, { color: C.ink });

// Rotation content
await addFormula(tex`\begin{bmatrix}
\cos\theta & -\sin\theta & 0 \\
\sin\theta & \cos\theta & 0 \\
0 & 0 & 1
\end{bmatrix}`, 7.05, 3.22, 2.34, 1.1, {
  fontSize: 13.6,
});
addText("列向量就是旋转后的基底", 9.6, 3.17, 2.05, 0.23, {
  fontSize: 12.5, bold: true, color: C.rose, margin: 0,
});
addText("第 1 列", 9.6, 3.45, 0.5, 0.18, { fontSize: 10.4, color: C.ink, margin: 0 });
await addFormula(tex`(\cos\theta,\sin\theta)`, 10.08, 3.45, 0.68, 0.15, { color: C.ink });
addText("：A 的 x 轴在 B 中的方向", 11.0, 3.45, 1.2, 0.3, { fontSize: 10.3, color: C.ink, margin: 0, fit: "shrink" });
addText("第 2 列", 9.6, 3.84, 0.5, 0.18, { fontSize: 10.4, color: C.ink, margin: 0 });
await addFormula(tex`(-\sin\theta,\cos\theta)`, 10.08, 3.84, 0.68, 0.15, { color: C.ink });
addText("：A 的 y 轴在 B 中的方向", 11.0, 3.84, 1.2, 0.3, { fontSize: 10.3, color: C.ink, margin: 0, fit: "shrink" });
await addFormula(tex`T=0`, 9.6, 4.25, 0.32, 0.13, { color: C.ink });
addText("：只改变方向，不移动原点", 10.05, 4.24, 1.55, 0.2, { fontSize: 10.4, color: C.ink, margin: 0, fit: "shrink" });
addBox(7.15, 4.62, 2.35, 1.72, "F8FAFC", "E2E8F0");
addMiniGrid(7.31, 4.78, 2.03, 1.37);
const roX = 8.05;
const roY = 5.86;
addAxis(roX, roY, 1.02, 0, "B", C.ink, true);
addAxis(roX, roY, 1.02, 45, "A", C.rose);
addDot(roX, roY, C.ink, "O", -0.22, 0.04);
slide.addShape(pptx.ShapeType.arc, {
  x: roX + 0.28, y: roY - 0.58, w: 0.62, h: 0.62,
  adjustPoint: 0.25,
  line: { color: C.amber, width: 1.1 },
  fill: { color: "FFFFFF", transparency: 100 },
});
await addFormula(tex`\theta`, roX + 0.73, roY - 0.54, 0.16, 0.14, { color: C.amber });
addBox(9.82, 4.62, 2.45, 1.72, C.tealSoft, "5EEAD4");
addText("坐标转化的例子", 10.02, 4.81, 1.25, 0.22, { fontSize: 12, bold: true, color: C.teal, margin: 0 });
await addFormula(tex`\begin{gathered}
P_A=(1,0),\quad \theta=45^\circ\\[-2pt]
\begin{bmatrix}
x_B\\
y_B\\
1
\end{bmatrix}
=
\begin{bmatrix}
\cos\theta&-\sin\theta&0\\
\sin\theta&\cos\theta&0\\
0&0&1
\end{bmatrix}
\begin{bmatrix}
1\\
0\\
1
\end{bmatrix}\\[-2pt]
\begin{aligned}
x_B&=\cos\theta\cdot1-\sin\theta\cdot0=\cos45^\circ\\
y_B&=\sin\theta\cdot1+\cos\theta\cdot0=\sin45^\circ\\
P_B&=(\cos45^\circ,\sin45^\circ)
\end{aligned}
\end{gathered}`, 9.92, 4.96, 2.0, 1.08, { color: C.ink });

// Footer takeaway
addBox(0.55, 7.09, 12.22, 0.25, C.ink, C.ink);
addText("一句话总结：A2B 的 R 负责“朝向”，T 负责“位置”；齐次坐标把二者统一到一个矩阵乘法里。", 0.72, 7.13, 10.7, 0.17, {
  fontSize: 9.7, color: "FFFFFF", bold: true, margin: 0,
});
addText("MVP_VP", 11.96, 7.13, 0.62, 0.17, {
  fontSize: 9.5, color: "FFFFFF", align: "right", margin: 0,
});

slide.addNotes(`讲稿建议：
1. 先从统一公式讲起：P_A 是同一个点在坐标系 A 中的坐标，P_B 是它在坐标系 B 中的坐标。A2B 描述“坐标系 A 如何嵌入坐标系 B”。
2. 平移例子：R 等于单位矩阵，所以 A 的 x/y 轴方向与 B 相同；T 等于 (tx, ty)，表示 A 的原点 O_A 在 B 坐标系中的位置。点 (0,0,1) 经过矩阵后变成 (tx,ty,1)，这正好说明第三列的物理意义。
3. 旋转例子：T 为 0，所以原点不动；R 的第一列是旋转后的 A-x 轴方向，第二列是旋转后的 A-y 轴方向。点 P_A=(x_A,y_A) 实际上是在说“沿 A-x 走 x_A，再沿 A-y 走 y_A”，矩阵乘法就是把这两个方向向量线性组合，得到 B 坐标系中的 P_B。
4. 收束到 MVP/VP：三维里的模型变换和 View 变换也是同一件事，只是 R/T 从二维扩展到三维；Viewport 变换则是把投影后的规范坐标再放到屏幕窗口。`);

await pptx.writeFile({ fileName: "/home/wwyhkq/workspace/PPT_MVP_VP/MVP_VP_2D_transform_meaning.pptx" });
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
