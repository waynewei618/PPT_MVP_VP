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

const pptx = new pptxgen();
pptx.defineLayout({ name: "CUSTOM_WIDE", width: 13.333, height: 7.5 });
pptx.layout = "CUSTOM_WIDE";
pptx.author = "Codex";
pptx.company = "OpenAI";
pptx.subject = "MVP and viewport transforms";
pptx.title = "计算机图形学中的 MVP 与 Viewport 变换";
pptx.lang = "zh-CN";
pptx.theme = {
  headFontFace: "Microsoft YaHei",
  bodyFontFace: "Microsoft YaHei",
  lang: "zh-CN",
};

const tex = String.raw;
const C = {
  ink: "172033",
  muted: "596577",
  bg: "F6F8FB",
  panel: "FFFFFF",
  line: "C9D3E2",
  cyan: "0E7490",
  cyanSoft: "D7F4FA",
  green: "15803D",
  greenSoft: "DDF8E8",
  red: "B4233A",
  redSoft: "FFE2E8",
  violet: "5B43A6",
  violetSoft: "ECE8FF",
  amber: "B36B00",
  amberSoft: "FFF0CC",
  black: "111827",
  white: "FFFFFF",
};

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
const texInput = new TeX({ packages: AllPackages });
const svgOutput = new SVG({ fontCache: "none" });
const mathDocument = mathjax.document("", { InputJax: texInput, OutputJax: svgOutput });

let slide;

async function formulaSvgData(formula, color = C.ink) {
  const node = mathDocument.convert(formula, { display: true });
  const html = adaptor.outerHTML(node);
  let svg = html.match(/<svg[\s\S]*<\/svg>/)?.[0];
  if (!svg) throw new Error(`MathJax did not produce SVG for formula: ${formula}`);
  svg = svg.replace("<svg ", `<svg color="#${color}" `);
  const viewBox = svg.match(/\bviewBox="([^"]+)"/)?.[1]?.trim().split(/\s+/).map(Number);
  let aspect = 1;
  if (viewBox && viewBox.length === 4 && viewBox[2] > 0 && viewBox[3] > 0) {
    aspect = viewBox[2] / viewBox[3];
  }
  return {
    data: `data:image/svg+xml;base64,${Buffer.from(svg).toString("base64")}`,
    aspect,
  };
}

async function addFormula(formula, x, y, w, h, opts = {}) {
  const rendered = await formulaSvgData(formula, opts.color ?? C.ink);
  let drawW = w;
  let drawH = w / rendered.aspect;
  if (drawH > h) {
    drawH = h;
    drawW = h * rendered.aspect;
  }
  const align = opts.align ?? "center";
  const valign = opts.valign ?? "mid";
  const drawX = align === "left" ? x : align === "right" ? x + w - drawW : x + (w - drawW) / 2;
  const drawY = valign === "top" ? y : valign === "bottom" ? y + h - drawH : y + (h - drawH) / 2;
  slide.addImage({
    data: rendered.data,
    x: drawX,
    y: drawY,
    w: drawW,
    h: drawH,
  });
}

function addText(text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x,
    y,
    w,
    h,
    margin: opts.margin ?? 0.04,
    fontFace: opts.fontFace ?? "Microsoft YaHei",
    fontSize: opts.fontSize ?? 15,
    color: opts.color ?? C.ink,
    bold: opts.bold ?? false,
    italic: opts.italic ?? false,
    align: opts.align ?? "left",
    valign: opts.valign ?? "top",
    fit: opts.fit ?? "shrink",
    breakLine: opts.breakLine,
    rotate: opts.rotate,
    bullet: opts.bullet,
  });
}

function addBox(x, y, w, h, fill = C.panel, line = C.line, radius = 0.08) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: radius,
    fill: { color: fill },
    line: { color: line, width: 1 },
  });
}

function addRect(x, y, w, h, fill = C.panel, line = C.line) {
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w,
    h,
    fill: { color: fill },
    line: { color: line, width: 1 },
  });
}

function addLine(x, y, w, h, color = C.ink, width = 1.5, arrow = "triangle", dash) {
  slide.addShape(pptx.ShapeType.line, {
    x,
    y,
    w,
    h,
    line: { color, width, endArrowType: arrow, dash },
  });
}

function addTitle(title, kicker, accent = C.cyan) {
  addRect(0.48, 0.38, 0.06, 0.84, accent, accent);
  addText(kicker, 0.62, 0.33, 2.8, 0.25, { fontSize: 9.8, color: accent, bold: true, margin: 0 });
  addText(title, 0.62, 0.62, 8.9, 0.55, { fontSize: 27, bold: true, color: C.ink, margin: 0 });
}

function addFooter(n) {
  addText(`MVP_VP | ${n}`, 11.55, 7.08, 1.15, 0.18, {
    fontSize: 9,
    color: C.muted,
    align: "right",
    margin: 0,
  });
}

function addAxis(cx, cy, scale, color, label, rot = 0, dashed = false) {
  const a = (rot * Math.PI) / 180;
  const xdx = Math.cos(a) * scale;
  const xdy = -Math.sin(a) * scale;
  const ydx = Math.cos(a + Math.PI / 2) * scale;
  const ydy = -Math.sin(a + Math.PI / 2) * scale;
  addLine(cx, cy, xdx, xdy, color, 1.6, "triangle", dashed ? "dash" : undefined);
  addLine(cx, cy, ydx, ydy, color, 1.6, "triangle", dashed ? "dash" : undefined);
  addText(`${label}x`, cx + xdx + 0.02, cy + xdy - 0.14, 0.32, 0.18, { fontSize: 9.5, color, bold: true, margin: 0 });
  addText(`${label}y`, cx + ydx - 0.16, cy + ydy - 0.16, 0.32, 0.18, { fontSize: 9.5, color, bold: true, margin: 0 });
}

function addStepPill(i, title, x, y, w, color) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x,
    y,
    w: 0.34,
    h: 0.34,
    fill: { color },
    line: { color },
  });
  addText(String(i), x, y + 0.055, 0.34, 0.1, { fontSize: 10, color: C.white, bold: true, align: "center", margin: 0 });
  addText(title, x + 0.44, y + 0.02, w - 0.44, 0.22, { fontSize: 13, color: C.ink, bold: true, margin: 0 });
}

async function addMatrixCard(title, formula, x, y, w, h, color, sub) {
  addBox(x, y, w, h, C.panel, color);
  addText(title, x + 0.18, y + 0.16, w - 0.36, 0.23, { fontSize: 13, bold: true, color, margin: 0 });
  await addFormula(formula, x + 0.22, y + 0.54, w - 0.44, h - 1.05, { color: C.ink });
  if (sub) {
    addText(sub, x + 0.22, y + h - 0.38, w - 0.44, 0.22, { fontSize: 10.5, color: C.muted, margin: 0 });
  }
}

function addSlideBase(bg = C.bg) {
  slide = pptx.addSlide();
  slide.background = { color: bg };
}

async function slideTitle() {
  addSlideBase(C.black);
  addText("计算机图形学中的", 0.78, 0.68, 4.2, 0.42, { fontSize: 22, color: "A8DADC", bold: true, margin: 0 });
  addText("MVP 与 Viewport 变换", 0.74, 1.12, 8.5, 0.78, { fontSize: 40, color: C.white, bold: true, margin: 0 });
  addText("从“坐标系如何放置”到“像素最终落在哪里”", 0.78, 2.02, 6.5, 0.34, { fontSize: 17, color: "D4DEE9", margin: 0 });
  await addFormula(tex`P_{\mathrm{screen}} = VP \cdot P \cdot V \cdot M \cdot P_{\mathrm{local}}`, 0.85, 2.82, 5.8, 0.55, { color: "FFFFFF" });
  addBox(7.25, 0.72, 4.95, 5.85, "1F2937", "374151");
  addAxis(8.22, 5.62, 1.25, "9CA3AF", "W", 0, true);
  addAxis(9.26, 4.58, 1.04, "22D3EE", "L", 28, false);
  addAxis(10.62, 3.25, 0.92, "F59E0B", "C", 0, false);
  addRect(10.1, 1.34, 1.45, 0.88, "111827", "F43F5E");
  addText("Viewport", 10.18, 1.68, 1.28, 0.2, { fontSize: 11.5, color: "FBCFE8", bold: true, align: "center", margin: 0 });
  addLine(9.65, 4.2, 0.68, -0.72, "A8DADC", 1.4, "triangle");
  addLine(10.74, 2.9, 0.03, -0.58, "F43F5E", 1.4, "triangle");
  addText("Model", 8.6, 5.88, 0.9, 0.2, { fontSize: 10, color: "22D3EE", margin: 0 });
  addText("View", 9.75, 4.0, 0.7, 0.2, { fontSize: 10, color: "A8DADC", margin: 0 });
  addText("Projection", 10.88, 2.52, 1.0, 0.2, { fontSize: 10, color: "F43F5E", margin: 0 });
  slide.addNotes(`开场讲稿：
这份讲稿从坐标系变换讲起。MVP 和 Viewport 本质上不是一堆孤立公式，而是一条坐标表达的流水线：局部坐标先被放到世界坐标，再从相机角度观察，再投影到标准立方体，最后映射到屏幕像素。`);
}

async function slideCoordinateBasics() {
  addSlideBase();
  addTitle("坐标变换的核心：换一种坐标系描述同一个点", "01 / 坐标基础", C.cyan);
  addFooter(1);
  addBox(0.7, 1.72, 4.08, 4.9, C.panel, C.line);
  addAxis(1.62, 5.62, 1.55, C.ink, "W", 0, true);
  addAxis(2.92, 4.45, 1.35, C.cyan, "L", 0, false);
  addLine(1.7, 5.5, 1.0, -0.88, C.amber, 1.4);
  addText("同一个几何点 P", 2.9, 3.5, 1.38, 0.24, { fontSize: 12, bold: true, color: C.red, margin: 0 });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 3.18,
    y: 3.86,
    w: 0.14,
    h: 0.14,
    fill: { color: C.red },
    line: { color: C.red },
  });
  addText("坐标值会变，几何点不变", 1.06, 6.17, 3.2, 0.28, { fontSize: 13, color: C.muted, align: "center", margin: 0 });
  addBox(5.25, 1.6, 7.25, 1.15, C.cyanSoft, "99E6F2");
  await addFormula(tex`L2W =
\begin{bmatrix}
R & T\\
0^T & 1
\end{bmatrix},
\qquad
P_W = L2W \cdot P_L`, 5.48, 1.86, 5.78, 0.42, { color: C.ink });
  addText("L2W 的物理含义：把 L 坐标系放进 W 坐标系", 5.58, 2.42, 5.8, 0.22, { fontSize: 12.5, color: C.cyan, bold: true, margin: 0 });
  addBox(5.25, 3.1, 3.34, 1.42, C.panel, C.line);
  await addFormula(tex`R=[X_L\;Y_L\;Z_L]`, 5.58, 3.45, 2.12, 0.28);
  addText("R 的列向量分别是 L 的 x、y、z 轴在 W 中的方向向量。", 5.58, 3.95, 2.62, 0.36, { fontSize: 11.5, color: C.muted, margin: 0 });
  addBox(9.03, 3.1, 3.34, 1.42, C.panel, C.line);
  await addFormula(tex`T=O_L^W`, 9.42, 3.45, 1.2, 0.28);
  addText("T 是 L 的原点在 W 坐标系中的位置。", 9.42, 3.95, 2.25, 0.36, { fontSize: 11.5, color: C.muted, margin: 0 });
  addBox(5.25, 4.85, 7.25, 1.12, C.amberSoft, "F5C366");
  addText("读矩阵时先读列：方向、方向、方向、位置。齐次坐标只是把旋转和平移合成一次矩阵乘法。", 5.58, 5.24, 6.38, 0.3, { fontSize: 15, color: C.ink, bold: true, margin: 0 });
  slide.addNotes(`讲稿建议：
先建立最重要的观念：坐标变换不是移动点本身，而是用另一个坐标系表达同一个点。L2W 表示 local to world。R 的三列告诉我们 local 坐标系的三个轴在 world 中指向哪里，T 告诉我们 local 原点在 world 中放在哪里。`);
}

async function slideL2WExamples() {
  addSlideBase();
  addTitle("L2W：平移、旋转，以及“先旋转再平移”", "02 / Local 到 World", C.green);
  addFooter(2);
  await addMatrixCard("只平移：原点改变，轴方向不变", tex`L2W =
\begin{bmatrix}
1&0&2\\
0&1&3\\
0&0&1
\end{bmatrix},
\quad
P_L=\begin{bmatrix}0\\0\\1\end{bmatrix},
\quad
P_W=\begin{bmatrix}2\\3\\1\end{bmatrix}`, 0.7, 1.55, 3.9, 2.05, C.green, "T=(2,3) 表示 L 的原点位于 W 的 (2,3)。");
  await addMatrixCard("只旋转：原点不动，基向量改变", tex`L2W =
\begin{bmatrix}
\cos30^\circ&-\sin30^\circ&0\\
\sin30^\circ&\cos30^\circ&0\\
0&0&1
\end{bmatrix}`, 4.88, 1.55, 3.9, 2.05, C.red, "第一列是 L 的 x 轴方向，第二列是 L 的 y 轴方向。");
  await addMatrixCard("旋转 + 平移：把坐标系整体放到世界中", tex`L2W =
\begin{bmatrix}
1&0&2\\
0&1&3\\
0&0&1
\end{bmatrix}
\begin{bmatrix}
\cos30^\circ&-\sin30^\circ&0\\
\sin30^\circ&\cos30^\circ&0\\
0&0&1
\end{bmatrix}`, 9.05, 1.55, 3.25, 2.05, C.violet, "矩阵乘法顺序对应操作顺序。");
  addBox(0.7, 4.18, 5.15, 2.35, C.panel, C.line);
  addAxis(1.55, 6.05, 1.22, C.ink, "W", 0, true);
  addAxis(3.32, 4.82, 1.0, C.violet, "L", 30, false);
  addLine(1.64, 5.94, 1.46, -0.92, C.amber, 1.4);
  await addFormula(tex`T=(2,3)`, 2.17, 5.2, 0.85, 0.18, { color: C.amber });
  await addFormula(tex`\theta=30^\circ`, 3.82, 5.3, 0.78, 0.18, { color: C.violet });
  addBox(6.25, 4.18, 6.05, 2.35, C.greenSoft, "A7E8C2");
  addText("矩阵列的快速检查", 6.55, 4.54, 2.4, 0.24, { fontSize: 16, bold: true, color: C.green, margin: 0 });
  addText("第 1 列：局部 x 轴在世界中的方向\n第 2 列：局部 y 轴在世界中的方向\n最后 1 列：局部原点在世界中的位置", 6.62, 4.98, 4.65, 0.92, { fontSize: 14, color: C.ink, breakLine: false, margin: 0.02 });
  await addFormula(tex`P_W = L2W \cdot P_L`, 8.12, 6.08, 2.2, 0.28, { color: C.green });
  slide.addNotes(`讲稿建议：
这里对应手稿第一页的三个例子。平移时 R 是单位阵，T 直接决定原点位置。旋转时 T 为 0，R 的列向量体现新坐标轴方向。组合时可以理解为先在局部坐标系中旋转，再把整个坐标系平移到世界坐标中。`);
}

async function slideW2LInverse() {
  addSlideBase();
  addTitle("W2L：世界坐标回到局部坐标，就是 L2W 的逆", "03 / 逆变换", C.red);
  addFooter(3);
  addBox(0.82, 1.72, 4.9, 4.82, C.panel, C.line);
  addAxis(1.6, 5.8, 1.55, C.ink, "W", 0, true);
  addAxis(3.08, 4.58, 1.35, C.red, "L", 28, false);
  slide.addShape(pptx.ShapeType.ellipse, { x: 3.76, y: 3.52, w: 0.14, h: 0.14, fill: { color: C.violet }, line: { color: C.violet } });
  addText("P", 3.92, 3.42, 0.25, 0.18, { fontSize: 12, color: C.violet, bold: true, margin: 0 });
  addLine(1.72, 5.67, 1.14, -0.88, C.amber, 1.2);
  addLine(3.1, 4.48, 0.52, -0.75, C.violet, 1.2);
  addText("先减去原点，再投影到 L 的轴上", 1.25, 6.2, 3.85, 0.24, { fontSize: 13, color: C.muted, align: "center", margin: 0 });
  addBox(6.2, 1.68, 5.95, 1.25, C.redSoft, "FF9DAF");
  await addFormula(tex`W2L=(L2W)^{-1}
=
\begin{bmatrix}
R^T & -R^T T\\
0^T & 1
\end{bmatrix}`, 6.58, 1.98, 4.05, 0.52, { color: C.ink });
  addBox(6.2, 3.35, 2.72, 1.45, C.panel, C.line);
  addStepPill(1, "先平移", 6.48, 3.68, 1.7, C.amber);
  await addFormula(tex`P_W-T`, 6.92, 4.15, 0.9, 0.2, { color: C.amber });
  addText("以 L 的原点为新参考点", 6.54, 4.48, 1.84, 0.2, { fontSize: 10.5, color: C.muted, margin: 0 });
  addBox(9.42, 3.35, 2.72, 1.45, C.panel, C.line);
  addStepPill(2, "再旋转", 9.7, 3.68, 1.7, C.red);
  await addFormula(tex`R^T(P_W-T)`, 10.05, 4.15, 1.25, 0.2, { color: C.red });
  addText("投影到 L 的基向量上", 9.76, 4.48, 1.9, 0.2, { fontSize: 10.5, color: C.muted, margin: 0 });
  addBox(6.2, 5.35, 5.95, 0.88, C.violetSoft, "C6B8FF");
  await addFormula(tex`P_L = W2L \cdot P_W = V_{\mathrm{object}} \cdot P_W`, 6.68, 5.62, 4.2, 0.26, { color: C.violet });
  slide.addNotes(`讲稿建议：
手稿第二页先讨论 W2L。由于旋转矩阵是正交矩阵，所以 R 的逆等于 R 的转置。逆变换的直觉是：先把世界坐标点减掉局部坐标系原点 T，再用 R 的转置把这个向量投影到局部坐标轴上。`);
}

async function slideModelView() {
  addSlideBase();
  addTitle("M 与 V：物体放入世界，再从相机观察", "04 / Model 与 View", C.violet);
  addFooter(4);
  addBox(0.72, 1.62, 11.88, 1.05, C.violetSoft, "C6B8FF");
  await addFormula(tex`P_W = M \cdot P_L,\qquad P_C = V \cdot P_W,\qquad P_C = V \cdot M \cdot P_L`, 1.5, 1.9, 7.9, 0.32, { color: C.ink });
  addText("M 是 L2W，V 是 C2W 的逆", 9.76, 1.93, 2.12, 0.25, { fontSize: 13, color: C.violet, bold: true, margin: 0 });
  addBox(0.88, 3.05, 3.35, 2.68, C.panel, C.line);
  addAxis(1.54, 5.22, 1.22, C.green, "L", 24, false);
  addText("局部坐标 L", 1.18, 3.3, 1.3, 0.24, { fontSize: 14, bold: true, color: C.green, margin: 0 });
  await addFormula(tex`P_L=\begin{bmatrix}x\\y\\z\\1\end{bmatrix}`, 2.38, 3.75, 1.02, 0.82, { color: C.green });
  addLine(4.42, 4.25, 1.3, 0, C.violet, 1.8);
  addText("M", 4.94, 3.92, 0.36, 0.22, { fontSize: 15, bold: true, color: C.violet, align: "center", margin: 0 });
  addBox(5.92, 3.05, 3.35, 2.68, C.panel, C.line);
  addAxis(6.62, 5.22, 1.24, C.ink, "W", 0, true);
  addAxis(7.7, 4.45, 0.82, C.red, "C", 0, false);
  addText("世界坐标 W", 6.22, 3.3, 1.45, 0.24, { fontSize: 14, bold: true, color: C.ink, margin: 0 });
  addText("相机坐标 C", 7.78, 3.3, 1.2, 0.24, { fontSize: 14, bold: true, color: C.red, margin: 0 });
  addLine(9.47, 4.25, 1.3, 0, C.red, 1.8);
  addText("V", 9.98, 3.92, 0.36, 0.22, { fontSize: 15, bold: true, color: C.red, align: "center", margin: 0 });
  addBox(10.98, 3.05, 1.32, 2.68, C.redSoft, "FF9DAF");
  addText("相机空间", 11.2, 3.42, 0.9, 0.24, { fontSize: 13, bold: true, color: C.red, align: "center", margin: 0 });
  await addFormula(tex`P_C`, 11.42, 4.45, 0.48, 0.24, { color: C.red });
  addBox(0.88, 6.17, 11.42, 0.45, C.panel, C.line);
  addText("相机也有自己的 C2W；为了把世界点表示成相机坐标，View 矩阵必须取逆：", 1.15, 6.3, 5.78, 0.18, { fontSize: 11.6, color: C.muted, margin: 0 });
  await addFormula(tex`V = W2C = (C2W)^{-1}`, 6.86, 6.25, 2.14, 0.22, { color: C.red });
  slide.addNotes(`讲稿建议：
M 是模型矩阵，把模型自己的局部坐标放到世界里。相机本身也可以看成一个坐标系，它在世界里有 C2W。但渲染时我们要把世界点表达为相机坐标，所以 View 矩阵等于 C2W 的逆。最终模型到相机就是 V 乘 M。`);
}

async function slidePerspectiveIdea() {
  addSlideBase();
  addTitle("透视投影：把视锥体压到规范立方体", "05 / Projection 直觉", C.amber);
  addFooter(5);
  addBox(0.78, 1.55, 5.05, 4.95, C.panel, C.line);
  addLine(1.25, 5.45, 3.8, -2.9, C.ink, 1.1, "none");
  addLine(1.25, 5.45, 3.8, 0.15, C.ink, 1.1, "none");
  addLine(2.98, 4.1, 0, 1.42, C.line, 1.0, "none", "dash");
  addLine(4.42, 3.0, 0, 2.58, C.line, 1.0, "none", "dash");
  addText("near", 2.72, 5.68, 0.62, 0.2, { fontSize: 10, color: C.muted, margin: 0 });
  addText("far", 4.25, 5.86, 0.45, 0.2, { fontSize: 10, color: C.muted, margin: 0 });
  addText("相机", 0.98, 5.6, 0.5, 0.2, { fontSize: 10.5, color: C.ink, margin: 0 });
  slide.addShape(pptx.ShapeType.ellipse, { x: 3.7, y: 3.18, w: 0.12, h: 0.12, fill: { color: C.red }, line: { color: C.red } });
  addText("P(x,y,z)", 3.86, 3.06, 1.0, 0.2, { fontSize: 10.5, color: C.red, margin: 0 });
  await addFormula(tex`y'=\frac{n}{z}y,\qquad x'=\frac{n}{z}x`, 1.65, 2.15, 2.7, 0.34, { color: C.amber });
  addBox(6.35, 1.55, 5.72, 1.2, C.amberSoft, "F5C366");
  addText("核心比例来自相似三角形", 6.68, 1.86, 2.7, 0.24, { fontSize: 16, bold: true, color: C.amber, margin: 0 });
  addText("离相机越远，投到 near 平面上的 x、y 越小。", 6.7, 2.2, 4.58, 0.24, { fontSize: 13, color: C.ink, margin: 0 });
  addBox(6.35, 3.2, 5.72, 2.0, C.panel, C.line);
  await addFormula(tex`\begin{bmatrix}x\\y\\z\\1\end{bmatrix}
\longrightarrow
\begin{bmatrix}nx\\ny\\?\\z\end{bmatrix}
\xrightarrow{\div w}
\begin{bmatrix}\frac{n}{z}x\\\frac{n}{z}y\\?\\1\end{bmatrix}`, 6.9, 3.65, 3.75, 0.72, { color: C.ink });
  addText("手稿里的关键：先让 w'=z，再由透视除法产生 1/z 的缩放。", 6.8, 4.65, 4.58, 0.26, { fontSize: 12.2, color: C.muted, margin: 0 });
  addBox(6.35, 5.62, 5.72, 0.72, C.redSoft, "FF9DAF");
  addText("注意：不同图形 API 的相机朝向、NDC z 范围和矩阵正负号会不同；这里按手稿的 z 为深度方向讲解。", 6.72, 5.87, 4.78, 0.22, { fontSize: 11.2, color: C.red, bold: true, margin: 0 });
  slide.addNotes(`讲稿建议：
透视投影最重要的是相似三角形。设 near 平面距离为 n，空间点深度为 z，那么投影到 near 平面时，x 和 y 都乘以 n/z。矩阵本身不能直接做除法，所以先把 w 分量构造成 z，再靠透视除法完成除以 z。`);
}

async function slidePerspectiveMatrix() {
  addSlideBase();
  addTitle("透视矩阵：x、y 归一化，z 也要映射到范围内", "06 / Projection 矩阵", C.red);
  addFooter(6);
  addBox(0.72, 1.52, 3.68, 4.78, C.panel, C.line);
  addText("视锥参数", 1.0, 1.84, 1.2, 0.24, { fontSize: 16, bold: true, color: C.red, margin: 0 });
  addText("near: n\nfar: f\nleft/right: l, r\nbottom/top: b, t", 1.05, 2.38, 2.18, 1.06, { fontSize: 14, color: C.ink, margin: 0.02 });
  await addFormula(tex`\tan\frac{fovy}{2}=\frac{t}{n},
\qquad
aspect=\frac{w}{h}=\frac{r}{t}`, 1.02, 4.08, 2.35, 0.46, { color: C.red });
  addText("这些量决定 near 平面的宽高。", 1.05, 4.98, 2.72, 0.24, { fontSize: 12, color: C.muted, margin: 0 });
  addBox(4.85, 1.52, 7.05, 2.45, C.redSoft, "FF9DAF");
  await addFormula(tex`P_{\mathrm{proj}}=
\begin{bmatrix}
\frac{2n}{r-l} & 0 & \frac{r+l}{r-l} & 0\\
0 & \frac{2n}{t-b} & \frac{t+b}{t-b} & 0\\
0 & 0 & \frac{n+f}{n-f} & \frac{-2nf}{n-f}\\
0 & 0 & 1 & 0
\end{bmatrix}`, 5.48, 1.86, 4.85, 1.25, { color: C.ink });
  addText("手稿重点：第 4 行让 w'=z，后续透视除法才会出现 x/z、y/z。", 5.42, 3.45, 5.15, 0.24, { fontSize: 12, color: C.red, bold: true, margin: 0 });
  addBox(4.85, 4.45, 3.25, 1.62, C.panel, C.line);
  addText("x / y 行", 5.13, 4.75, 1.0, 0.22, { fontSize: 14, color: C.cyan, bold: true, margin: 0 });
  addText("把 near 平面的矩形区间映射到 NDC 的 [-1, 1]。", 5.13, 5.2, 2.35, 0.38, { fontSize: 12, color: C.muted, margin: 0 });
  addBox(8.58, 4.45, 3.32, 1.62, C.panel, C.line);
  addText("z 行", 8.86, 4.75, 0.65, 0.22, { fontSize: 14, color: C.violet, bold: true, margin: 0 });
  await addFormula(tex`Az+B`, 9.54, 4.73, 0.62, 0.18, { color: C.violet });
  addText("选择 A、B，让 near 和 far 映射到规定深度范围。", 8.86, 5.2, 2.48, 0.38, { fontSize: 12, color: C.muted, margin: 0 });
  slide.addNotes(`讲稿建议：
透视矩阵可以分成两块讲。前两行负责把视锥的左右上下范围归一化。第四行负责制造 w 等于 z。第三行负责处理深度，因为深度也要被映射到规范范围中，才能用于裁剪和深度测试。`);
}

async function slideViewport() {
  addSlideBase();
  addTitle("Viewport：从规范坐标到屏幕像素", "07 / Viewport 变换", C.cyan);
  addFooter(7);
  addBox(0.82, 1.52, 5.22, 4.95, C.panel, C.line);
  addRect(1.52, 2.15, 2.1, 2.1, "F9FAFB", C.ink);
  addLine(2.57, 4.25, 0, -2.1, C.line, 1, "none");
  addLine(1.52, 3.2, 2.1, 0, C.line, 1, "none");
  addText("NDC", 2.2, 1.8, 0.68, 0.22, { fontSize: 13, color: C.ink, bold: true, align: "center", margin: 0 });
  addText("(-1,1)", 1.1, 2.0, 0.72, 0.18, { fontSize: 9.5, color: C.muted, margin: 0 });
  addText("(1,-1)", 3.35, 4.26, 0.72, 0.18, { fontSize: 9.5, color: C.muted, margin: 0 });
  addLine(3.92, 3.2, 0.95, 0, C.cyan, 1.8);
  addRect(4.92, 2.02, 0.65, 2.35, C.cyanSoft, C.cyan);
  addText("屏幕", 4.98, 3.08, 0.52, 0.2, { fontSize: 11, color: C.cyan, bold: true, align: "center", margin: 0 });
  await addFormula(tex`[-1,1]^3 \rightarrow [0,W]\times[0,H]\times[0,1]`, 1.26, 5.15, 3.68, 0.28, { color: C.cyan });
  addBox(6.55, 1.52, 5.5, 1.62, C.cyanSoft, "99E6F2");
  await addFormula(tex`M_{VP}=
\begin{bmatrix}
\frac{W}{2}&0&0&\frac{W}{2}\\
0&\frac{H}{2}&0&\frac{H}{2}\\
0&0&1&0\\
0&0&0&1
\end{bmatrix}`, 7.15, 1.82, 3.4, 0.82, { color: C.ink });
  addBox(6.55, 3.62, 5.5, 1.08, C.panel, C.line);
  await addFormula(tex`\begin{aligned}
x_{\mathrm{screen}}&=\frac{W}{2}x_{\mathrm{ndc}}+\frac{W}{2}\\
y_{\mathrm{screen}}&=\frac{H}{2}y_{\mathrm{ndc}}+\frac{H}{2}
\end{aligned}`, 7.0, 3.88, 3.42, 0.44, { color: C.ink });
  addBox(6.55, 5.12, 5.5, 0.82, C.amberSoft, "F5C366");
  addText("如果屏幕坐标 y 轴向下，实际工程里会把 y 的符号或偏移作相应调整。", 6.92, 5.41, 4.72, 0.22, { fontSize: 12.2, color: C.amber, bold: true, margin: 0 });
  slide.addNotes(`讲稿建议：
投影之后得到的是 NDC，范围通常是 -1 到 1。Viewport 变换做的是尺度和平移：把 -1 到 1 的区间缩放到屏幕宽高，再平移到屏幕坐标原点附近。不同系统的屏幕 y 轴方向可能不同，所以实际矩阵可能在 y 上有负号。`);
}

async function slidePipeline() {
  addSlideBase();
  addTitle("完整流水线：每一步都在回答一个坐标系问题", "08 / MVP + VP 总览", C.violet);
  addFooter(8);
  const y = 2.2;
  const xs = [0.65, 2.93, 5.22, 7.52, 9.82];
  const names = ["Local", "World", "Camera", "NDC", "Screen"];
  const colors = [C.green, C.violet, C.red, C.amber, C.cyan];
  for (let i = 0; i < xs.length; i++) {
    addBox(xs[i], y, 1.58, 1.08, i === 4 ? C.cyanSoft : C.panel, colors[i]);
    addText(names[i], xs[i] + 0.2, y + 0.34, 1.18, 0.23, { fontSize: 14, bold: true, color: colors[i], align: "center", margin: 0 });
    if (i < xs.length - 1) addLine(xs[i] + 1.72, y + 0.54, 0.68, 0, colors[i + 1], 1.8);
  }
  addText("M", 2.35, y + 0.18, 0.36, 0.2, { fontSize: 15, bold: true, color: C.violet, align: "center", margin: 0 });
  addText("V", 4.62, y + 0.18, 0.36, 0.2, { fontSize: 15, bold: true, color: C.red, align: "center", margin: 0 });
  addText("P", 6.9, y + 0.18, 0.36, 0.2, { fontSize: 15, bold: true, color: C.amber, align: "center", margin: 0 });
  addText("VP", 9.0, y + 0.18, 0.42, 0.2, { fontSize: 15, bold: true, color: C.cyan, align: "center", margin: 0 });
  addBox(1.18, 4.05, 10.85, 1.02, C.violetSoft, "C6B8FF");
  await addFormula(tex`P_{\mathrm{screen}}
= M_{VP}\cdot P_{\mathrm{proj}}\cdot V\cdot M\cdot P_{\mathrm{local}}`, 2.35, 4.38, 6.52, 0.34, { color: C.ink });
  addBox(1.18, 5.65, 2.36, 0.68, C.greenSoft, "A7E8C2");
  addText("M：物体如何放进世界", 1.45, 5.9, 1.82, 0.18, { fontSize: 11.8, color: C.green, bold: true, align: "center", margin: 0 });
  addBox(3.86, 5.65, 2.36, 0.68, C.redSoft, "FF9DAF");
  addText("V：世界如何被相机看见", 4.08, 5.9, 1.92, 0.18, { fontSize: 11.8, color: C.red, bold: true, align: "center", margin: 0 });
  addBox(6.55, 5.65, 2.36, 0.68, C.amberSoft, "F5C366");
  addText("P：视锥如何压成盒子", 6.82, 5.9, 1.82, 0.18, { fontSize: 11.8, color: C.amber, bold: true, align: "center", margin: 0 });
  addBox(9.23, 5.65, 2.36, 0.68, C.cyanSoft, "99E6F2");
  addText("VP：盒子如何落到屏幕", 9.46, 5.9, 1.9, 0.18, { fontSize: 11.8, color: C.cyan, bold: true, align: "center", margin: 0 });
  slide.addNotes(`讲稿建议：
最后把所有内容串起来。M、V、P、VP 都是在改变坐标表达。M 从局部到世界，V 从世界到相机，P 从相机到裁剪空间和 NDC，VP 从 NDC 到屏幕。只要每一步都问清楚“当前坐标系是什么、下一步坐标系是什么”，公式就不会混乱。`);
}

async function slideSummary() {
  addSlideBase(C.black);
  addText("总结", 0.78, 0.68, 1.5, 0.36, { fontSize: 20, color: "A8DADC", bold: true, margin: 0 });
  addText("把矩阵看成坐标系说明书", 0.78, 1.12, 7.4, 0.62, { fontSize: 34, color: C.white, bold: true, margin: 0 });
  const items = [
    ["R", "描述新坐标轴在旧坐标系中的方向"],
    ["T", "描述新原点在旧坐标系中的位置"],
    ["V", "相机坐标系 C2W 的逆变换"],
    ["P", "用 w'=z 和透视除法制造近大远小"],
    ["VP", "把 NDC 的 [-1,1] 映射到屏幕像素"],
  ];
  for (let i = 0; i < items.length; i++) {
    const yy = 2.2 + i * 0.72;
    slide.addShape(pptx.ShapeType.ellipse, { x: 0.92, y: yy, w: 0.42, h: 0.42, fill: { color: "22D3EE" }, line: { color: "22D3EE" } });
    addText(items[i][0], 0.92, yy + 0.105, 0.42, 0.12, { fontSize: 10.5, color: C.black, bold: true, align: "center", margin: 0 });
    addText(items[i][1], 1.58, yy + 0.08, 5.7, 0.23, { fontSize: 16, color: C.white, margin: 0 });
  }
  addBox(7.65, 2.04, 4.22, 2.68, "1F2937", "374151");
  await addFormula(tex`\boxed{P_{\mathrm{screen}} = VP\cdot P\cdot V\cdot M\cdot P_L}`, 8.08, 2.52, 3.3, 0.42, { color: C.white });
  addText("记忆方式", 8.1, 3.42, 0.9, 0.22, { fontSize: 13, color: "A8DADC", bold: true, margin: 0 });
  addText("局部 -> 世界 -> 相机 -> 投影 -> 屏幕", 8.1, 3.86, 2.72, 0.26, { fontSize: 14, color: "D4DEE9", margin: 0 });
  addText("坐标系清楚，矩阵就清楚。", 8.1, 4.24, 2.18, 0.24, { fontSize: 13, color: "FDE68A", bold: true, margin: 0 });
  slide.addNotes(`收尾讲稿：
总结时强调：MVP 和 Viewport 的难点不在矩阵符号，而在坐标系关系。看到矩阵先问 R 和 T 在描述哪个坐标系，再问点从哪个坐标表达变到哪个坐标表达。这样就能把模型、相机、投影和屏幕映射统一起来。`);
}

async function main() {
  await slideTitle();
  await slideCoordinateBasics();
  await slideL2WExamples();
  await slideW2LInverse();
  await slideModelView();
  await slidePerspectiveIdea();
  await slidePerspectiveMatrix();
  await slideViewport();
  await slidePipeline();
  await slideSummary();
  await pptx.writeFile({ fileName: "/home/sil/workspace/PPT_MVP_VP/MVP_VP_handwritten.pptx" });
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
