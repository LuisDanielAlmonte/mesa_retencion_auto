/* Exportación a PowerPoint editable.
   Convierte un slide HTML (lienzo de 1920×1080 px) en objetos nativos de PowerPoint con PptxGenJS:
   - cada caja con fondo o borde → forma (rectángulo, rectángulo redondeado o elipse), con degradado si lo tiene
   - cada texto → cuadro de texto editable con su fuente, tamaño, color y peso
   - cada gráfico SVG → formas vectoriales (líneas, curvas, barras, puntos y etiquetas), todas editables
   Uso: xpPptx.slide(pptx, pptSlide, slideEl) por cada slide y al final await xpPptx.save(pptx, nombre).
   El slide debe estar visible y sin escalar (transform: none) mientras se convierte. */
(function () {
  const PX = 1 / 144;            // 1920 px → 13.333 in
  const PT = 0.5;                // 1 px del lienzo → 0.5 pt
  const r2 = v => Math.round(v * 1000) / 1000;
  let gradSeq = 0;
  const grads = {};              // nombre de forma → { fill?: xml, line?: xml }

  /* ---------- Colores ---------- */
  function parseColor(s) {
    if (!s) return null;
    s = s.trim();
    if (s === "transparent" || s === "none") return null;
    let m = s.match(/^rgba?\(([^)]+)\)$/);
    if (m) {
      const p = m[1].split(/[\s,\/]+/).filter(Boolean).map(parseFloat);
      return { r: p[0], g: p[1], b: p[2], a: p.length > 3 ? p[3] : 1 };
    }
    m = s.match(/^#([0-9a-f]{3,8})$/i);
    if (m) {
      let h = m[1];
      if (h.length <= 4) h = h.split("").map(c => c + c).join("");
      return { r: parseInt(h.slice(0, 2), 16), g: parseInt(h.slice(2, 4), 16), b: parseInt(h.slice(4, 6), 16), a: h.length === 8 ? parseInt(h.slice(6, 8), 16) / 255 : 1 };
    }
    // nombres u otros formatos: que el navegador lo resuelva
    const t = document.createElement("i");
    t.style.color = s; document.body.appendChild(t);
    const c = getComputedStyle(t).color; t.remove();
    return c && c !== s ? parseColor(c) : null;
  }
  const hex = c => [c.r, c.g, c.b].map(v => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, "0")).join("").toUpperCase();
  const blend = (c, bg) => ({ r: c.r * c.a + bg.r * (1 - c.a), g: c.g * c.a + bg.g * (1 - c.a), b: c.b * c.a + bg.b * (1 - c.a), a: 1 });
  const transp = a => Math.round((1 - Math.max(0, Math.min(1, a))) * 100);

  // Divide por comas de primer nivel (fuera de paréntesis)
  function splitTop(s) {
    const out = []; let depth = 0, cur = "";
    for (const ch of s) {
      if (ch === "(") depth++;
      if (ch === ")") depth--;
      if (ch === "," && !depth) { out.push(cur.trim()); cur = ""; } else cur += ch;
    }
    if (cur.trim()) out.push(cur.trim());
    return out;
  }
  // linear-gradient(...) → { angle (grados CSS), stops: [{c, pos 0..1}] }
  // Fondo con varias capas: se usa la capa lineal de base; los brillos radiales semitransparentes se omiten
  function parseBackground(img) {
    if (!img || img === "none") return null;
    const layers = splitTop(img).map(parseGradient).filter(Boolean);
    const lin = layers.filter(g => !g.radial);
    if (lin.length) return lin[lin.length - 1];
    const rad = layers.find(g => g.stops.every(s => s.c.a > .5));
    return rad || null;
  }
  function parseGradient(img) {
    const m = img.match(/(?:repeating-)?(linear|radial)-gradient\((.*)\)$/s);
    if (!m) return null;
    const parts = splitTop(m[2]);
    let angle = 180;
    if (m[1] === "radial") { if (!/rgb|#/.test(parts[0])) parts.shift(); }
    else if (/deg|turn|rad/.test(parts[0])) {
      const p = parts.shift();
      angle = p.endsWith("turn") ? parseFloat(p) * 360 : p.endsWith("rad") ? parseFloat(p) * 180 / Math.PI : parseFloat(p);
    } else if (/^to /.test(parts[0])) {
      const p = parts.shift();
      const dirs = { "to top": 0, "to right": 90, "to bottom": 180, "to left": 270, "to top right": 45, "to right top": 45, "to bottom right": 135, "to right bottom": 135, "to bottom left": 225, "to left bottom": 225, "to top left": 315, "to left top": 315 };
      angle = dirs[p] ?? 180;
    }
    const stops = parts.map(p => {
      const cm = p.match(/^(rgba?\([^)]*\)|#[0-9a-f]+|[a-z]+)\s*(.*)$/i);
      if (!cm) return null;
      const c = parseColor(cm[1]);
      const pos = cm[2] ? parseFloat(cm[2]) / 100 : null;
      return c ? { c, pos } : null;
    }).filter(Boolean);
    if (!stops.length) return null;
    stops.forEach((s, i) => { if (s.pos == null || isNaN(s.pos)) s.pos = stops.length === 1 ? 0 : i / (stops.length - 1); });
    return { angle, stops, radial: m[1] === "radial" };
  }
  function gradXml(g, opacity) {
    const gs = g.stops.map(s => {
      const a = s.c.a * opacity;
      return `<a:gs pos="${Math.round(Math.max(0, Math.min(1, s.pos)) * 100000)}"><a:srgbClr val="${hex(s.c)}">${a < 1 ? `<a:alpha val="${Math.round(a * 100000)}"/>` : ""}</a:srgbClr></a:gs>`;
    }).join("");
    const ang = (((g.ooxmlAngle ?? (g.angle - 90)) % 360) + 360) % 360;
    const shade = g.radial ? `<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>` : `<a:lin ang="${Math.round(ang * 60000)}" scaled="0"/>`;
    return `<a:gradFill rotWithShape="1"><a:gsLst>${gs}</a:gsLst>${shade}</a:gradFill>`;
  }
  const midColor = g => g.stops[Math.floor((g.stops.length - 1) / 2)].c;

  /* ---------- Utilidades de estilo ---------- */
  // Fondo aproximado detrás de un elemento (para mezclar textos semitransparentes)
  function backdrop(el, root) {
    for (let e = el; e && e !== root.parentElement; e = e.parentElement) {
      const cs = getComputedStyle(e);
      const g = parseBackground(cs.backgroundImage);
      if (g) return midColor(g);
      const c = parseColor(cs.backgroundColor);
      if (c && c.a > .5) return c;
    }
    return { r: 255, g: 255, b: 255, a: 1 };
  }
  function textColor(cs, el, root, extraOpacity = 1) {
    let c = parseColor(cs.color) || { r: 0, g: 0, b: 0, a: 1 };
    c = { ...c, a: c.a * extraOpacity };
    return hex(c.a < 1 ? blend(c, backdrop(el, root)) : c);
  }
  const fontFace = cs => (cs.fontFamily.split(",")[0] || "Calibri").replace(/["']/g, "").trim();
  function transformText(t, cs) {
    if (cs.textTransform === "uppercase") return t.toUpperCase();
    if (cs.textTransform === "lowercase") return t.toLowerCase();
    if (cs.textTransform === "capitalize") return t.replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toUpperCase());
    return t;
  }
  function runOpts(cs, el, root, opacity) {
    const ls = parseFloat(cs.letterSpacing);
    const o = {
      fontFace: fontFace(cs),
      fontSize: r2(parseFloat(cs.fontSize) * PT),
      color: textColor(cs, el, root, opacity),
      bold: (parseInt(cs.fontWeight, 10) || 400) >= 600,
      italic: cs.fontStyle === "italic",
    };
    if (ls) o.charSpacing = r2(ls * PT);
    if (/underline/.test(cs.textDecorationLine)) o.underline = { style: "sng" };
    if (/line-through/.test(cs.textDecorationLine)) o.strike = "sngStrike";
    return o;
  }
  const alignOf = cs => /center/.test(cs.textAlign) ? "center" : /right|end/.test(cs.textAlign) ? "right" : "left";

  /* ---------- Formas HTML ---------- */
  function addBox(ctx, el, cs, op) {
    const g = parseBackground(cs.backgroundImage);
    const bg = parseColor(cs.backgroundColor);
    const bw = ["Top", "Right", "Bottom", "Left"].map(s => cs[`border${s}Style`] !== "none" && cs[`border${s}Style`] !== "hidden" ? parseFloat(cs[`border${s}Width`]) || 0 : 0);
    const bc = ["Top", "Right", "Bottom", "Left"].map(s => parseColor(cs[`border${s}Color`]));
    const hasBorder = bw.some((w, i) => w > 0 && bc[i] && bc[i].a > 0);
    if (!g && !(bg && bg.a > 0) && !hasBorder) return;

    // Geometría: caja sin transformar + rotación de la matriz (las franjas diagonales van rotadas)
    const rect = el.getBoundingClientRect();
    let w = el.offsetWidth, h = el.offsetHeight, rot = 0;
    if (cs.transform && cs.transform !== "none") {
      const m = new DOMMatrix(cs.transform);
      rot = Math.atan2(m.b, m.a) * 180 / Math.PI;
      w *= Math.hypot(m.a, m.b); h *= Math.hypot(m.c, m.d);
    }
    if (!w || !h || el.offsetWidth === undefined) { w = rect.width; h = rect.height; }
    if (w < .5 || h < .5) return;
    // Adornos que sobresalen de un contenedor con overflow:hidden (círculos de esquina): PowerPoint no recorta, se omiten
    for (let e = el.parentElement; e && e !== ctx.root; e = e.parentElement) {
      const ecs = getComputedStyle(e);
      if (ecs.overflow === "visible" && ecs.overflowX === "visible") continue;
      const er = e.getBoundingClientRect();
      if (rect.left < er.left - 2 || rect.top < er.top - 2 || rect.right > er.right + 2 || rect.bottom > er.bottom + 2) return;
      break;
    }
    const cx = rect.left + rect.width / 2 - ctx.ox, cy = rect.top + rect.height / 2 - ctx.oy;
    const x = cx - w / 2, y = cy - h / 2;

    const rad = Math.min(parseFloat(cs.borderTopLeftRadius) || 0, w / 2, h / 2);
    const isEllipse = rad > 0 && Math.abs(rad - Math.min(w, h) / 2) < 1 && /%/.test(cs.borderTopLeftRadius) ? true
      : rad > 0 && Math.abs(w - h) < 1 && rad >= w / 2 - .5;
    const shape = isEllipse ? "ellipse" : rad > .5 ? "roundRect" : "rect";
    const o = { x: r2(x * PX), y: r2(y * PX), w: r2(w * PX), h: r2(h * PX) };
    if (rot) o.rotate = r2(((rot % 360) + 360) % 360);
    if (shape === "roundRect") o.rectRadius = r2(rad * PX);

    if (g) {
      const name = "xpg-" + (++gradSeq);
      grads[name] = { fill: gradXml(g, op) };
      o.objectName = name;
      o.fill = { color: hex(midColor(g)) };
    } else if (bg && bg.a > 0) {
      o.fill = { color: hex(bg), transparency: transp(bg.a * op) };
    } else o.fill = { type: "none" };

    const uniform = bw.every(v => v === bw[0]) && bc.every(c => c && bc[0] && hex(c) === hex(bc[0]) && c.a === bc[0].a);
    if (hasBorder && uniform) {
      o.line = { color: hex(bc[0]), width: r2(bw[0] * PT), transparency: transp(bc[0].a * op) };
      if (/dash/.test(cs.borderTopStyle)) o.line.dashType = "dash";
      if (/dot/.test(cs.borderTopStyle)) o.line.dashType = "sysDot";
    }
    // Sombra suave (aproximada)
    const sh = cs.boxShadow && cs.boxShadow !== "none" ? cs.boxShadow.match(/(rgba?\([^)]*\))\s+(-?[\d.]+)px\s+(-?[\d.]+)px\s+([\d.]+)px/) : null;
    if (sh && !rot) {
      const sc = parseColor(sh[1]);
      if (sc && sc.a > .05) o.shadow = { type: "outer", color: hex(sc), opacity: r2(Math.min(.35, sc.a * .45) * op), blur: r2(Math.min(+sh[4], 40) * PT), offset: r2(Math.max(0, +sh[3]) * PT * .4), angle: 90 };
    }
    ctx.s.addShape(shape, o);

    // Bordes de un solo lado (subrayados de encabezado, separadores)
    if (hasBorder && !uniform) {
      const r = { l: rect.left - ctx.ox, t: rect.top - ctx.oy, w: rect.width, h: rect.height };
      [[0, r.l, r.t, r.w, bw[0]], [1, r.l + r.w - bw[1], r.t, bw[1], r.h], [2, r.l, r.t + r.h - bw[2], r.w, bw[2]], [3, r.l, r.t, bw[3], r.h]].forEach(([i, X, Y, W, H]) => {
        if (!bw[i] || !bc[i] || !bc[i].a) return;
        ctx.s.addShape("rect", { x: r2(X * PX), y: r2(Y * PX), w: r2(Math.max(W, .5) * PX), h: r2(Math.max(H, .5) * PX), fill: { color: hex(bc[i]), transparency: transp(bc[i].a * op) } });
      });
    }
  }

  /* ---------- Textos HTML ---------- */
  const INLINE = new Set(["inline", "inline-block"]);
  // ¿Contenedor de texto "puro"? (solo texto y elementos en línea, sin elementos de bloque dentro)
  function isTextBlock(el) {
    let hasText = false;
    const ok = node => {
      for (const ch of node.childNodes) {
        if (ch.nodeType === 3) { if (ch.data.trim()) hasText = true; continue; }
        if (ch.nodeType !== 1) continue;
        if (ch.tagName === "BR") continue;
        if (ch instanceof SVGElement) return false;
        const d = getComputedStyle(ch).display;
        if (d === "none") continue;
        if (d !== "inline") return false;
        if (!ok(ch)) return false;
      }
      return true;
    };
    return ok(el) && hasText;
  }
  function textRects(node) {
    const rg = document.createRange();
    rg.selectNodeContents(node);
    return [...rg.getClientRects()].filter(r => r.width > 0 && r.height > 0);
  }
  function unionRect(rs) {
    const l = Math.min(...rs.map(r => r.left)), t = Math.min(...rs.map(r => r.top));
    const R = Math.max(...rs.map(r => r.right)), B = Math.max(...rs.map(r => r.bottom));
    return { left: l, top: t, width: R - l, height: B - t };
  }
  function addTextBox(ctx, runs, rects, cs, el) {
    if (!runs.length || !rects.length) return;
    const u = unionRect(rects);
    const lines = new Set(rects.map(r => Math.round(r.top))).size;
    const align = alignOf(cs);
    const multi = lines > 1;
    // margen extra para que la diferencia de métricas de fuente no corte ni parta el texto
    const extra = multi ? u.width * .03 + 4 : u.width * .12 + 12;
    let x = u.left - ctx.ox, w = u.width + extra;
    if (align === "center") x -= extra / 2;
    else if (align === "right") x -= extra;
    const lh = parseFloat(cs.lineHeight);
    const o = {
      x: r2(x * PX), y: r2((u.top - ctx.oy) * PX), w: r2(w * PX), h: r2(Math.max(u.height, 4) * PX),
      margin: 0, valign: multi ? "top" : "middle", align, wrap: multi, isTextBox: true, fit: "none",
      paraSpaceBefore: 0, paraSpaceAfter: 0,
    };
    if (multi && lh) o.lineSpacing = r2(lh * PT);
    ctx.s.addText(runs, o);
  }
  // Texto de un contenedor puro: un solo cuadro con varios formatos (negritas, colores…)
  function textBlock(ctx, el, op) {
    const runs = [], rects = [];
    const walk = (node, inhOp) => {
      for (const ch of node.childNodes) {
        if (ch.nodeType === 3) {
          const t = ch.data.replace(/\s+/g, " ");
          if (!t) continue;
          const pcs = getComputedStyle(ch.parentElement);
          runs.push({ text: transformText(t, pcs), options: runOpts(pcs, ch.parentElement, ctx.root, inhOp) });
          rects.push(...textRects(ch));
        } else if (ch.tagName === "BR") {
          if (runs.length) runs[runs.length - 1].options.breakLine = true;
        } else if (ch.nodeType === 1) {
          const ccs = getComputedStyle(ch);
          if (ccs.display === "none" || ccs.visibility === "hidden") continue;
          walk(ch, inhOp * parseFloat(ccs.opacity || 1));
        }
      }
    };
    walk(el, op);
    // quitar espacios al principio y al final
    while (runs.length && !runs[0].text.trim()) runs.shift();
    while (runs.length && !runs[runs.length - 1].text.trim()) runs.pop();
    if (!runs.length) return;
    runs[0].text = runs[0].text.replace(/^\s+/, "");
    runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, "");
    addTextBox(ctx, runs, rects, getComputedStyle(el), el);
  }
  function textNode(ctx, node, op) {
    const t = node.data.replace(/\s+/g, " ").trim();
    if (!t) return;
    const el = node.parentElement, cs = getComputedStyle(el);
    addTextBox(ctx, [{ text: transformText(t, cs), options: runOpts(cs, el, ctx.root, op) }], textRects(node), cs, el);
  }

  /* ---------- SVG ---------- */
  function pathPoints(d) {
    // Devuelve subtrayectos con comandos absolutos M/L/C/Q/Z (arcos convertidos a curvas)
    const toks = d.match(/[a-zA-Z]|-?(?:\d*\.\d+|\d+\.?)(?:e[-+]?\d+)?/g) || [];
    const out = []; let i = 0, cmd = "", x = 0, y = 0, sx = 0, sy = 0, lc = null;
    const num = () => parseFloat(toks[i++]);
    while (i < toks.length) {
      if (/[a-zA-Z]/.test(toks[i])) cmd = toks[i++];
      const rel = cmd === cmd.toLowerCase(), C = cmd.toUpperCase();
      const ax = v => rel ? x + v : v, ay = v => rel ? y + v : v;
      if (C === "Z") { out.push({ t: "Z" }); x = sx; y = sy; lc = null; continue; }
      if (C === "M") { x = ax(num()); y = ay(num()); sx = x; sy = y; out.push({ t: "M", x, y }); cmd = rel ? "l" : "L"; lc = null; continue; }
      if (C === "L") { x = ax(num()); y = ay(num()); out.push({ t: "L", x, y }); lc = null; continue; }
      if (C === "H") { x = rel ? x + num() : num(); out.push({ t: "L", x, y }); lc = null; continue; }
      if (C === "V") { y = rel ? y + num() : num(); out.push({ t: "L", x, y }); lc = null; continue; }
      if (C === "C" || C === "S") {
        let x1, y1;
        if (C === "C") { x1 = ax(num()); y1 = ay(num()); }
        else { x1 = lc ? 2 * x - lc[0] : x; y1 = lc ? 2 * y - lc[1] : y; }
        const x2 = ax(num()), y2 = ay(num()), nx = ax(num()), ny = ay(num());
        out.push({ t: "C", x1, y1, x2, y2, x: nx, y: ny }); lc = [x2, y2]; x = nx; y = ny; continue;
      }
      if (C === "Q" || C === "T") {
        let x1, y1;
        if (C === "Q") { x1 = ax(num()); y1 = ay(num()); } else { x1 = lc ? 2 * x - lc[0] : x; y1 = lc ? 2 * y - lc[1] : y; }
        const nx = ax(num()), ny = ay(num());
        out.push({ t: "Q", x1, y1, x: nx, y: ny }); lc = [x1, y1]; x = nx; y = ny; continue;
      }
      if (C === "A") {
        const rx = Math.abs(num()), ry = Math.abs(num()), phi = num() * Math.PI / 180, fa = num(), fs = num();
        const nx = ax(num()), ny = ay(num());
        arcToCubic(x, y, rx, ry, phi, fa, fs, nx, ny).forEach(c => out.push(c));
        x = nx; y = ny; lc = null; continue;
      }
      i++; // comando desconocido
    }
    return out;
  }
  function arcToCubic(x1, y1, rx, ry, phi, fa, fs, x2, y2) {
    if (!rx || !ry) return [{ t: "L", x: x2, y: y2 }];
    const cos = Math.cos(phi), sin = Math.sin(phi);
    const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2;
    const x1p = cos * dx + sin * dy, y1p = -sin * dx + cos * dy;
    let lam = x1p * x1p / (rx * rx) + y1p * y1p / (ry * ry);
    if (lam > 1) { rx *= Math.sqrt(lam); ry *= Math.sqrt(lam); }
    const sgn = fa === fs ? -1 : 1;
    const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
    const co = sgn * Math.sqrt(Math.max(0, num / (rx * rx * y1p * y1p + ry * ry * x1p * x1p)));
    const cxp = co * rx * y1p / ry, cyp = -co * ry * x1p / rx;
    const cx = cos * cxp - sin * cyp + (x1 + x2) / 2, cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
    const ang = (ux, uy, vx, vy) => { const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy); return a; };
    let t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
    let dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
    if (!fs && dt > 0) dt -= 2 * Math.PI; else if (fs && dt < 0) dt += 2 * Math.PI;
    const segs = Math.ceil(Math.abs(dt) / (Math.PI / 2)), d = dt / segs, k = 4 / 3 * Math.tan(d / 4);
    const pt = (t) => [cx + rx * Math.cos(t) * cos - ry * Math.sin(t) * sin, cy + rx * Math.cos(t) * sin + ry * Math.sin(t) * cos];
    const dpt = (t) => [-rx * Math.sin(t) * cos - ry * Math.cos(t) * sin, -rx * Math.sin(t) * sin + ry * Math.cos(t) * cos];
    const out = [];
    for (let s = 0; s < segs; s++) {
      const a = t1 + s * d, b = a + d, p0 = pt(a), p1 = pt(b), d0 = dpt(a), d1 = dpt(b);
      out.push({ t: "C", x1: p0[0] + k * d0[0], y1: p0[1] + k * d0[1], x2: p1[0] - k * d1[0], y2: p1[1] - k * d1[1], x: p1[0], y: p1[1] });
    }
    return out;
  }
  function svgPaint(ctx, el, cs, kind) {
    const v = kind === "fill" ? cs.fill : cs.stroke;
    if (!v || v === "none") return null;
    const op = parseFloat(kind === "fill" ? cs.fillOpacity : cs.strokeOpacity) * ctx.op(el);
    const m = v.match(/url\("?#([^")]+)"?\)/);
    if (m) {
      const gEl = document.getElementById(m[1]);
      if (!gEl) return null;
      const stops = [...gEl.querySelectorAll("stop")].map(s => {
        const scs = getComputedStyle(s);
        const c = parseColor(scs.stopColor) || { r: 0, g: 0, b: 0, a: 1 };
        const off = s.getAttribute("offset") || "0";
        return { c: { ...c, a: c.a * parseFloat(scs.stopOpacity || 1) }, pos: off.endsWith("%") ? parseFloat(off) / 100 : parseFloat(off) };
      });
      if (!stops.length) return null;
      const g = { stops, radial: gEl.tagName.toLowerCase() === "radialgradient" };
      if (!g.radial) {
        const num = (a, d) => { const s = gEl.getAttribute(a); return s == null ? d : s.endsWith("%") ? parseFloat(s) / 100 : parseFloat(s); };
        const x1 = num("x1", 0), y1 = num("y1", 0), x2 = num("x2", 1), y2 = num("y2", 0);
        let dx = x2 - x1, dy = y2 - y1;
        if (gEl.getAttribute("gradientUnits") === "userSpaceOnUse") {
          const M = el.getScreenCTM(); // dirección en pantalla
          const vx = M.a * dx + M.c * dy, vy = M.b * dx + M.d * dy; dx = vx; dy = vy;
        }
        g.ooxmlAngle = Math.atan2(dy, dx) * 180 / Math.PI;
      }
      return { grad: g, op, color: midColor(g) };
    }
    const c = parseColor(v);
    return c ? { color: c, op: op * c.a } : null;
  }
  // Patrón de guiones real (null si es la animación de "dibujar la línea", que al final queda continua)
  function dashInfo(el, cs) {
    const da = cs.strokeDasharray;
    if (!da || da === "none") return null;
    const first = parseFloat(da);
    const pl = parseFloat(el.getAttribute("pathLength"));
    let total = pl || 0;
    if (!total && el.getTotalLength) { try { total = el.getTotalLength(); } catch (e) { total = 0; } }
    if (total && first >= total * .99) return { solid: true, frac: Math.max(0, Math.min(1, (total - (parseFloat((cs.strokeDashoffset.match(/-?[\d.]+/) || [0])[0]) || 0)) / total)) };
    return { dash: first };
  }
  function applyPaint(o, fill, stroke, cs, scale, isLine, el) {
    const named = [];
    if (fill && !isLine) {
      if (fill.grad) named.push(["fill", gradXml(fill.grad, fill.op)]);
      o.fill = { color: hex(fill.color), transparency: transp(fill.op) };
    } else o.fill = { type: "none" };
    if (stroke) {
      const sw = parseFloat(cs.strokeWidth) * scale;
      o.line = { color: hex(stroke.color), width: r2(Math.max(sw, .25) * PT), transparency: transp(stroke.op) };
      const di = el && dashInfo(el, cs);
      if (di && di.dash) {
        const first = di.dash * scale;
        o.line.dashType = first <= sw * 1.5 ? "sysDot" : first < 6 ? "sysDash" : "dash";
      }
      if (cs.strokeLinecap === "round") o.line.capType = "round";
      if (stroke.grad) named.push(["line", gradXml(stroke.grad, stroke.op)]);
    }
    if (named.length) {
      const name = "xpg-" + (++gradSeq);
      grads[name] = Object.fromEntries(named);
      o.objectName = name;
    }
  }
  function addPath(ctx, el, cs, segs, closedHint) {
    const M = el.getScreenCTM();
    const tp = (x, y) => { const p = new DOMPoint(x, y).matrixTransform(M); return [p.x - ctx.ox, p.y - ctx.oy]; };
    const all = [];
    const pts = segs.map(s => {
      if (s.t === "Z") return s;
      const o = { ...s };
      [o.x, o.y] = tp(s.x, s.y); all.push([o.x, o.y]);
      if (s.t === "C" || s.t === "Q") { [o.x1, o.y1] = tp(s.x1, s.y1); all.push([o.x1, o.y1]); }
      if (s.t === "C") { [o.x2, o.y2] = tp(s.x2, s.y2); all.push([o.x2, o.y2]); }
      return o;
    });
    if (!all.length) return;
    let l = Math.min(...all.map(p => p[0])), t = Math.min(...all.map(p => p[1]));
    let R = Math.max(...all.map(p => p[0])), B = Math.max(...all.map(p => p[1]));
    if (R - l < 1) { l -= .5; R += .5; }
    if (B - t < 1) { t -= .5; B += .5; }
    const rel = (x, y) => ({ x: r2((x - l) * PX), y: r2((y - t) * PX) });
    const points = pts.map((s, i) => {
      if (s.t === "Z") return { close: true };
      const p = rel(s.x, s.y);
      if (s.t === "M") return { ...p, moveTo: true };
      if (s.t === "L") return p;
      const c1 = rel(s.x1, s.y1);
      if (s.t === "Q") return { ...p, curve: { type: "quadratic", x1: c1.x, y1: c1.y } };
      const c2 = rel(s.x2, s.y2);
      return { ...p, curve: { type: "cubic", x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y } };
    });
    const scale = Math.hypot(M.a, M.b);
    const fill = svgPaint(ctx, el, cs, "fill"), stroke = svgPaint(ctx, el, cs, "stroke");
    if (!fill && !stroke) return;
    const o = { x: r2(l * PX), y: r2(t * PX), w: r2((R - l) * PX), h: r2((B - t) * PX), points };
    applyPaint(o, fill, stroke, cs, scale, false, el);
    ctx.s.addShape("custGeom", o);
  }
  function svgShape(ctx, el) {
    const cs = getComputedStyle(el);
    if (cs.display === "none" || cs.visibility === "hidden" || ctx.op(el) <= 0.01) return;
    const tag = el.tagName.toLowerCase();
    const M = el.getScreenCTM();
    if (!M) return;
    const scale = Math.hypot(M.a, M.b);
    const n = a => parseFloat(el.getAttribute(a)) || 0;
    if (tag === "circle" && cs.stroke !== "none") {
      const di = dashInfo(el, cs);
      if (di && di.solid && di.frac < .999) {
        if (di.frac <= .001) return;
        const cx = n("cx"), cy = n("cy"), r = n("r"), a = di.frac * 2 * Math.PI;
        const ex = cx + r * Math.cos(a), ey = cy + r * Math.sin(a);
        addPath(ctx, el, cs, [{ t: "M", x: cx + r, y: cy }, ...arcToCubic(cx + r, cy, r, r, 0, a > Math.PI ? 1 : 0, 1, ex, ey)]);
        return;
      }
    }
    if (tag === "rect" || tag === "circle" || tag === "ellipse") {
      let x, y, w, h, shape = "rect", rad = 0;
      if (tag === "rect") { x = n("x"); y = n("y"); w = el.width.baseVal.value; h = el.height.baseVal.value; rad = n("rx") || n("ry"); }
      else { const rx = tag === "circle" ? n("r") : n("rx"), ry = tag === "circle" ? n("r") : n("ry"); x = n("cx") - rx; y = n("cy") - ry; w = 2 * rx; h = 2 * ry; shape = "ellipse"; }
      if (w <= 0 || h <= 0) return;
      const rot = Math.atan2(M.b, M.a) * 180 / Math.PI;
      const c = new DOMPoint(x + w / 2, y + h / 2).matrixTransform(M);
      const W = w * scale, H = h * Math.hypot(M.c, M.d);
      const o = { x: r2((c.x - ctx.ox - W / 2) * PX), y: r2((c.y - ctx.oy - H / 2) * PX), w: r2(W * PX), h: r2(H * PX) };
      if (Math.abs(rot) > .1) o.rotate = r2(((rot % 360) + 360) % 360);
      if (shape === "rect" && rad > 0) { shape = "roundRect"; o.rectRadius = r2(Math.min(rad * scale, W / 2, H / 2) * PX); }
      const fill = svgPaint(ctx, el, cs, "fill"), stroke = svgPaint(ctx, el, cs, "stroke");
      if (!fill && !stroke) return;
      applyPaint(o, fill, stroke, cs, scale, false, el);
      ctx.s.addShape(shape, o);
    } else if (tag === "line") {
      addPath(ctx, el, cs, [{ t: "M", x: n("x1"), y: n("y1") }, { t: "L", x: n("x2"), y: n("y2") }]);
    } else if (tag === "polyline" || tag === "polygon") {
      const p = (el.getAttribute("points") || "").trim().split(/[\s,]+/).map(parseFloat);
      const segs = [];
      for (let i = 0; i + 1 < p.length; i += 2) segs.push({ t: i ? "L" : "M", x: p[i], y: p[i + 1] });
      if (tag === "polygon") segs.push({ t: "Z" });
      addPath(ctx, el, cs, segs);
    } else if (tag === "path") {
      addPath(ctx, el, cs, pathPoints(el.getAttribute("d") || ""));
    } else if (tag === "text") {
      const t = el.textContent.replace(/\s+/g, " ").trim();
      if (!t) return;
      const bb = el.getBBox();
      if (!bb.width) return;
      const c = new DOMPoint(bb.x + bb.width / 2, bb.y + bb.height / 2).matrixTransform(M);
      const rot = Math.atan2(M.b, M.a) * 180 / Math.PI;
      const W = bb.width * scale, H = bb.height * scale, extra = W * .12 + 10;
      const anchor = cs.textAnchor;
      const align = Math.abs(rot) > 1 ? "center" : anchor === "middle" ? "center" : anchor === "end" ? "right" : "left";
      let x = c.x - ctx.ox - W / 2;
      if (align === "center") x -= extra / 2; else if (align === "right") x -= extra;
      const fill = svgPaint(ctx, el, cs, "fill") || { color: { r: 0, g: 0, b: 0, a: 1 }, op: 1 };
      const col = fill.op < 1 ? blend({ ...fill.color, a: fill.op }, { r: 255, g: 255, b: 255, a: 1 }) : fill.color;
      const ls = parseFloat(cs.letterSpacing);
      const o = {
        x: r2(x * PX), y: r2((c.y - ctx.oy - H / 2) * PX), w: r2((W + extra) * PX), h: r2(H * PX),
        margin: 0, valign: "middle", align, wrap: false, isTextBox: true, fit: "none",
        fontFace: fontFace(cs), fontSize: r2(parseFloat(cs.fontSize) * scale * PT), color: hex(col),
        bold: (parseInt(cs.fontWeight, 10) || 400) >= 600, italic: cs.fontStyle === "italic",
      };
      if (ls) o.charSpacing = r2(ls * scale * PT);
      if (Math.abs(rot) > .1) o.rotate = r2(((rot % 360) + 360) % 360);
      ctx.s.addText(transformText(t, cs), o);
      return;
    }
  }
  function walkSvg(ctx, node) {
    for (const ch of node.children) {
      const tag = ch.tagName.toLowerCase();
      if (["defs", "clippath", "lineargradient", "radialgradient", "marker", "mask", "pattern", "symbol", "style", "title", "desc", "animate", "animatemotion", "animatetransform", "set", "metadata"].includes(tag)) continue;
      if (ch.querySelector(":scope > animateMotion")) continue;   // partículas animadas
      if (tag === "g" || tag === "a" || tag === "svg") {
        const cs = getComputedStyle(ch);
        if (cs.display === "none" || cs.visibility === "hidden") continue;
        walkSvg(ctx, ch);
      } else if (tag === "foreignobject") {
        for (const c of ch.children) visit(ctx, c);
      } else svgShape(ctx, ch);
    }
  }

  /* ---------- Recorrido del slide ---------- */
  function visit(ctx, el) {
    const cs = getComputedStyle(el);
    if (cs.display === "none" || cs.visibility === "hidden") return;
    const op = ctx.op(el);
    if (op <= 0.01) return;
    if (el instanceof SVGSVGElement) { walkSvg(ctx, el); return; }
    if (el instanceof SVGElement) return;
    if (el.tagName === "IMG" && el.complete && el.naturalWidth) {
      const r = el.getBoundingClientRect();
      try {
        const cv = document.createElement("canvas"); cv.width = el.naturalWidth; cv.height = el.naturalHeight;
        cv.getContext("2d").drawImage(el, 0, 0);
        ctx.s.addImage({ data: cv.toDataURL("image/png"), x: r2((r.left - ctx.ox) * PX), y: r2((r.top - ctx.oy) * PX), w: r2(r.width * PX), h: r2(r.height * PX) });
      } catch (e) { /* imagen de otro dominio: se omite */ }
      return;
    }
    addBox(ctx, el, cs, op);
    if (isTextBlock(el)) {
      // fondos de los elementos en línea (pastillas, resaltados) antes que el texto
      el.querySelectorAll("*").forEach(ch => { const ccs = getComputedStyle(ch); if (ccs.display !== "none") addBox(ctx, ch, ccs, ctx.op(ch)); });
      textBlock(ctx, el, op);
      return;
    }
    for (const ch of el.childNodes) {
      if (ch.nodeType === 3) textNode(ctx, ch, op);
      else if (ch.nodeType === 1) visit(ctx, ch);
    }
  }
  // Los ::before / ::after se convierten en elementos reales mientras dura la conversión
  function materializePseudos(root) {
    const made = [];
    const style = document.createElement("style");
    style.textContent = ".xp-np::before,.xp-np::after{content:none!important}";
    const hosts = [root, ...root.querySelectorAll("*")].filter(e => !(e instanceof SVGElement));
    const todo = [];
    hosts.forEach(el => ["::before", "::after"].forEach(p => {
      const pcs = getComputedStyle(el, p);
      if (!pcs.content || pcs.content === "none" || pcs.content === "normal") return;
      const sp = document.createElement("span");
      for (let i = 0; i < pcs.length; i++) { const k = pcs[i]; sp.style.setProperty(k, pcs.getPropertyValue(k)); }
      const txt = pcs.content.match(/^"(.*)"$/s);
      sp.textContent = txt ? txt[1].replace(/\\"/g, '"') : "";
      todo.push([el, p, sp]);
    }));
    todo.forEach(([el, p, sp]) => {
      if (p === "::before") el.insertBefore(sp, el.firstChild); else el.appendChild(sp);
      el.classList.add("xp-np"); made.push(sp);
    });
    document.head.appendChild(style);
    return () => { made.forEach(s => s.remove()); hosts.forEach(e => e.classList.remove("xp-np")); style.remove(); };
  }

  function slide(pptx, s, root) {
    const undo = materializePseudos(root);
    try {
      const r = root.getBoundingClientRect();
      const opCache = new Map();
      const ctx = {
        s, root, ox: r.left, oy: r.top,
        op: el => {
          if (opCache.has(el)) return opCache.get(el);
          let o = 1;
          for (let e = el; e && e !== root.parentElement; e = e.parentElement) {
            if (opCache.has(e) && e !== el) { o *= opCache.get(e); break; }
            o *= parseFloat(getComputedStyle(e).opacity || 1);
          }
          opCache.set(el, o);
          return o;
        },
      };
      const cs = getComputedStyle(root);
      const g = parseBackground(cs.backgroundImage);
      const bg = parseColor(cs.backgroundColor);
      if (!g && bg) s.background = { color: hex(bg) };
      // el propio slide: su fondo (degradado en la portada) y luego su contenido
      if (g) addBox(ctx, root, cs, 1);
      for (const ch of root.childNodes) {
        if (ch.nodeType === 3) textNode(ctx, ch, 1);
        else if (ch.nodeType === 1) visit(ctx, ch);
      }
    } finally { undo(); }
  }

  // Guarda el archivo; aplica los degradados (PptxGenJS solo escribe rellenos sólidos)
  async function save(pptx, fileName) {
    const buf = await pptx.write({ outputType: "arraybuffer" });
    const zip = await JSZip.loadAsync(buf);
    const files = Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f));
    for (const f of files) {
      let xml = await zip.file(f).async("string");
      if (!xml.includes('name="xpg-')) continue;
      xml = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, sp => {
        const m = sp.match(/name="(xpg-\d+)"/);
        if (!m || !grads[m[1]]) return sp;
        const gd = grads[m[1]];
        const i = sp.indexOf("<p:spPr>"), j = sp.indexOf("</p:spPr>");
        if (i < 0 || j < 0) return sp;
        let pr = sp.slice(i, j);
        const lnM = pr.slice(Math.max(0, pr.indexOf("</a:custGeom>"))).search(/<a:ln[ >]/);
        const lnAt = lnM < 0 ? -1 : lnM + Math.max(0, pr.indexOf("</a:custGeom>"));
        if (gd.line && lnAt >= 0) {
          const ln = pr.slice(lnAt).replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, gd.line);
          pr = pr.slice(0, lnAt) + ln;
        }
        if (gd.fill) {
          const head = lnAt >= 0 ? pr.slice(0, lnAt) : pr, tail = lnAt >= 0 ? pr.slice(lnAt) : "";
          const cg = head.indexOf("</a:custGeom>");
          const at = cg < 0 ? 0 : cg;
          pr = head.slice(0, at) + head.slice(at).replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, gd.fill) + tail;
        }
        return sp.slice(0, i) + pr + sp.slice(j);
      });
      zip.file(f, xml);
    }
    const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", compression: "DEFLATE" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = fileName;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  window.xpPptx = { slide, save };
})();
