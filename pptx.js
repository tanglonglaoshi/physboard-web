/* ==========================================================================
 * pptx.js — 轻量 PPTX 渲染器（离线、无外部依赖）
 * 把 .pptx 的每张幻灯片画到 canvas 上，供白板作为页面背景使用。
 * 支持：纯色/图片/渐变背景、文本框、常见自选图形、直线/连接线、
 *       图片（含简单裁剪）、组合、表格；布局/母版的占位符继承。
 * 不支持：图表、SmartArt、动画、艺术字、3D 等复杂效果（会画占位块）。
 * ========================================================================== */
(function (global) {
  'use strict';

  const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const EMU_PER_PT = 12700;

  /* ----------------------------- ZIP ----------------------------- */
  function findEOCD(b) {
    const min = Math.max(0, b.length - 66000);
    for (let i = b.length - 22; i >= min; i--) {
      if (b[i] === 0x50 && b[i + 1] === 0x4b && b[i + 2] === 0x05 && b[i + 3] === 0x06) return i;
    }
    return -1;
  }

  function Zip(buffer) {
    const b = new Uint8Array(buffer);
    this.buf = b;
    this.entries = Object.create(null);
    const eocd = findEOCD(b);
    if (eocd < 0) throw new Error('不是有效的 .pptx 文件（ZIP 结构缺失）');
    const count = b[eocd + 10] | (b[eocd + 11] << 8);
    let off = b[eocd + 16] | (b[eocd + 17] << 8) | (b[eocd + 18] << 16) | (b[eocd + 19] << 24);
    const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    const dec = new TextDecoder('utf-8');
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(off, true) !== 0x02014b50) break;
      const method = dv.getUint16(off + 10, true);
      const compSize = dv.getUint32(off + 20, true);
      const nameLen = dv.getUint16(off + 28, true);
      const extraLen = dv.getUint16(off + 30, true);
      const commentLen = dv.getUint16(off + 32, true);
      const localOff = dv.getUint32(off + 42, true);
      const name = dec.decode(b.subarray(off + 46, off + 46 + nameLen));
      this.entries[name] = { method, compSize, localOff };
      off += 46 + nameLen + extraLen + commentLen;
    }
  }
  Zip.prototype.has = function (name) { return !!this.entries[name]; };
  Zip.prototype.names = function () { return Object.keys(this.entries); };
  Zip.prototype.bytes = async function (name) {
    const e = this.entries[name];
    if (!e) return null;
    const b = this.buf;
    const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    if (dv.getUint32(e.localOff, true) !== 0x04034b50) throw new Error('ZIP 局部头损坏');
    const nameLen = dv.getUint16(e.localOff + 26, true);
    const extraLen = dv.getUint16(e.localOff + 28, true);
    const start = e.localOff + 30 + nameLen + extraLen;
    const data = b.subarray(start, start + e.compSize);
    if (e.method === 0) return data;
    if (e.method !== 8) throw new Error('不支持的压缩方式：' + e.method);
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('当前浏览器不支持解压 .pptx（缺少 DecompressionStream）');
    }
    const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  };

  /* ----------------------------- XML ----------------------------- */
  const parser = new DOMParser();
  function parseXml(bytes) {
    let text;
    if (bytes[0] === 0xff && bytes[1] === 0xfe) text = new TextDecoder('utf-16le').decode(bytes);
    else if (bytes[0] === 0xfe && bytes[1] === 0xff) text = new TextDecoder('utf-16be').decode(bytes);
    else text = new TextDecoder('utf-8').decode(bytes);
    const doc = parser.parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('XML 解析失败');
    return doc;
  }
  function kids(el, name) {
    const out = [];
    if (!el) return out;
    for (let n = el.firstElementChild; n; n = n.nextElementSibling) {
      if (n.localName === name) out.push(n);
    }
    return out;
  }
  function kid(el, name) { return kids(el, name)[0] || null; }
  function localAll(el, name) {
    return el ? Array.from(el.getElementsByTagName('*')).filter(n => n.localName === name) : [];
  }
  function numAttr(el, name, def) {
    if (!el) return def;
    const v = el.getAttribute(name);
    const n = parseFloat(v);
    return isFinite(n) ? n : def;
  }
  function rid(el) { return el ? (el.getAttributeNS(REL_NS, 'id') || el.getAttribute('r:embed') || el.getAttribute('r:id')) : null; }

  function resolvePath(basePart, target) {
    if (!target) return null;
    if (target[0] === '/') return target.slice(1);
    const base = basePart.slice(0, basePart.lastIndexOf('/') + 1);
    const parts = (base + target).split('/');
    const out = [];
    for (const p of parts) {
      if (!p || p === '.') continue;
      if (p === '..') out.pop(); else out.push(p);
    }
    return out.join('/');
  }

  async function readRels(zip, part) {
    const p = part.slice(0, part.lastIndexOf('/') + 1) + '_rels/' +
      part.slice(part.lastIndexOf('/') + 1) + '.rels';
    const map = Object.create(null);
    if (!zip.has(p)) return map;
    const doc = parseXml(await zip.bytes(p));
    for (const rel of localAll(doc, 'Relationship')) {
      map[rel.getAttribute('Id')] = {
        target: rel.getAttribute('Target'),
        type: rel.getAttribute('Type') || ''
      };
    }
    return map;
  }

  /* --------------------------- 颜色/主题 --------------------------- */
  function schemeValue(name) {
    const m = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' };
    return m[name] || name;
  }
  function parseTheme(doc) {
    const out = { colors: {}, major: '', minor: '', majorEa: '', minorEa: '' };
    if (!doc) return out;
    const cs = kid(doc.documentElement, 'themeElements');
    const clr = cs && kid(cs, 'clrScheme');
    if (clr) {
      for (const c of clr.children) {
        const srgb = kid(c, 'srgbClr');
        const sys = kid(c, 'sysClr');
        let hex = null;
        if (srgb) hex = srgb.getAttribute('val');
        else if (sys) hex = sys.getAttribute('lastClr');
        if (hex) out.colors[c.localName] = '#' + hex;
      }
    }
    const fs = cs && kid(cs, 'fontScheme');
    if (fs) {
      const major = kid(fs, 'majorFont'), minor = kid(fs, 'minorFont');
      out.major = (major && kid(major, 'latin') && major.getElementsByTagName('*')[0] && kid(major, 'latin').getAttribute('typeface')) || '';
      out.minor = (minor && kid(minor, 'latin') && kid(minor, 'latin').getAttribute('typeface')) || '';
      const mea = major && kid(major, 'ea');
      const nea = minor && kid(minor, 'ea');
      out.majorEa = mea ? mea.getAttribute('typeface') : '';
      out.minorEa = nea ? nea.getAttribute('typeface') : '';
    }
    return out;
  }

  function colorOf(el, theme, fallback) {
    if (!el) return fallback || null;
    // 传进来的可能是 solidFill / 颜色节点
    let node = el;
    if (node.localName === 'solidFill' || node.localName === 'gradFill') {
      node = node.firstElementChild;
    }
    if (!node) return fallback || null;
    let hex = null;
    if (node.localName === 'srgbClr') hex = '#' + node.getAttribute('val');
    else if (node.localName === 'schemeClr') {
      const v = schemeValue(node.getAttribute('val'));
      hex = theme.colors[v] || '#000000';
    } else if (node.localName === 'sysClr') {
      hex = '#' + (node.getAttribute('lastClr') || '000000');
    } else if (node.localName === 'prstClr') {
      hex = node.getAttribute('val') || '#000000';
    }
    if (!hex) return fallback || null;
    let alpha = 1;
    const a = kid(node, 'alpha');
    if (a) alpha = Math.max(0, Math.min(1, numAttr(a, 'val', 100000) / 100000));
    if (alpha >= 0.999) return hex;
    const r = parseInt(hex.slice(1, 3), 16), g = parseInt(hex.slice(3, 5), 16), b = parseInt(hex.slice(5, 7), 16);
    return 'rgba(' + r + ',' + g + ',' + b + ',' + alpha.toFixed(3) + ')';
  }

  function fillOf(shape, theme) {
    const spPr = kid(shape, 'spPr');
    if (!spPr) return null;
    const solid = kid(spPr, 'solidFill');
    if (solid) return colorOf(solid, theme, null);
    if (kid(spPr, 'noFill')) return null;
    const grad = kid(spPr, 'gradFill');
    if (grad) {
      const stops = localAll(grad, 'gs').map(gs => colorOf(gs, theme, null)).filter(Boolean);
      if (stops.length) return stops[0];
    }
    const style = kid(shape, 'style');
    const ref = style && kid(style, 'fillRef');
    if (ref && numAttr(ref, 'idx', 0) > 0) {
      const inherit = ref.firstElementChild;
      if (inherit) return colorOf(inherit, theme, null);
    }
    return null;
  }

  function lineOf(shape, theme) {
    const spPr = kid(shape, 'spPr');
    const ln = spPr && kid(spPr, 'ln');
    let color = null, width = 0;
    if (ln) {
      if (kid(ln, 'noFill')) return null;
      const solid = kid(ln, 'solidFill');
      if (solid) color = colorOf(solid, theme, null);
      width = numAttr(ln, 'w', 0);
    }
    if (!color) {
      const style = kid(shape, 'style');
      const ref = style && kid(style, 'lnRef');
      if (ref && numAttr(ref, 'idx', 0) > 0) {
        const inherit = ref.firstElementChild;
        if (inherit) color = colorOf(inherit, theme, null);
        if (color && !width) width = 12700;
      }
    }
    if (!color) return null;
    return { color: color, width: width || 12700 };
  }

  /* --------------------------- 几何/路径 --------------------------- */
  function xfrmOf(node) {
    const spPr = kid(node, 'spPr') || kid(node, 'grpSpPr');
    const xfrm = spPr && kid(spPr, 'xfrm');
    if (!xfrm) return null;
    const off = kid(xfrm, 'off'), ext = kid(xfrm, 'ext');
    if (!off || !ext) return null;
    return {
      x: numAttr(off, 'x', 0), y: numAttr(off, 'y', 0),
      cx: numAttr(ext, 'cx', 0), cy: numAttr(ext, 'cy', 0),
      rot: numAttr(xfrm, 'rot', 0),
      flipH: xfrm.getAttribute('flipH') === '1',
      flipV: xfrm.getAttribute('flipV') === '1'
    };
  }

  function presetOf(node) {
    const spPr = kid(node, 'spPr');
    const prst = spPr && kid(spPr, 'prstGeom');
    return prst ? prst.getAttribute('prst') : '';
  }

  function pathFor(ctx, prst, x, y, w, h) {
    const h2 = h / 2, w2 = w / 2;
    switch (prst) {
      case 'ellipse':
      case 'circle':
        ctx.beginPath();
        ctx.ellipse(x + w2, y + h2, Math.abs(w2), Math.abs(h2), 0, 0, Math.PI * 2);
        return;
      case 'roundRect':
      case 'round2SameRect': {
        const r = Math.min(Math.abs(w), Math.abs(h)) * 0.16;
        ctx.beginPath();
        ctx.moveTo(x + r, y);
        ctx.lineTo(x + w - r, y);
        ctx.quadraticCurveTo(x + w, y, x + w, y + r);
        ctx.lineTo(x + w, y + h - r);
        ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
        ctx.lineTo(x + r, y + h);
        ctx.quadraticCurveTo(x, y + h, x, y + h - r);
        ctx.lineTo(x, y + r);
        ctx.quadraticCurveTo(x, y, x + r, y);
        ctx.closePath();
        return;
      }
      case 'triangle':
        ctx.beginPath();
        ctx.moveTo(x + w2, y); ctx.lineTo(x + w, y + h); ctx.lineTo(x, y + h); ctx.closePath();
        return;
      case 'rtTriangle':
        ctx.beginPath();
        ctx.moveTo(x, y); ctx.lineTo(x, y + h); ctx.lineTo(x + w, y + h); ctx.closePath();
        return;
      case 'diamond':
        ctx.beginPath();
        ctx.moveTo(x + w2, y); ctx.lineTo(x + w, y + h2); ctx.lineTo(x + w2, y + h); ctx.lineTo(x, y + h2);
        ctx.closePath();
        return;
      case 'rightArrow':
      case 'leftArrow':
      case 'upArrow':
      case 'downArrow': {
        const t = 0.4;
        ctx.beginPath();
        if (prst === 'rightArrow') {
          ctx.moveTo(x, y + h * 0.3); ctx.lineTo(x + w * 0.6, y + h * 0.3);
          ctx.lineTo(x + w * 0.6, y); ctx.lineTo(x + w, y + h2);
          ctx.lineTo(x + w * 0.6, y + h); ctx.lineTo(x + w * 0.6, y + h * 0.7);
          ctx.lineTo(x, y + h * 0.7);
        } else if (prst === 'leftArrow') {
          ctx.moveTo(x + w, y + h * 0.3); ctx.lineTo(x + w * 0.4, y + h * 0.3);
          ctx.lineTo(x + w * 0.4, y); ctx.lineTo(x, y + h2);
          ctx.lineTo(x + w * 0.4, y + h); ctx.lineTo(x + w * 0.4, y + h * 0.7);
          ctx.lineTo(x + w, y + h * 0.7);
        } else if (prst === 'upArrow') {
          ctx.moveTo(x + w * 0.3, y + h); ctx.lineTo(x + w * 0.3, y + h * 0.4);
          ctx.lineTo(x, y + h * 0.4); ctx.lineTo(x + w2, y);
          ctx.lineTo(x + w, y + h * 0.4); ctx.lineTo(x + w * 0.7, y + h * 0.4);
          ctx.lineTo(x + w * 0.7, y + h);
        } else {
          ctx.moveTo(x + w * 0.3, y); ctx.lineTo(x + w * 0.3, y + h * 0.6);
          ctx.lineTo(x, y + h * 0.6); ctx.lineTo(x + w2, y + h);
          ctx.lineTo(x + w, y + h * 0.6); ctx.lineTo(x + w * 0.7, y + h * 0.6);
          ctx.lineTo(x + w * 0.7, y);
        }
        ctx.closePath();
        return;
      }
      case 'chevron': {
        ctx.beginPath();
        ctx.moveTo(x, y); ctx.lineTo(x + w * 0.75, y); ctx.lineTo(x + w, y + h2);
        ctx.lineTo(x + w * 0.75, y + h); ctx.lineTo(x, y + h); ctx.lineTo(x + w * 0.25, y + h2);
        ctx.closePath();
        return;
      }
      case 'parallelogram': {
        const o = w * 0.2;
        ctx.beginPath();
        ctx.moveTo(x + o, y); ctx.lineTo(x + w, y); ctx.lineTo(x + w - o, y + h); ctx.lineTo(x, y + h);
        ctx.closePath();
        return;
      }
      case 'trapezoid': {
        const o = w * 0.2;
        ctx.beginPath();
        ctx.moveTo(x + o, y); ctx.lineTo(x + w - o, y); ctx.lineTo(x + w, y + h); ctx.lineTo(x, y + h);
        ctx.closePath();
        return;
      }
      case 'line':
      case 'straightConnector1':
        return; // 单独处理
      default:
        ctx.beginPath();
        ctx.rect(x, y, w, h);
    }
  }

  function hasPlaceholder(sp) {
    const nv = kid(sp, 'nvSpPr');
    const nvPr = nv && kid(nv, 'nvPr');
    return !!(nvPr && kid(nvPr, 'ph'));
  }

  /* 自定义几何（任意形状）：按 a:path 指令绘制 */
  function pathForCustom(ctx, cust, x, y, w, h) {
    const pathLst = kid(cust, 'pathLst');
    if (!pathLst) return false;
    const paths = kids(pathLst, 'path');
    if (!paths.length) return false;
    ctx.beginPath();
    let any = false;
    for (const path of paths) {
      const pw = numAttr(path, 'w', w) || w || 1;
      const ph = numAttr(path, 'h', h) || h || 1;
      const sx = w / pw, sy = h / ph;
      for (const node of path.children) {
        const ln = node.localName;
        if (ln === 'moveTo' || ln === 'lnTo') {
          const pt = kid(node, 'pt');
          if (!pt) continue;
          const px = x + numAttr(pt, 'x', 0) * sx;
          const py = y + numAttr(pt, 'y', 0) * sy;
          if (ln === 'moveTo') ctx.moveTo(px, py); else ctx.lineTo(px, py);
          any = true;
        } else if (ln === 'cubicBezTo' || ln === 'quadBezTo') {
          const pts = kids(node, 'pt');
          if (ln === 'cubicBezTo' && pts.length >= 3) {
            ctx.bezierCurveTo(
              x + numAttr(pts[0], 'x', 0) * sx, y + numAttr(pts[0], 'y', 0) * sy,
              x + numAttr(pts[1], 'x', 0) * sx, y + numAttr(pts[1], 'y', 0) * sy,
              x + numAttr(pts[2], 'x', 0) * sx, y + numAttr(pts[2], 'y', 0) * sy);
            any = true;
          } else if (pts.length >= 2) {
            ctx.quadraticCurveTo(
              x + numAttr(pts[0], 'x', 0) * sx, y + numAttr(pts[0], 'y', 0) * sy,
              x + numAttr(pts[1], 'x', 0) * sx, y + numAttr(pts[1], 'y', 0) * sy);
            any = true;
          }
        } else if (ln === 'close') {
          ctx.closePath();
        }
      }
    }
    return any;
  }

  /* ----------------------------- 渲染器 ----------------------------- */
  function Deck(zip, cx, cy, slideParts, theme) {
    this.zip = zip;
    this.cx = cx;
    this.cy = cy;
    this.slideParts = slideParts;
    this.theme = theme;
    this._slides = [];
    this._imageCache = new Map(); // path -> Promise<HTMLImageElement|null>
    this._urls = [];
  }
  Deck.prototype.slideCount = function () { return this.slideParts.length; };

  Deck.prototype._readPart = async function (name) {
    if (!this.zip.has(name)) return null;
    return parseXml(await this.zip.bytes(name));
  };

  Deck.prototype._loadPlaceholders = function (doc) {
    const map = Object.create(null);
    if (!doc) return map;
    for (const sp of localAll(doc, 'sp')) {
      const nv = kid(sp, 'nvSpPr');
      const nvPr = nv && kid(nv, 'nvPr');
      const ph = nvPr && kid(nvPr, 'ph');
      if (!ph) continue;
      const type = ph.getAttribute('type') || 'obj';
      const idx = ph.getAttribute('idx') || '0';
      const xf = xfrmOf(sp);
      if (!xf) continue;
      map['t:' + type] = map['t:' + type] || xf;
      map['i:' + idx] = map['i:' + idx] || xf;
    }
    return map;
  };

  Deck.prototype._slide = async function (i) {
    if (this._slides[i]) return this._slides[i];
    const part = this.slideParts[i];
    const doc = await this._readPart(part);
    if (!doc) throw new Error('第 ' + (i + 1) + ' 页数据缺失');
    const rels = await readRels(this.zip, part);
    let layoutPart = null, masterPart = null;
    for (const id in rels) {
      if (/\/slideLayout$/.test(rels[id].type)) layoutPart = resolvePath(part, rels[id].target);
    }
    if (layoutPart) {
      const lrels = await readRels(this.zip, layoutPart);
      for (const id in lrels) {
        if (/\/slideMaster$/.test(lrels[id].type)) masterPart = resolvePath(layoutPart, lrels[id].target);
      }
    }
    const layoutDoc = layoutPart ? await this._readPart(layoutPart) : null;
    const masterDoc = masterPart ? await this._readPart(masterPart) : null;
    const layoutRels = layoutPart ? await readRels(this.zip, layoutPart) : {};
    const masterRels = masterPart ? await readRels(this.zip, masterPart) : {};
    const slide = {
      doc, part, rels,
      layout: { doc: layoutDoc, part: layoutPart, rels: layoutRels, ph: this._loadPlaceholders(layoutDoc) },
      master: { doc: masterDoc, part: masterPart, rels: masterRels, ph: this._loadPlaceholders(masterDoc) },
      images: null
    };
    this._slides[i] = slide;
    return slide;
  };

  Deck.prototype._imageFor = function (part, rels, embedId) {
    const rel = rels[embedId];
    if (!rel) return Promise.resolve(null);
    const path = resolvePath(part, rel.target);
    if (this._imageCache.has(path)) return this._imageCache.get(path);
    const p = (async () => {
      if (!this.zip.has(path)) return null;
      const bytes = await this.zip.bytes(path);
      const ext = (path.split('.').pop() || '').toLowerCase();
      const mime = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg'
        : ext === 'gif' ? 'image/gif' : ext === 'bmp' ? 'image/bmp'
        : ext === 'svg' ? 'image/svg+xml' : 'application/octet-stream';
      const blob = new Blob([bytes], { type: mime });
      const url = URL.createObjectURL(blob);
      this._urls.push(url);
      const img = new Image();
      const ok = await new Promise(resolve => {
        let done = false;
        const fin = (v) => { if (done) return; done = true; resolve(v); };
        img.onload = () => fin(true);
        img.onerror = () => fin(false);
        setTimeout(() => fin(false), 8000);
        img.src = url;
        if (img.complete && (img.naturalWidth || img.width)) fin(true);
      });
      return ok ? img : null;
    })();
    this._imageCache.set(path, p);
    return p;
  };

  Deck.prototype._collectImages = async function (slide) {
    if (slide.images) return;
    slide.images = Object.create(null);
    const parts = [
      { part: slide.part, rels: slide.rels, doc: slide.doc },
      { part: slide.layout.part, rels: slide.layout.rels, doc: slide.layout.doc },
      { part: slide.master.part, rels: slide.master.rels, doc: slide.master.doc }
    ];
    for (const P of parts) {
      if (!P.part || !P.doc) continue;
      const map = new Map();
      for (const blip of localAll(P.doc, 'blip')) {
        const id = rid(blip);
        if (!id || map.has(id)) continue;
        const img = await this._imageFor(P.part, P.rels, id);
        if (img) map.set(id, img);
      }
      slide.images[P.part] = map;
    }
  };

  Deck.prototype._bgSource = function (slide) {
    const pick = (doc, part) => {
      if (!doc) return null;
      const cSld = kid(doc.documentElement, 'cSld');
      const bg = cSld && kid(cSld, 'bg');
      return bg ? { bg: bg, part: part } : null;
    };
    return pick(slide.doc, slide.part) ||
      pick(slide.layout.doc, slide.layout.part) ||
      pick(slide.master.doc, slide.master.part);
  };

  Deck.prototype._paintBackground = async function (ctx, slide, k, cw, ch) {
    const src = this._bgSource(slide);
    const bg = src && src.bg;
    const owner = src && src.part;
    const imgMap = (owner && slide.images && slide.images[owner]) || null;
    let filled = false;
    if (bg) {
      const bgPr = kid(bg, 'bgPr') || kid(bg, 'bgRef');
      const solid = bgPr && kid(bgPr, 'solidFill');
      if (solid) {
        const color = colorOf(solid, this.theme, null);
        if (color) { ctx.fillStyle = color; ctx.fillRect(0, 0, cw, ch); filled = true; }
      } else if (bgPr) {
        const grad = kid(bgPr, 'gradFill');
        if (grad) {
          const gs = localAll(grad, 'gs');
          const stops = gs.map(g => ({ c: colorOf(g, this.theme, null), p: numAttr(g, 'pos', 0) / 1000 })).filter(s => s.c);
          if (stops.length) {
            const grad2 = ctx.createLinearGradient(0, 0, cw, ch);
            stops.forEach(s => grad2.addColorStop(Math.min(1, Math.max(0, s.p / 100)), s.c));
            ctx.fillStyle = grad2;
            ctx.fillRect(0, 0, cw, ch); filled = true;
          }
        } else {
          const blipFill = kid(bgPr, 'blipFill');
          const blip = blipFill && localAll(blipFill, 'blip')[0];
          if (blip) {
            const img = imgMap && imgMap.get(rid(blip));
            if (img) { ctx.drawImage(img, 0, 0, cw, ch); filled = true; }
          }
        }
      }
    }
    if (!filled) {
      const lt1 = this.theme.colors.lt1;
      if (lt1 && lt1.toLowerCase() !== '#ffffff') {
        ctx.fillStyle = lt1; ctx.fillRect(0, 0, cw, ch);
      }
    }
  };

  Deck.prototype._xfrmFor = function (sp, slide) {
    const own = xfrmOf(sp);
    if (own) return own;
    const nv = kid(sp, 'nvSpPr');
    const nvPr = nv && kid(nv, 'nvPr');
    const ph = nvPr && kid(nvPr, 'ph');
    if (!ph) return null;
    const type = ph.getAttribute('type') || 'obj';
    const idx = ph.getAttribute('idx') || '0';
    for (const src of [slide.layout, slide.master]) {
      if (src.ph['i:' + idx]) return src.ph['i:' + idx];
      if (src.ph['t:' + type]) return src.ph['t:' + type];
    }
    return null;
  };

  Deck.prototype._shapeText = function (ctx, body, rect, k, shape) {
    const pxPerPt = k * EMU_PER_PT;
    const bodyPr = kid(body, 'bodyPr');
    const lIns = numAttr(bodyPr, 'lIns', 91440) * k;
    const tIns = numAttr(bodyPr, 'tIns', 45720) * k;
    const rIns = numAttr(bodyPr, 'rIns', 91440) * k;
    const bIns = numAttr(bodyPr, 'bIns', 45720) * k;
    const anchor = (bodyPr && bodyPr.getAttribute('anchor')) || 't';
    const wrap = !bodyPr || bodyPr.getAttribute('wrap') !== 'none';
    const contentW = Math.max(4, rect.w - lIns - rIns);
    const paras = kids(body, 'p');
    const lines = [];

    const style = kid(shape, 'style');
    const fontRef = style && kid(style, 'fontRef');
    const styleColor = fontRef ? colorOf(fontRef.firstElementChild, this.theme, null) : null;

    for (const p of paras) {
      const pPr = kid(p, 'pPr');
      const algn = (pPr && pPr.getAttribute('algn')) || 'l';
      const marL = numAttr(pPr, 'marL', 0) * k;
      const indent = numAttr(pPr, 'indent', 0) * k;
      const lnSpcEl = pPr && kid(pPr, 'lnSpc');
      const lnSpcPct = lnSpcEl && kid(lnSpcEl, 'spcPct');
      const lnMul = lnSpcPct ? numAttr(lnSpcPct, 'val', 100000) / 100000 : 1;
      const buChar = pPr && kid(pPr, 'buChar');
      const buAuto = pPr && kid(pPr, 'buAutoNum');
      const buNone = pPr && kid(pPr, 'buNone');
      const bullet = (!buNone && (buChar || buAuto)) ? (buChar ? buChar.getAttribute('char') : '•') : '';

      const defRPr = (pPr && kid(pPr, 'defRPr'));
      const runs = [];
      for (const r of kids(p, 'r')) {
        const rPr = kid(r, 'rPr');
        const t = kid(r, 't');
        runs.push({
          text: t ? t.textContent : '',
          size: numAttr(rPr, 'sz', numAttr(defRPr, 'sz', 1800)),
          bold: (rPr && (rPr.getAttribute('b') === '1')) || (defRPr && defRPr.getAttribute('b') === '1'),
          italic: (rPr && rPr.getAttribute('i') === '1') || (defRPr && defRPr.getAttribute('i') === '1'),
          underline: rPr && rPr.getAttribute('u') && rPr.getAttribute('u') !== 'none',
          color: (rPr && colorOf(kid(rPr, 'solidFill'), this.theme, null)) ||
                 (defRPr && colorOf(kid(defRPr, 'solidFill'), this.theme, null)) ||
                 styleColor || '#000000',
          font: this._fontOf(rPr, defRPr)
        });
      }
      const brs = kids(p, 'br');
      if (brs.length) {
        // 按 br 拆分简单处理
        const parts = [];
        let acc = '';
        for (const n of p.children) {
          if (n.localName === 'r') acc += (kid(n, 't') ? kid(n, 't').textContent : '');
          else if (n.localName === 'br') { parts.push(acc); acc = ''; }
        }
        parts.push(acc);
        for (let i = 0; i < parts.length; i++) {
          this._layoutParagraph(ctx, parts[i], runs[0] || { size: 1800, color: '#000', font: '' }, pxPerPt, contentW, wrap, lines, algn, marL, indent, lnMul, i === 0 ? bullet : '');
        }
      } else {
        const text = runs.map(r => r.text).join('');
        const first = runs[0] || { size: 1800, color: '#000', font: '', bold: false, italic: false, underline: false };
        if (!text) { lines.push({ blank: true, size: first.size, color: first.color, font: first.font, algn, marL, lnMul }); continue; }
        // 混排：按 run 逐字拆分
        const chars = [];
        for (const r of runs) {
          const arr = Array.from(r.text);
          for (const ch of arr) chars.push({ ch, run: r });
        }
        this._layoutRuns(ctx, chars, pxPerPt, contentW, wrap, lines, algn, marL, indent, lnMul, bullet);
      }
    }

    if (!lines.length) return;
    let totalH = 0;
    for (const line of lines) {
      line.h = line.blank ? line.size / 100 * pxPerPt * 1.2 * line.lnMul : line.size / 100 * pxPerPt * 1.25 * line.lnMul;
      totalH += line.h;
    }
    let y = rect.y + tIns;
    if (anchor === 'ctr') y = rect.y + (rect.h - totalH) / 2;
    else if (anchor === 'b') y = rect.y + rect.h - bIns - totalH;

    for (const line of lines) {
      if (line.blank) { y += line.h; continue; }
      ctx.font = (line.italic ? 'italic ' : '') + (line.bold ? '700 ' : '') + (line.size / 100 * pxPerPt) + 'px ' + (line.font || '"Microsoft YaHei","微软雅黑",sans-serif');
      ctx.fillStyle = line.color;
      ctx.textBaseline = 'alphabetic';
      let x = rect.x + lIns + line.marL;
      if (line.bullet) {
        ctx.fillText(line.bullet, x + line.indent, y + line.h * 0.78);
      }
      let tx = x + (line.bullet ? Math.max(0, line.indent) : 0);
      if (line.algn === 'ctr') tx = rect.x + lIns + (contentW - line.width) / 2;
      else if (line.algn === 'r') tx = rect.x + lIns + contentW - line.width;
      ctx.fillText(line.text, tx, y + line.h * 0.78);
      y += line.h;
    }
  };

  Deck.prototype._fontOf = function (rPr, defRPr) {
    const get = (el) => {
      if (!el) return '';
      const latin = kid(el, 'latin');
      const ea = kid(el, 'ea');
      return (ea && ea.getAttribute('typeface')) || (latin && latin.getAttribute('typeface')) || '';
    };
    return get(rPr) || get(defRPr) || '';
  };

  Deck.prototype._layoutParagraph = function (ctx, text, run, pxPerPt, width, wrap, lines, algn, marL, indent, lnMul, bullet) {
    ctx.font = (run.italic ? 'italic ' : '') + (run.bold ? '700 ' : '') + (run.size / 100 * pxPerPt) + 'px ' + (run.font || '"Microsoft YaHei","微软雅黑",sans-serif');
    const words = wrap ? String(text).match(/[\u4e00-\u9fff]|[^\s\u4e00-\u9fff]+|\s+/g) || [] : [String(text)];
    let cur = '', curW = 0;
    const push = (s, w) => {
      lines.push({ text: s, width: w, size: run.size, color: run.color, font: run.font, bold: run.bold, italic: run.italic, algn, marL, indent, lnMul, bullet: bullet || '' });
      bullet = '';
    };
    for (const word of words) {
      const w = ctx.measureText(word).width;
      if (cur && curW + w > width) { push(cur, curW); cur = ''; curW = 0; }
      if (!wrap || w <= width || !cur) { cur += word; curW += w; }
      else { push(word, w); }
    }
    if (cur) push(cur, curW);
    if (!words.length) push('', 0);
  };

  Deck.prototype._layoutRuns = function (ctx, chars, pxPerPt, width, wrap, lines, algn, marL, indent, lnMul, bullet) {
    let cur = '', curW = 0, maxSize = 0, color = '#000', font = '', bold = false, italic = false;
    const push = () => {
      lines.push({ text: cur, width: curW, size: maxSize || 1800, color, font, bold, italic, algn, marL, indent, lnMul, bullet: bullet || '' });
      bullet = '';
      cur = ''; curW = 0;
    };
    for (const c of chars) {
      const r = c.run;
      ctx.font = (r.italic ? 'italic ' : '') + (r.bold ? '700 ' : '') + (r.size / 100 * pxPerPt) + 'px ' + (r.font || '"Microsoft YaHei","微软雅黑",sans-serif');
      const w = ctx.measureText(c.ch).width;
      if (wrap && cur && curW + w > width) push();
      cur += c.ch; curW += w;
      if (r.size > maxSize) maxSize = r.size;
      color = r.color; font = r.font; bold = r.bold; italic = r.italic;
    }
    if (cur) push();
  };

  Deck.prototype._paintPicture = async function (ctx, pic, slide, k, part) {
    const xf = xfrmOf(pic);
    if (!xf) return;
    const blip = localAll(pic, 'blip')[0];
    if (!blip) return;
    const map = (part && slide.images && slide.images[part]) || null;
    const img = map && map.get(rid(blip));
    if (!img) return;
    const x = xf.x * k, y = xf.y * k, w = xf.cx * k, h = xf.cy * k;
    let sx = 0, sy = 0, sw = img.naturalWidth || img.width, sh = img.naturalHeight || img.height;
    const blipFill = kid(pic, 'blipFill');
    const srcRect = blipFill && kid(blipFill, 'srcRect');
    if (srcRect) {
      const l = numAttr(srcRect, 'l', 0) / 100000, t = numAttr(srcRect, 't', 0) / 100000;
      const r = numAttr(srcRect, 'r', 0) / 100000, b = numAttr(srcRect, 'b', 0) / 100000;
      sx = sw * l; sy = sh * t; sw = sw * (1 - l - r); sh = sh * (1 - t - b);
    }
    ctx.save();
    if (xf.rot) {
      ctx.translate(x + w / 2, y + h / 2); ctx.rotate(xf.rot / 60000 * Math.PI / 180);
      ctx.translate(-(x + w / 2), -(y + h / 2));
    }
    try { ctx.drawImage(img, sx, sy, sw, sh, x, y, w, h); } catch (e) {}
    ctx.restore();
  };

  Deck.prototype._paintSp = function (ctx, sp, slide, k) {
    const xf = this._xfrmFor(sp, slide);
    if (!xf) return;
    const x = xf.x * k, y = xf.y * k, w = Math.abs(xf.cx) * k, h = Math.abs(xf.cy) * k;
    const prst = presetOf(sp);
    const fill = fillOf(sp, this.theme);
    const line = lineOf(sp, this.theme);

    ctx.save();
    if (xf.rot) {
      ctx.translate(x + w / 2, y + h / 2); ctx.rotate(xf.rot / 60000 * Math.PI / 180);
      ctx.translate(-(x + w / 2), -(y + h / 2));
    }
    if (prst === 'line' || prst === 'straightConnector1') {
      if (line) {
        let x1 = x, y1 = y, x2 = x + w, y2 = y + h;
        if (xf.flipH) { x1 = x + w; x2 = x; }
        if (xf.flipV) { y1 = y + h; y2 = y; }
        ctx.strokeStyle = line.color;
        ctx.lineWidth = Math.max(1, line.width * k);
        ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke();
      }
    } else {
      const spPr = kid(sp, 'spPr');
      const cust = spPr && kid(spPr, 'custGeom');
      let ok = true;
      if (cust) ok = pathForCustom(ctx, cust, x, y, w, h);
      else pathFor(ctx, prst, x, y, w, h);
      if (ok) {
        if (fill) { ctx.fillStyle = fill; ctx.fill(); }
        if (line) {
          ctx.strokeStyle = line.color;
          ctx.lineWidth = Math.max(1, line.width * k);
          ctx.stroke();
        }
      }
    }
    ctx.restore();

    const body = kid(sp, 'txBody');
    if (body) this._shapeText(ctx, body, { x, y, w, h }, k, sp);
  };

  Deck.prototype._paintTable = function (ctx, frame, slide, k) {
    const tbl = localAll(frame, 'tbl')[0];
    if (!tbl) return;
    const xf = xfrmOf(frame);
    const x0 = xf ? xf.x * k : 0, y0 = xf ? xf.y * k : 0;
    const cols = kids(tbl, 'tblGrid').flatMap(g => kids(g, 'gridCol').map(c => numAttr(c, 'w', 0) * k));
    let y = y0;
    for (const tr of kids(tbl, 'tr')) {
      const rowH = numAttr(tr, 'h', 0) * k;
      let x = x0;
      const cells = kids(tr, 'tc');
      for (let i = 0; i < cells.length; i++) {
        const cw = cols[i] || 0;
        const tc = cells[i];
        const tcPr = kid(tc, 'tcPr');
        const fill = tcPr && colorOf(kid(tcPr, 'solidFill'), this.theme, null);
        if (fill) { ctx.fillStyle = fill; ctx.fillRect(x, y, cw, rowH); }
        ctx.strokeStyle = 'rgba(0,0,0,.25)';
        ctx.lineWidth = Math.max(1, 12700 * k);
        ctx.strokeRect(x, y, cw, rowH);
        const body = kid(tc, 'txBody');
        if (body) this._shapeText(ctx, body, { x, y, w: cw, h: rowH }, k, tc);
        x += cw;
      }
      y += rowH;
    }
  };

  Deck.prototype._paintTree = async function (ctx, container, slide, k, part, skipPh) {
    for (let n = container.firstElementChild; n; n = n.nextElementSibling) {
      const name = n.localName;
      if (name === 'sp' || name === 'cxnSp') {
        if (skipPh && hasPlaceholder(n)) continue;
        this._paintSp(ctx, n, slide, k);
      } else if (name === 'pic') {
        await this._paintPicture(ctx, n, slide, k, part);
      } else if (name === 'grpSp') {
        const gp = kid(n, 'grpSpPr');
        const gx = gp && kid(gp, 'xfrm');
        if (!gx) { await this._paintTree(ctx, n, slide, k, part, skipPh); continue; }
        const off = kid(gx, 'off'), ext = kid(gx, 'ext');
        const chOff = kid(gx, 'chOff'), chExt = kid(gx, 'chExt');
        if (!off || !ext || !chOff || !chExt || !numAttr(chExt, 'cx', 0) || !numAttr(chExt, 'cy', 0)) {
          await this._paintTree(ctx, n, slide, k, part, skipPh);
          continue;
        }
        ctx.save();
        ctx.translate(numAttr(off, 'x', 0) * k, numAttr(off, 'y', 0) * k);
        ctx.scale(numAttr(ext, 'cx', 0) / numAttr(chExt, 'cx', 1), numAttr(ext, 'cy', 0) / numAttr(chExt, 'cy', 1));
        ctx.translate(-numAttr(chOff, 'x', 0) * k, -numAttr(chOff, 'y', 0) * k);
        await this._paintTree(ctx, n, slide, k, part, skipPh);
        ctx.restore();
      } else if (name === 'graphicFrame') {
        if (localAll(n, 'tbl').length) this._paintTable(ctx, n, slide, k);
      }
    }
  };

  Deck.prototype.renderSlideCanvas = async function (index, targetWidth) {
    const slide = await this._slide(index);
    await this._collectImages(slide);
    const k = Math.max(0.00001, targetWidth / this.cx);
    const cw = Math.max(1, Math.round(this.cx * k));
    const ch = Math.max(1, Math.round(this.cy * k));
    const cv = document.createElement('canvas');
    cv.width = cw; cv.height = ch;
    const ctx = cv.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, cw, ch);
    await this._paintBackground(ctx, slide, k, cw, ch);
    const showMaster = slide.doc.documentElement.getAttribute('showMasterSp') !== '0';
    if (showMaster) {
      if (slide.master.doc) {
        const mc = kid(slide.master.doc.documentElement, 'cSld');
        const mt = mc && kid(mc, 'spTree');
        if (mt) await this._paintTree(ctx, mt, slide, k, slide.master.part, true);
      }
      if (slide.layout.doc) {
        const lc = kid(slide.layout.doc.documentElement, 'cSld');
        const lt = lc && kid(lc, 'spTree');
        if (lt) await this._paintTree(ctx, lt, slide, k, slide.layout.part, true);
      }
    }
    const cSld = kid(slide.doc.documentElement, 'cSld');
    const tree = cSld && kid(cSld, 'spTree');
    if (tree) await this._paintTree(ctx, tree, slide, k, slide.part, false);
    return cv;
  };

  Deck.prototype.dispose = function () {
    for (const u of this._urls) { try { URL.revokeObjectURL(u); } catch (e) {} }
    this._urls = [];
    this._imageCache.clear();
    this._slides = [];
  };

  /* ----------------------------- 入口 ----------------------------- */
  async function load(arrayBuffer) {
    const zip = new Zip(arrayBuffer);
    if (!zip.has('ppt/presentation.xml')) {
      throw new Error('不是有效的 .pptx 文件（旧版 .ppt 请先在 PowerPoint 中另存为 .pptx）');
    }
    const pres = parseXml(await zip.bytes('ppt/presentation.xml'));
    const rels = await readRels(zip, 'ppt/presentation.xml');
    const sldSz = localAll(pres, 'sldSz')[0];
    const cx = numAttr(sldSz, 'cx', 12192000);
    const cy = numAttr(sldSz, 'cy', 6858000);

    let slideParts = [];
    const idLst = localAll(pres, 'sldIdLst')[0];
    if (idLst) {
      for (const sldId of kids(idLst, 'sldId')) {
        const r = sldId.getAttributeNS(REL_NS, 'id');
        const rel = rels[r];
        if (rel) slideParts.push(resolvePath('ppt/presentation.xml', rel.target));
      }
    }
    if (!slideParts.length) {
      slideParts = zip.names().filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
        .sort((a, b) => parseInt(a.match(/(\d+)\.xml$/)[1], 10) - parseInt(b.match(/(\d+)\.xml$/)[1], 10));
    }
    if (!slideParts.length) throw new Error('PPT 中没有找到幻灯片');

    let themeDoc = null;
    const themePath = zip.names().find(n => /^ppt\/theme\/theme1\.xml$/.test(n));
    if (themePath) themeDoc = parseXml(await zip.bytes(themePath));
    const theme = parseTheme(themeDoc);

    const deck = new Deck(zip, cx, cy, slideParts, theme);
    // 预热首页
    await deck._slide(0);
    return deck;
  }

  global.PPTX = { load: load };
})(window);
