/* doc_to_docx.js: run by tools/doc_to_docx.py in Chrome on a built kit document, in print media at the width of the
   page's text. PC_DOCX() walks the rendered page and returns what Word needs, lengths in CSS px: paragraphs with runs
   and computed styles, tables, boxes, blocks side by side, pictures to capture, page breaks, named pages, the cover,
   the contents and the page set-up. */
window.PC_DOCX = function (mode) {
  'use strict';
  var NOTE = '.pc-footnote, .pc-sidenote, .pc-fn';
  var PIC = 'img, svg, canvas, video, picture, iframe, object, .pc-chart, .pc-media, .pc-flow, .pc-qr, .pc-map, [data-docx="picture"]';
  var cv = document.createElement('canvas'), g, seen = {}, pics = [], fonts = {}, counters = {}, lists = new WeakMap(), nlist = 0;
  cv.width = cv.height = 1;
  g = cv.getContext('2d', {willReadFrequently: true});
  var rootCS = getComputedStyle(document.documentElement);
  var docEl = document.body, docCS = getComputedStyle(docEl), htmlEl = document.documentElement;

  function rgba(c) {
    if (!c || c === 'transparent' || c === 'none') return null;
    if (c in seen) return seen[c];
    g.clearRect(0, 0, 1, 1); g.fillStyle = '#000'; g.fillStyle = c; g.fillRect(0, 0, 1, 1);
    var d = g.getImageData(0, 0, 1, 1).data;
    return (seen[c] = d[3] ? [d[0], d[1], d[2], d[3] / 255] : null);
  }
  function over(a, b) {
    if (!a) return b;
    if (a[3] >= 0.999 || !b) return a;
    return [0, 1, 2].map(function (i) { return a[i] * a[3] + b[i] * (1 - a[3]); }).concat(1);
  }
  function hex(v) { return v ? v.slice(0, 3).map(function (n) { return ('0' + Math.round(n).toString(16)).slice(-2); }).join('') : null; }
  function num(v) { return parseFloat(v) || 0; }
  function family(ff) { return ff.split(',')[0].trim().replace(/^["']|["']$/g, ''); }
  function weight(w) { return w === 'bold' ? 700 : w === 'normal' ? 400 : parseInt(w, 10) || 400; }
  function rect(el) { var r = el.getBoundingClientRect(); return {x: r.left + scrollX, y: r.top + scrollY, w: r.width, h: r.height}; }
  function hidden(el, cs) {
    return cs.display === 'none' || cs.visibility === 'hidden' || cs.visibility === 'collapse' || cs.opacity === '0' ||
      (el.dataset && el.dataset.docx === 'skip') || el.matches('.pcx-src, template, script, style, noscript, link, meta');
  }
  function side(cs, s) {
    var w = num(cs['border' + s + 'Width']), st = cs['border' + s + 'Style'];
    return w > 0 && st !== 'none' && st !== 'hidden' ? {w: w, st: st, c: hex(rgba(cs['border' + s + 'Color']))} : null;
  }
  function borders(cs) { return {t: side(cs, 'Top'), r: side(cs, 'Right'), b: side(cs, 'Bottom'), l: side(cs, 'Left')}; }
  function pads(cs) { return [num(cs.paddingTop), num(cs.paddingRight), num(cs.paddingBottom), num(cs.paddingLeft)]; }
  function textIn(el) { return /\S/.test(el.textContent || ''); }

  // ---- CSS counters, so ::before and ::after text reads as Chrome draws it
  function pairs(v, d) {
    var out = [];
    if (!v || v === 'none') return out;
    var t = v.trim().split(/\s+/);
    for (var i = 0; i < t.length; i++) {
      var k = parseInt(t[i + 1], 10);
      if (!isNaN(k)) { out.push([t[i], k]); i++; } else out.push([t[i], d]);
    }
    return out;
  }
  function enter(cs, frame) {
    pairs(cs.counterReset, 0).forEach(function (p) {
      var s = counters[p[0]] = counters[p[0]] || [];
      if (frame.indexOf(p[0]) >= 0) s[s.length - 1] = p[1]; else { s.push(p[1]); frame.push(p[0]); }
    });
    pairs(cs.counterSet, 0).forEach(function (p) { var s = counters[p[0]] = counters[p[0]] || [0]; s[s.length - 1] = p[1]; });
    pairs(cs.counterIncrement, 1).forEach(function (p) { var s = counters[p[0]] = counters[p[0]] || [0]; s[s.length - 1] += p[1]; });
  }
  function leave(frame) { frame.forEach(function (n) { counters[n].pop(); }); }
  function roman(n) {
    var v = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1], s = ['m', 'cm', 'd', 'cd', 'c', 'xc', 'l', 'xl', 'x', 'ix', 'v', 'iv', 'i'], o = '';
    for (var i = 0; i < v.length; i++) while (n >= v[i]) { o += s[i]; n -= v[i]; }
    return o;
  }
  function alpha(n) { var o = ''; while (n > 0) { n--; o = String.fromCharCode(97 + n % 26) + o; n = Math.floor(n / 26); } return o; }
  function fmt(n, st) {
    st = (st || 'decimal').trim();
    if (st === 'none') return '';
    if (/roman/.test(st)) return st[0] === 'u' ? roman(n).toUpperCase() : roman(n);
    if (/alpha|latin/.test(st)) return st[0] === 'u' ? alpha(n).toUpperCase() : alpha(n);
    if (st === 'decimal-leading-zero') return (n < 10 ? '0' : '') + n;
    return String(n);
  }
  function unesc(s) {
    return s.replace(/\\([0-9a-fA-F]{1,6})\s?/g, function (m, h) { return String.fromCodePoint(parseInt(h, 16)); }).replace(/\\(.)/g, '$1');
  }
  function content(el, pcs) {
    var c = pcs.content, out = '', m;
    if (!c || c === 'none' || c === 'normal') return '';
    var re = /"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)'|(counters?)\(([^)]*)\)|attr\(([^)]*)\)|(open-quote|close-quote)/g;
    while ((m = re.exec(c))) {
      if (m[1] !== undefined || m[2] !== undefined) out += unesc(m[1] !== undefined ? m[1] : m[2]);
      else if (m[3]) {
        var a = m[4].split(',').map(function (s) { return s.trim(); }), s = counters[a[0]] || [0];
        if (m[3] === 'counter') out += fmt(s[s.length - 1] || 0, a[1]);
        else out += s.map(function (n) { return fmt(n, a[2]); }).join(unesc((a[1] || '""').replace(/^["']|["']$/g, '')));
      } else if (m[5]) out += el.getAttribute(m[5].trim()) || '';
      else if (m[6]) out += m[6] === 'open-quote' ? '“' : '”';
    }
    return out;
  }

  // ---- runs
  function run(t, cs, ctx, el) {
    var r = {t: t, f: family(cs.fontFamily), w: weight(cs.fontWeight), sz: num(cs.fontSize), c: hex(over(rgba(cs.color), ctx.bg))};
    if (cs.fontStyle !== 'normal') r.i = 1;
    if (cs.textTransform === 'uppercase') r.caps = 1;
    if (/small-caps/.test(cs.fontVariantCaps || '')) r.sc = 1;
    var d = cs.textDecorationLine || '';
    if (/underline/.test(d)) r.u = 1;
    if (/line-through/.test(d)) r.s = 1;
    if (num(cs.letterSpacing)) r.ls = num(cs.letterSpacing);
    if (cs.verticalAlign === 'super' || cs.verticalAlign === 'sub') {
      r.v = cs.verticalAlign;
      r.sz = num(getComputedStyle(el.parentElement || el).fontSize) || r.sz;
    }
    if (/^inline/.test(cs.display) && rgba(cs.backgroundColor)) r.bg = hex(over(rgba(cs.backgroundColor), ctx.bg));
    var a = el.closest && el.closest('a[href]');
    if (a) r.href = a.getAttribute('href');
    var l = el.closest && el.closest('[lang]');
    if (l && l.lang) r.lang = l.lang;
    if (/[\u0D80-\u0DFF\u0B80-\u0BFF]/.test(t)) r.ff = cs.fontFamily;
    var fr = el.closest && el.closest('.pc-fnref');
    if (fr && !r.v) { r.v = 'super'; r.fnref = 1; r.sz = num(getComputedStyle(fr.parentElement || fr).fontSize) || r.sz; }
    fonts[r.f + '|' + r.w + '|' + (r.i ? 1 : 0)] = 1;
    return r;
  }
  function textOf(node, cs) {
    var t = node.data;
    if (!/pre/.test(cs.whiteSpace)) t = t.replace(/[\t\n\r ]+/g, ' ');
    else if (cs.whiteSpace === 'pre-line') t = t.replace(/[\t ]+/g, ' ');
    if (cs.textTransform === 'lowercase') t = t.toLowerCase();
    else if (cs.textTransform === 'capitalize') t = t.replace(/(^|\s)(\S)/g, function (m, a, b) { return a + b.toUpperCase(); });
    return t;
  }
  function pseudo(el, which, ctx, p) {
    var pcs = getComputedStyle(el, which);
    if (pcs.display === 'none') return;
    var t = content(el, pcs), abs = /absolute|fixed/.test(pcs.position), r = null;
    if (t && /\S/.test(t)) r = run(t.replace(/[\t\n\r ]+/g, ' '), pcs, ctx, el);
    else if (t === '' && which === '::before' && ctx.bullet && rgba(pcs.backgroundColor)) {
      r = run('• ', pcs, ctx, el); r.c = hex(over(rgba(pcs.backgroundColor), ctx.bg));
    }
    if (!r) return;
    var blk = !abs && !/^inline|^contents$/.test(pcs.display);
    if (abs) {   // a marker drawn beside the text, as lists do; a filled one (a step's circle) keeps its fill
      r.pre = 1;
      var fill = rgba(pcs.backgroundColor);
      if (fill && hex(over(fill, ctx.bg)) !== hex(ctx.bg)) { r.bg = hex(over(fill, ctx.bg)); r.t = ' ' + r.t.trim() + ' '; }
    }
    else if (which === '::before' && !blk && num(pcs.marginRight) > 0 && !/\s$/.test(r.t)) r.t += ' ';
    // a block marker (a chapter's number over its title) becomes a paragraph of its own in emit(), with its own spacing
    var cut = blk && {which: which, sz: num(pcs.fontSize), mt: num(pcs.marginTop), mb: num(pcs.marginBottom),
                      lh: pcs.lineHeight === 'normal' ? num(pcs.fontSize) * 1.2 : num(pcs.lineHeight)};
    if (blk && which === '::after') p.runs.push({br: 1, split: cut});
    p.runs.push(r);
    if (blk && which === '::before') p.runs.push({br: 1, split: cut});
  }
  function tick(el, cs) {
    if (el.tagName === 'INPUT' && /checkbox|radio/.test(el.type)) return el.checked ? '☒' : '☐';
    if (textIn(el) && !/[✓✔✗xX]/.test(el.textContent.trim())) return '';
    var r = el.getBoundingClientRect(), s = num(cs.fontSize) * 2.6, b = borders(cs);
    if (!r.width || r.width > s || r.height > s || Math.abs(r.width - r.height) > 0.35 * Math.max(r.width, r.height)) return '';
    if (!(b.t && b.r && b.b && b.l)) return '';
    return textIn(el) || el.querySelector('svg, img') || /checked|done|is-on/.test(el.className) ? '☒' : '☐';
  }
  function fillLine(el, cs) {
    if (textIn(el) || el.children.length || !/inline/.test(cs.display)) return 0;
    var b = borders(cs), w = el.getBoundingClientRect().width;
    return b.b && !b.t && w > num(cs.fontSize) * 1.5 ? w : 0;
  }
  function each(el, fn) {
    var frame = [];
    for (var n = el.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 3) { fn(n, null); continue; }
      if (n.nodeType !== 1) continue;
      var cs = getComputedStyle(n);
      if (hidden(n, cs)) continue;
      enter(cs, frame);
      fn(n, cs);
    }
    leave(frame);
  }
  function inline(node, cs, ctx, p) {
    if (node.nodeType === 3) {
      var pcs = getComputedStyle(node.parentElement), t = textOf(node, pcs);
      if (t) p.runs.push(run(t, pcs, ctx, node.parentElement));
      return;
    }
    var el = node, k;
    if (el.tagName === 'BR') { p.runs.push({br: 1}); return; }
    if (el.matches(NOTE)) {
      var fp = {runs: []};
      each(el, function (n, c) { if (!(n.nodeType === 1 && n.matches('.pc-fn-n'))) inline(n, c, ctx, fp); });
      p.runs.push({fn: fp.runs});
      return;
    }
    if (el.matches(PIC)) { if (!ctx.flat) { p.runs.push({pic: picture(el)}); gapRun(el, ctx, p); } return; }
    if ((k = tick(el, cs))) { p.runs.push(run(k, cs, ctx, el)); return; }
    if ((k = fillLine(el, cs))) {
      var r = run(new Array(Math.max(4, Math.round(k / (num(cs.fontSize) * 0.26)))).join(' '), cs, ctx, el);
      r.u = 1; r.c = hex(rgba(cs.borderBottomColor)) || r.c; p.runs.push(r);
      return;
    }
    pseudo(el, '::before', ctx, p);
    each(el, function (n, c) { inline(n, c, ctx, p); });
    pseudo(el, '::after', ctx, p);
    if ((/^inline-(block|flex|grid)/.test(cs.display) && num(cs.marginRight) > 0) || el.matches('.pc-toc-n')) {
      var ps = el.parentElement || el, at = el.matches('.pc-toc-n') ? nextLeft(el) : null, sp = run(at === null ? ' ' : '\t', getComputedStyle(ps), ctx, ps);
      if (at !== null) sp.tab_at = at - ctx.left;   // a contents number in its own column: a tab to where the title starts
      p.runs.push(sp);
    } else gapRun(el, ctx, p);
  }
  function gapRun(el, ctx, p) {   // the gap between flex or grid items laid out as one line reads as a space
    var ps = el.parentElement, nx = el.nextSibling;
    while (nx && nx.nodeType !== 1 && !/\S/.test(nx.textContent || '')) nx = nx.nextSibling;
    if (!ps || !nx) return;
    var pc = getComputedStyle(ps);
    if (/flex|grid/.test(pc.display) && num(pc.columnGap) > 0) p.runs.push(run(' ', pc, ctx, ps));
  }
  function lineRow(el, cs) {   // a flex row that reads as one line of words (a callout label: icon and text) becomes one paragraph
    if (!/flex/.test(cs.display) || /column/.test(cs.flexDirection)) return false;
    var fs = num(cs.fontSize), r = rect(el), kids = [].filter.call(el.children, function (k) { return !hidden(k, getComputedStyle(k)); });
    if (!kids.length || kids.length > 6 || r.h > 2.6 * fs) return false;
    var rs = kids.map(rect).sort(function (a, b) { return a.x - b.x; });
    for (var i = 0; i < kids.length; i++) {
      var k = kids[i], kr = rect(k);
      if (/absolute|fixed/.test(getComputedStyle(k).position) || kr.y >= r.y + r.h || kr.y + kr.h <= r.y) return false;
      var icon = k.matches(PIC) || (!/\S/.test(k.textContent) && k.querySelector(PIC));   // an icon, bare or in its wrapper
      if (icon ? kr.h > 2.5 * fs || kr.w > 4 * fs : blockKids(k) || k.matches(NOTE)) return false;
    }
    for (i = 1; i < rs.length; i++) if (rs[i].x - (rs[i - 1].x + rs[i - 1].w) > Math.max(12, 1.5 * fs)) return false;
    return true;
  }
  function nextLeft(el) {   // where the first visible text or box after el starts (document px), or null
    var rg = document.createRange();
    for (var n = el.nextSibling; n; n = n.nextSibling) {
      if (n.nodeType === 3) {
        var i = n.data.search(/\S/);
        if (i < 0) continue;
        rg.setStart(n, i); rg.setEnd(n, i + 1);
        var q = rg.getClientRects()[0];
        if (q) return q.left + scrollX;
      } else if (n.nodeType === 1) {
        var b = n.getBoundingClientRect();
        if (b.width > 0) return b.left + scrollX;
      }
    }
    return null;
  }
  function picture(el, rr) {
    var r = rr || rect(el), id = pics.length;
    var jpeg = el.tagName === 'IMG' && /\.jpe?g(\?|$)|^data:image\/jpeg/i.test(el.currentSrc || el.src || '');
    pics.push({id: id, x: r.x, y: r.y, w: r.w, h: r.h, jpeg: !!jpeg, alt: el.getAttribute('alt') || el.getAttribute('aria-label') || ''});
    return {id: id, w: r.w, h: r.h};
  }

  // ---- blocks
  function inlineLevel(el, cs) { return el.matches(NOTE) || /^inline|^contents$/.test(cs.display) && !(el.matches(PIC) && el.getBoundingClientRect().width > 0.6 * docW); }
  function blockKids(el) {
    var kids = el.children;
    for (var i = 0; i < kids.length; i++) {
      var cs = getComputedStyle(kids[i]);
      if (!hidden(kids[i], cs) && !inlineLevel(kids[i], cs)) return true;
    }
    return false;
  }
  function pf(el, cs, ctx) {
    var r = rect(el), b = borders(cs), bg = rgba(cs.backgroundColor), deco = ctx.noBox !== el && !!(bg || b.t || b.b || b.l || b.r), pd = pads(cs);
    var left = ctx.colLeft ? ctx.colLeft(r.x) : ctx.left;
    var o = {
      top: r.y + (deco ? 0 : pd[0] + num(cs.borderTopWidth)), bottom: r.y + r.h - (deco ? 0 : pd[2] + num(cs.borderBottomWidth)),
      il: r.x + num(cs.borderLeftWidth) + (deco ? 0 : pd[3]) - left, ir: ctx.right - (r.x + r.w - num(cs.borderRightWidth) - (deco ? 0 : pd[1])),
      ti: num(cs.textIndent), al: cs.textAlign, lh: cs.lineHeight === 'normal' ? 0 : num(cs.lineHeight), sz: num(cs.fontSize),
      kn: /avoid/.test(cs.breakAfter) || /^H[1-6]$/.test(el.tagName) ? 1 : 0, kl: /avoid/.test(cs.breakInside) ? 1 : 0
    };
    if (ctx.colLeft) o.ir = Math.max(0, ctx.colW - (r.x + r.w - left));
    if (deco) {
      o.bd = b; o.pad = pd;
      if (bg) o.shade = hex(over(bg, ctx.bg));
    }
    o.mt = num(cs.marginTop);
    if (el.closest && el.closest('.pc-cover .pc-meta')) o.meta = 1;
    return o;
  }
  function para(el, cs, ctx) {
    var p = {k: 'p', tag: el.tagName.toLowerCase(), cls: el.className && el.className.baseVal === undefined ? el.className : '', runs: []};
    p.pf = pf(el, cs, ctx);
    if (el.id) p.id = el.id;
    var bctx = ctx;
    if (rgba(cs.backgroundColor)) bctx = Object.assign({}, ctx, {bg: over(rgba(cs.backgroundColor), ctx.bg)});
    pseudo(el, '::before', bctx, p);
    each(el, function (n, c) { inline(n, c, bctx, p); });
    pseudo(el, '::after', bctx, p);
    return p;
  }
  function emit(p, ctx) {
    var si = -1;
    for (var i = 0; i < p.runs.length; i++) if (p.runs[i].split) { si = i; break; }
    if (si >= 0) {
      var sp = p.runs[si].split;
      delete p.runs[si].split;
      if (!ctx.list && !p.pf.bd && !p.pf.shade && p.k === 'p') {   // a marker on a line of its own: two paragraphs
        var h = sp.lh, q = {k: 'p', tag: 'pseudo', cls: p.cls, runs: null, pf: Object.assign({}, p.pf, {lh: sp.lh, sz: sp.sz, ti: 0})};
        if (sp.which === '::before') {
          q.runs = p.runs.slice(0, si); p.runs = p.runs.slice(si + 1);
          q.pf.top = p.pf.top + sp.mt; q.pf.bottom = q.pf.top + h; q.pf.kn = 1;
          p.pf.top = q.pf.bottom + sp.mb; p.pf.mt = sp.mb;
          emit(q, ctx); emit(p, ctx);
        } else {
          q.runs = p.runs.slice(si + 1); p.runs = p.runs.slice(0, si);
          q.pf.bottom = p.pf.bottom - sp.mb; q.pf.top = q.pf.bottom - h; q.pf.mt = sp.mt;
          p.pf.bottom = q.pf.top - sp.mt; p.pf.kn = 1;
          emit(p, ctx); emit(q, ctx);
        }
        return;
      }
    }
    if (!p.runs.some(function (r) { return r.pic || r.fn || r.br || (r.t && /\S/.test(r.t)); })) return;
    if (ctx.brk) { p.pf.pb = 1; ctx.brk = 0; }
    if (ctx.list) { p.list = ctx.list; ctx.list = null; }
    if (ctx.pid) { p.ids = ctx.pid; ctx.pid = null; }
    ctx.out.push(p);
  }
  function anon(el, cs, ctx) { var p = {k: 'p', tag: 'anon', runs: [], range: document.createRange()}; p.pf = pf(el, cs, ctx); p.first = null; return p; }
  function closeAnon(p, ctx) {
    if (p.first) {
      p.range.setStartBefore(p.first); p.range.setEndAfter(p.last);
      var r = p.range.getBoundingClientRect();
      if (r.height) { p.pf.top = r.top + scrollY; p.pf.bottom = r.bottom + scrollY; }
    }
    delete p.range; delete p.first; delete p.last;
    emit(p, ctx);
  }
  function flow(el, ctx) {
    var p = null, ecs = getComputedStyle(el);
    each(el, function (n, cs) {
      if (n.nodeType === 3 || inlineLevel(n, cs)) {
        if (n.nodeType === 3 && !p && !/\S/.test(n.data)) return;
        p = p || anon(el, ecs, ctx);
        p.first = p.first || n; p.last = n;
        inline(n, cs, ctx, p);
        return;
      }
      if (p) { closeAnon(p, ctx); p = null; }
      block(n, cs, ctx);
    });
    if (p) closeAnon(p, ctx);
  }
  function sub(ctx, left, right, bg, extra) {
    return Object.assign({out: [], left: left, right: right, bg: bg || ctx.bg, brk: 0, list: null, flat: ctx.flat}, extra || {});
  }
  function isBox(el, cs, ctx) {
    if (ctx.flat || ctx.noBox === el) return false;
    var bg = rgba(cs.backgroundColor), b = borders(cs), n = (b.t ? 1 : 0) + (b.r ? 1 : 0) + (b.b ? 1 : 0) + (b.l ? 1 : 0);
    var kids = 0;   // a plain loop: each() would run the children's counter increments a second time
    for (var k = el.firstElementChild; k; k = k.nextElementSibling) {
      var c = getComputedStyle(k);
      if (!hidden(k, c) && !inlineLevel(k, c)) kids++;
    }
    if (bg && hex(over(bg, ctx.bg)) === hex(ctx.bg)) bg = null;
    return (bg || n >= 3 || (n && kids >= 2)) && textIn(el);
  }
  function rowKids(el, cs) {
    if (!/flex|grid/.test(cs.display) || (/flex/.test(cs.display) && /column/.test(cs.flexDirection))) return null;
    var kids = [];
    (function add(p) {   // a display:contents child (dl.pc-facts' rows) gives its own children to the grid
      [].forEach.call(p.children, function (k) {
        var s = getComputedStyle(k);
        if (s.display === 'contents' && !hidden(k, s)) add(k);
        else if (!hidden(k, s) && !/absolute|fixed/.test(s.position) && k.getBoundingClientRect().width > 0) kids.push(k);
      });
    })(el);
    if (kids.length < 2) return null;
    var rs = kids.map(rect);
    for (var i = 0; i < rs.length; i++) for (var j = i + 1; j < rs.length; j++)
      if (Math.min(rs[i].y + rs[i].h, rs[j].y + rs[j].h) - Math.max(rs[i].y, rs[j].y) > 2 && Math.abs(rs[i].x - rs[j].x) > 2) return {kids: kids, rs: rs};
    return null;
  }
  function cell(el, cs, ctx, r) {
    var b = borders(cs), pd = pads(cs), bg = over(rgba(cs.backgroundColor), ctx.bg);
    var c = sub(ctx, r.x + num(cs.borderLeftWidth) + pd[3], r.x + r.w - num(cs.borderRightWidth) - pd[1], bg, {noBox: el});
    c.top0 = r.y + num(cs.borderTopWidth) + pd[0];
    var va = cs.verticalAlign === 'middle' || /center/.test(getComputedStyle(el.parentElement).alignItems || '') ? 'center' :
      cs.verticalAlign === 'bottom' ? 'bottom' : 'top';
    if (blockKids(el) || rowKids(el, cs)) { pseudoBlock(el, '::before', cs, c); inner(el, cs, c); pseudoBlock(el, '::after', cs, c); } else emit(para(el, cs, c), c);
    var shade = rgba(cs.backgroundColor) ? hex(bg) : null;
    return {x0: r.x, x1: r.x + r.w, bd: b, pad: pd, shade: shade, va: va, blocks: c.out, top0: c.top0, bot0: r.y + r.h - num(cs.borderBottomWidth) - pd[2]};
  }
  function grid(el, cs, ctx, rk) {
    var items = rk.kids.map(function (k, i) { return {el: k, r: rk.rs[i]}; }).sort(function (a, b) { return a.r.y - b.r.y || a.r.x - b.r.x; });
    var rows = [], cur = null;
    items.forEach(function (it) {
      if (!cur || it.r.y >= cur.bottom - 2) { cur = {top: it.r.y, bottom: it.r.y + it.r.h, items: []}; rows.push(cur); }
      cur.items.push(it); cur.bottom = Math.max(cur.bottom, it.r.y + it.r.h);
    });
    var r = rect(el), out = [];
    rows.forEach(function (row, i) {
      if (i && row.top - rows[i - 1].bottom > 2) out.push({gap: row.top - rows[i - 1].bottom});
      out.push({cells: row.items.sort(function (a, b) { return a.r.x - b.r.x; }).map(function (it) {
        return cell(it.el, getComputedStyle(it.el), ctx, it.r);
      })});
    });
    return {k: 'table', kind: 'grid', x: r.x, w: r.w, top: r.y, bottom: r.y + r.h, il: r.x - ctx.left, rows: out};
  }
  function table(el, cs, ctx) {
    var r = rect(el), out = [];
    [].forEach.call(el.rows || [], function (tr) {
      var tcs = getComputedStyle(tr);
      if (hidden(tr, tcs)) return;
      var bg = over(rgba(tcs.backgroundColor), ctx.bg), rctx = Object.assign({}, ctx, {bg: bg});
      var row = {head: tr.parentElement.tagName === 'THEAD' ? 1 : 0, cells: []};
      [].forEach.call(tr.cells, function (td) {
        var dcs = getComputedStyle(td);
        if (hidden(td, dcs)) return;
        var c = cell(td, dcs, rctx, rect(td));
        if (td.rowSpan > 1) c.rspan = td.rowSpan;
        if (!c.shade && rgba(tcs.backgroundColor)) c.shade = hex(bg);
        row.cells.push(c);
      });
      out.push(row);
    });
    return {k: 'table', kind: 'table', x: r.x, w: r.w, top: r.y, bottom: r.y + r.h, il: r.x - ctx.left, rows: out};
  }
  function box(el, cs, ctx) {
    var r = rect(el), c = cell(el, cs, ctx, r);
    return {k: 'table', kind: 'box', x: r.x, w: r.w, top: r.y, bottom: r.y + r.h, il: r.x - ctx.left, rows: [{cells: [c]}]};
  }
  function inner(el, cs, ctx) {
    var rk = rowKids(el, cs);
    if (rk && !ctx.flat) { var t = grid(el, cs, ctx, rk); if (ctx.brk) { t.pb = 1; ctx.brk = 0; } ctx.out.push(t); }
    else flow(el, ctx);
  }
  function rule(el, cs, ctx) {
    var b = borders(cs), r = rect(el);
    if (textIn(el) || el.children.length || !(b.b || b.t) || (b.l || b.r)) return null;
    return {k: 'rule', top: r.y, bottom: r.y + r.h, il: r.x - ctx.left, ir: ctx.right - (r.x + r.w), side: b.b ? 'b' : 't', bd: b.b || b.t};
  }
  function toc(el, ctx) {
    var head = el.querySelector('h1, h2, .pc-toc-title');
    if (head) { var h = para(head, getComputedStyle(head), ctx); h.toc_title = 1; emit(h, ctx); }
    var entries = [];
    [].forEach.call(el.querySelectorAll('a'), function (a) {
      var acs = getComputedStyle(a);
      if (hidden(a, acs)) return;
      var lvl = 0, n = a, li = a.closest('li'), lm = li && /pc-toc-(\d)/.exec(li.className);
      while ((n = n.parentElement) && n !== el) if (n.tagName === 'OL' || n.tagName === 'UL') lvl++;
      if (lm) lvl = +lm[1];
      var kids = [].slice.call(a.children), last = kids[kids.length - 1], page = '', p = para(a, acs, ctx);
      if (last && /^\s*[\divxlc]+\s*$/i.test(last.textContent)) {
        page = last.textContent.trim();
        var lp = {runs: []};
        inline(last, getComputedStyle(last), ctx, lp);
        p.runs = p.runs.filter(function (r) { return !(r.t && r.t.trim() === page); });
        p.page_run = lp.runs[0];
      }
      p.lvl = Math.max(1, lvl); p.page = page; p.href = a.getAttribute('href');
      var dots = kids.filter(function (k) { return getComputedStyle(k).borderBottomStyle === 'dotted'; })[0];
      p.leader = dots ? 'dot' : 'none';
      if (dots) { var dr = rect(dots); if (dr.w > 0) p.dots_end = dr.x + dr.w - ctx.left; }   // the dots stop short of the page number
      entries.push(p);
    });
    var r = rect(el);
    ctx.out.push({k: 'toc', top: r.y, bottom: r.y + r.h, entries: entries, pb: ctx.brk});
    ctx.brk = 0;
  }
  function cover(el, cs, ctx) {
    var r = rect(el), c = sub(ctx, r.x, r.x + r.w, over(rgba(cs.backgroundColor), ctx.bg), {flat: true});
    flow(el, c);
    c.out.forEach(function (b) { b.ax = r.x; b.ay = r.y; });
    ctx.out.push({k: 'cover', x: r.x, y: r.y, w: r.w, h: r.h, blocks: c.out});
  }
  function block(el, cs, ctx) {
    if (ctx.top) {
      var pg = cs.page && cs.page !== 'auto' ? cs.page : '';
      if (pg !== ctx.page) { ctx.out.push({k: 'sect', page: pg}); ctx.page = pg; ctx.brk = 0; }
    }
    if (/page|left|right|recto|verso/.test(cs.breakBefore)) ctx.brk = 1;
    if (ctx.top && el.matches('.pc-cover')) { cover(el, cs, ctx); return; }
    if (el.matches('.pc-toc')) { toc(el, ctx); if (/page|left|right|recto|verso/.test(cs.breakAfter)) ctx.brk = 1; return; }
    var t;
    if (el.matches(PIC)) {
      if (!ctx.flat) {   // a figure's caption stays text: the picture is the rest of the figure
        var r = rect(el), fc = el.tagName === 'FIGURE' && el.querySelector(':scope > figcaption'), fcs = fc && getComputedStyle(fc), capBelow = false;
        if (fc && !hidden(fc, fcs)) {
          var cr = rect(fc);
          capBelow = cr.y > r.y + r.h / 2;
          r = capBelow ? {x: r.x, y: r.y, w: r.w, h: cr.y - num(fcs.marginTop) - r.y}
                       : {x: r.x, y: cr.y + cr.h + num(fcs.marginBottom), w: r.w, h: r.y + r.h - (cr.y + cr.h + num(fcs.marginBottom))};
          if (!capBelow) { var cp0 = para(fc, fcs, ctx); cp0.pf.kn = 1; emit(cp0, ctx); }
        } else fc = null;
        var pc = picture(el, r), il = r.x - ctx.left, ir = ctx.right - (r.x + r.w);
        ctx.out.push({k: 'pic', pic: pc, top: r.y, bottom: r.y + r.h, il: il, al: Math.abs(il - ir) < 4 && il > 4 ? 'center' : il > ir + 4 ? 'right' : 'left', pb: ctx.brk});
        ctx.brk = 0;
        if (fc && capBelow) emit(para(fc, fcs, ctx), ctx);
      }
    } else if (cs.display === 'table' || el.tagName === 'TABLE') {
      if (ctx.flat) flow(el, ctx);
      else {   // a caption becomes its own paragraph ("Table 1." from its ::before), kept with the table
        var cap = el.tagName === 'TABLE' && el.caption, ccs = cap && getComputedStyle(cap), up = cap && !hidden(cap, ccs) && ccs.captionSide !== 'bottom';
        if (up) { var cp = para(cap, ccs, ctx); cp.pf.kn = 1; emit(cp, ctx); }
        t = table(el, cs, ctx); t.pb = ctx.brk; ctx.brk = 0;
        if (up) { var cr = rect(cap); t.top = Math.max(t.top, cr.y + cr.h + num(ccs.marginBottom)); }
        ctx.out.push(t);
        if (cap && !hidden(cap, ccs) && !up) emit(para(cap, ccs, ctx), ctx);
      }
    } else if (cs.display === 'list-item') {
      var list = el.parentElement, id = lists.get(list);
      if (id === undefined) { id = ++nlist; lists.set(list, id); }
      var lvl = 0, n = el;
      while ((n = n.parentElement) && !n.matches('td, th, .pc-doc, body')) if (n.tagName === 'UL' || n.tagName === 'OL') lvl++;
      var mk = getComputedStyle(el, '::marker'), lr = rect(list), lcs = getComputedStyle(list);
      ctx.list = {id: id, lvl: Math.max(0, lvl - 1), type: cs.listStyleType, start: list.tagName === 'OL' ? (list.start || 1) : 1,
        mc: hex(over(rgba(mk.color), ctx.bg)), mf: family(mk.fontFamily), mw: weight(mk.fontWeight), ms: num(mk.fontSize),
        hang: Math.max(0, rect(el).x - (lr.x + num(lcs.paddingLeft) * 0)) };
      var bm = getComputedStyle(el, '::before');   // a marker drawn by an absolute ::before (steps): the text hangs past it
      if (cs.listStyleType === 'none' && bm.position === 'absolute' && bm.content && bm.content !== 'none')
        ctx.list.hang = Math.max(ctx.list.hang, num(cs.paddingLeft) + num(cs.borderLeftWidth) - num(bm.left));
      var bctx = Object.assign(ctx, {bullet: cs.listStyleType === 'none'});
      if (blockKids(el)) flow(el, bctx); else emit(para(el, cs, bctx), bctx);
      ctx.bullet = false; ctx.list = null;
    } else if (!ctx.flat && cs.columnCount !== 'auto' && parseInt(cs.columnCount, 10) > 1) {
      var r2 = rect(el), n2 = parseInt(cs.columnCount, 10), gap = num(cs.columnGap), cw = (r2.w - gap * (n2 - 1)) / n2;
      var cc = sub(ctx, r2.x, r2.x + cw, ctx.bg, {colW: cw, colLeft: function (x) { return r2.x + Math.max(0, Math.floor((x - r2.x + 1) / (cw + gap))) * (cw + gap); }});
      flow(el, cc);
      ctx.out.push({k: 'cols', n: n2, gap: gap, top: r2.y, bottom: r2.y + r2.h, blocks: cc.out, pb: ctx.brk});
      ctx.brk = 0;
    } else if ((t = rule(el, cs, ctx))) {
      ctx.out.push(t);
    } else if (isBox(el, cs, ctx)) {
      t = box(el, cs, ctx); t.pb = ctx.brk; ctx.brk = 0; ctx.out.push(t);
    } else if (!ctx.flat && lineRow(el, cs)) {
      var lp = para(el, cs, ctx);
      if (el.id) lp.id = el.id;
      emit(lp, ctx);
    } else if (blockKids(el) || (!ctx.flat && rowKids(el, cs))) {
      var top = ctx.top;
      if (el.id) ctx.pid = (ctx.pid || []).concat(el.id);
      ctx.top = false;
      pseudoBlock(el, '::before', cs, ctx);
      inner(el, cs, ctx);
      pseudoBlock(el, '::after', cs, ctx);
      ctx.top = top;
    } else {
      var p = para(el, cs, ctx);
      if (el.id) p.id = el.id;
      emit(p, ctx);
    }
    if (/page|left|right|recto|verso/.test(cs.breakAfter)) ctx.brk = 1;
  }
  function pseudoBlock(el, which, cs, ctx) {
    var pcs = getComputedStyle(el, which), t = content(el, pcs);
    if (pcs.display === 'none' || !t || !/\S/.test(t)) return;
    var p = {k: 'p', tag: 'pseudo', runs: [run(t, pcs, ctx, el)]};
    p.pf = pf(el, cs, ctx);
    p.pf.bottom = p.pf.top + num(pcs.lineHeight) || p.pf.top + num(pcs.fontSize) * 1.2;
    emit(p, ctx);
  }

  // ---- the page set-up, from the @page rules and their margin boxes
  function resolve(v) {
    for (var i = 0; i < 5 && /var\(/.test(v); i++)
      v = v.replace(/var\((--[\w-]+)(?:,\s*((?:[^()]|\([^()]*\))*))?\)/g, function (m, n, d) {
        var x = rootCS.getPropertyValue(n) || docCS.getPropertyValue(n);
        return x && x.trim() ? x.trim() : (d || '');
      });
    return v;
  }
  function boxStyle(sts) {
    // A margin box's style from every @page rule that styles it, in cascade order (the default page's rules, then the
    // named page's), so a theme rule that adds only a rule line keeps the engine's content and font.
    function last(p) {
      var v = '';
      sts.forEach(function (st) { var x = st.getPropertyValue(p); if (x && x.trim()) v = x; });
      return resolve(v).trim();
    }
    var probe = document.createElement('span');
    probe.style.cssText = sts.map(function (st) { return resolve(st.cssText); }).join(';');
    document.body.appendChild(probe);
    var cs = getComputedStyle(probe);
    var o = {f: family(cs.fontFamily), w: weight(cs.fontWeight), sz: num(cs.fontSize), c: hex(rgba(cs.color)), ls: num(cs.letterSpacing),
      caps: /uppercase/.test(last('text-transform')) ? 1 : 0, pt: num(cs.paddingTop), pb: num(cs.paddingBottom),
      mt: num(cs.marginTop), mb: num(cs.marginBottom), bt: side(cs, 'Top'), bb: side(cs, 'Bottom'),
      va: last('vertical-align') || 'middle', parts: []};
    probe.remove();
    var c = last('content');
    o.page = /counter\(\s*page\s*\)/.test(c) ? 1 : 0;
    var pr = /"((?:[^"\\]|\\.)*)"|counter\(\s*(pages?)\s*\)/g, pm;
    while ((pm = pr.exec(c))) o.parts.push(pm[2] ? {f: pm[2] === 'page' ? 'PAGE' : 'NUMPAGES'} : {t: unesc(pm[1])});
    o.text = (c.match(/"((?:[^"\\]|\\.)*)"/g) || []).map(function (s) { return unesc(s.slice(1, -1)); }).join('');
    if (o.text || o.parts.length) fonts[o.f + '|' + o.w + '|0'] = 1;
    return o;
  }
  function setup() {
    var pages = {}, styles = {};
    [].forEach.call(document.styleSheets, function (sh) {
      var rules;
      try { rules = sh.cssRules; } catch (e) { return; }
      (function scan(list) {
        [].forEach.call(list, function (r) {
          if (r.type === CSSRule.PAGE_RULE) {
            var name = (r.selectorText || '').trim(), o = pages[name] = pages[name] || {boxes: {}}, bs = styles[name] = styles[name] || {};
            ['size', 'margin-top', 'margin-right', 'margin-bottom', 'margin-left', 'margin'].forEach(function (p) {
              var v = r.style.getPropertyValue(p);
              if (v) o[p] = resolve(v);
            });
            [].forEach.call(r.cssRules || [], function (m) { if (m.name) (bs[m.name] = bs[m.name] || []).push(m.style); });
          } else if (r.cssRules) scan(r.cssRules);
        });
      })(rules);
    });
    var base = styles[''] || {};
    Object.keys(pages).forEach(function (name) {
      var own = styles[name] || {}, names = Object.keys(own);
      if (name) Object.keys(base).forEach(function (b) { if (names.indexOf(b) < 0) names.push(b); });
      names.forEach(function (b) {
        pages[name].boxes[b] = boxStyle(name ? (base[b] || []).concat(own[b] || []) : own[b]);
      });
    });
    return pages;
  }

  var data = {page: htmlEl.getAttribute('data-page') || '', orient: htmlEl.getAttribute('data-orient') || ''};
  if (mode === 'pages') return {pages: setup(), data: data};
  var body = docEl.getBoundingClientRect(), docW = body.width;
  var pageBg = over(rgba(docCS.backgroundColor), over(rgba(getComputedStyle(document.body).backgroundColor), over(rgba(rootCS.backgroundColor), [255, 255, 255, 1])));
  var ctx = {out: [], left: body.left + scrollX, right: body.right + scrollX, bg: pageBg, brk: 0, list: null, top: true, page: ''};
  flow(docEl, ctx);
  return {blocks: ctx.out, pics: pics, fonts: Object.keys(fonts), pages: setup(), width: docW, top0: body.top + scrollY,
    bg: hex(pageBg), lang: document.documentElement.lang || '', title: document.title || '',
    data: data};
};
