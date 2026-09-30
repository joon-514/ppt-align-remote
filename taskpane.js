/* 정렬 리모컨 — PowerPoint 연동 + 패널 UI */
(function () {
  "use strict";
  var Core = window.AlignCore;

  // ---------- 상태 ----------
  var state = {
    mode: load("mode", "selection"),   // selection | key | slide
    unit: load("unit", "cm"),
    lock: load("lock", "0") === "1",
    keysOn: load("keysOn", "1") === "1",
    keys: {},         // 단축키 동작 id → 조합 문자열 (예: "Command+Option+1")
    keysApi: false,   // 단축키 변경 API 지원 여부
    slidePreset: load("slidePreset", "auto"),
    customW: load("customW", ""),
    customH: load("customH", ""),
    order: [],        // 선택한 순서대로 쌓이는 id 목록
    manualKey: null,  // 목록에서 직접 찍은 기준 오브젝트
    shapes: [],       // 현재 선택 {id,name,left,top,width,height,rotation,adj}
    detected: null,   // 자동 감지된 슬라이드 크기 (pt)
    adjScale: null,   // 'frac'(0~0.5) | 'ooxml'(0~50000) — 처음 읽은 값으로 판별
    inOffice: false,
    api110: false,    // 회전·라운드 지원 여부
  };
  var history = [], future = [];
  var busy = false;

  function load(k, d) { try { var v = localStorage.getItem("alignRemote." + k); return v === null ? d : v; } catch (e) { return d; } }
  function save(k, v) { try { localStorage.setItem("alignRemote." + k, v); } catch (e) {} }

  var $ = function (s) { return document.querySelector(s); };
  var $$ = function (s) { return Array.prototype.slice.call(document.querySelectorAll(s)); };

  // ---------- 값 변환 ----------
  function fmt(pt) { return pt === null || pt === undefined ? "" : String(parseFloat(Core.fromPt(pt, state.unit).toFixed(2))); }
  function readNum(el) { var v = parseFloat(el.value); return isFinite(v) ? v : null; }
  function normRot(r) { r = ((r % 360) + 360) % 360; return r > 180 ? r - 360 : r; }
  function toFrac(v) { return (state.adjScale === "ooxml") ? v / 100000 : v; }
  function fromFrac(f) { return (state.adjScale === "ooxml") ? Math.round(f * 100000) : f; }
  function noteAdjScale(v) {
    if (state.adjScale || !(v > 0)) return;
    state.adjScale = v > 1.0001 ? "ooxml" : "frac";
  }

  // ---------- 기준 / 슬라이드 크기 ----------
  function effectiveKey() {
    var ids = state.shapes.map(function (s) { return s.id; });
    if (state.manualKey && ids.indexOf(state.manualKey) !== -1) return state.manualKey;
    return state.order.length ? state.order[0] : null;
  }

  function slideSize() {
    if (state.slidePreset === "auto") return state.detected || { width: 960, height: 540 };
    if (state.slidePreset === "custom") {
      var w = parseFloat(state.customW), h = parseFloat(state.customH);
      if (w > 0 && h > 0) return { width: Core.toPt(w, "cm"), height: Core.toPt(h, "cm") };
      return state.detected || { width: 960, height: 540 };
    }
    var p = state.slidePreset.split("x");
    return { width: +p[0], height: +p[1] };
  }

  function coreOpts(extra) {
    var o = { mode: state.mode, keyId: effectiveKey(), slide: slideSize(), gap: null };
    for (var k in extra) o[k] = extra[k];
    return o;
  }

  // ---------- PowerPoint 읽기 ----------
  async function readSel(ctx, withAdj) {
    var coll = ctx.presentation.getSelectedShapes();
    coll.load("items/id,items/name,items/left,items/top,items/width,items/height,items/type" + (state.api110 ? ",items/rotation" : ""));
    var slides = ctx.presentation.getSelectedSlides();
    slides.load("items/id");
    await ctx.sync();
    await detectSlideSize(ctx);

    var adj = {};
    if (withAdj && state.api110) {
      try {
        var geo = coll.items.filter(function (s) { return s.type === "GeometricShape"; });
        geo.forEach(function (s) { s.adjustments.load("count"); });
        await ctx.sync();
        var res = [];
        geo.forEach(function (s) { if (s.adjustments.count > 0) res.push([s.id, s.adjustments.get(0)]); });
        await ctx.sync();
        res.forEach(function (p) { adj[p[0]] = p[1].value; noteAdjScale(p[1].value); });
      } catch (e) { /* 조절점 없는 도형 */ }
    }

    var shapes = coll.items.map(function (s) {
      return {
        id: s.id, name: s.name, left: s.left, top: s.top, width: s.width, height: s.height,
        rotation: state.api110 ? s.rotation : null,
        adj: adj.hasOwnProperty(s.id) ? adj[s.id] : undefined,
      };
    });
    return { coll: coll, shapes: shapes, slideId: slides.items.length ? slides.items[0].id : null };
  }

  async function detectSlideSize(ctx) {
    if (state.detected) return;
    try {
      if (state.api110 && ctx.presentation.pageSetup) {
        var ps = ctx.presentation.pageSetup;
        ps.load("slideWidth,slideHeight");
        await ctx.sync();
        if (ps.slideWidth > 0 && ps.slideHeight > 0) state.detected = { width: ps.slideWidth, height: ps.slideHeight };
      }
    } catch (e) { /* 구버전: 수동 설정 사용 */ }
  }

  function applySelection(list) {
    state.shapes = list;
    state.order = Core.updateOrder(state.order, list.map(function (s) { return s.id; }));
    if (state.manualKey && !list.some(function (s) { return s.id === state.manualKey; })) state.manualKey = null;
    render();
  }

  async function refreshSelection() {
    if (!state.inOffice || busy) return;
    busy = true;
    try {
      await PowerPoint.run(async function (ctx) {
        var sel = await readSel(ctx, true);
        applySelection(sel.shapes);
      });
    } catch (e) {
      applySelection([]);
    } finally { busy = false; }
  }

  var refreshTimer = null;
  function scheduleRefresh() { clearTimeout(refreshTimer); refreshTimer = setTimeout(refreshSelection, 60); }

  // ---------- 변경 적용 (+ 되돌리기 기록) ----------
  // change: { left, top, width, height, rotation, adj }  (adj = 0~0.5 비율)
  function applyChange(item, c) {
    if (c.width !== undefined) item.width = c.width;
    if (c.height !== undefined) item.height = c.height;
    if (c.left !== undefined) item.left = c.left;
    if (c.top !== undefined) item.top = c.top;
    if (c.rotation !== undefined && c.rotation !== null && state.api110) item.rotation = ((c.rotation % 360) + 360) % 360;
    if (c.adj !== undefined && c.adj !== null && state.api110) item.adjustments.set(0, fromFrac(c.adj));
  }

  function snapshotOf(s, slideId) {
    return { slideId: slideId, id: s.id, left: s.left, top: s.top, width: s.width, height: s.height,
      rotation: s.rotation, adj: s.adj === undefined ? undefined : toFrac(s.adj) };
  }

  // planner(shapes) → { error } | { changes: {id: change}, msg }
  async function mutate(label, planner, needAdj) {
    if (!state.inOffice) return demoMutate(label, planner);
    if (busy) return;
    busy = true;
    try {
      await PowerPoint.run(async function (ctx) {
        var sel = await readSel(ctx, true);
        state.shapes = sel.shapes;
        state.order = Core.updateOrder(state.order, sel.shapes.map(function (s) { return s.id; }));
        if (!sel.shapes.length) { setStatus("먼저 오브젝트를 선택하세요.", "err"); return; }

        var plan = planner(sel.shapes);
        if (plan.error) { setStatus(plan.error, "err"); return; }
        var ids = Object.keys(plan.changes);
        if (!ids.length) { setStatus("이미 그 값이에요.", "ok"); return; }

        var byId = {};
        sel.shapes.forEach(function (s) { byId[s.id] = s; });
        var snap = ids.map(function (id) { return snapshotOf(byId[id], sel.slideId); });

        sel.coll.items.forEach(function (item) { var c = plan.changes[item.id]; if (c) applyChange(item, c); });
        await ctx.sync();

        pushHistory({ label: label, snap: snap });
        setStatus((plan.msg || label) + " · ⌘Z로 되돌리기", "ok");
      });
    } catch (e) {
      setStatus("실행 실패: " + (e && e.message ? e.message : e), "err");
    } finally {
      busy = false;
      refreshSelection();
    }
  }

  function pushHistory(entry) {
    history.push(entry);
    if (history.length > 50) history.shift();
    future = [];
  }

  // 되돌리기/다시하기: 기록해 둔 값으로 되돌리고, 지금 값은 반대쪽 스택에
  async function restore(from, to, word) {
    var entry = from.pop();
    if (!entry) { setStatus(word + "할 작업이 없어요.", ""); return; }
    if (!state.inOffice) return demoRestore(entry, to, word);
    if (busy) { from.push(entry); return; }
    busy = true;
    try {
      await PowerPoint.run(async function (ctx) {
        var fallback = null;
        var rows = entry.snap.map(function (r) {
          var sh;
          if (r.slideId) sh = ctx.presentation.slides.getItem(r.slideId).shapes.getItem(r.id);
          else { fallback = fallback || ctx.presentation.getSelectedShapes(); sh = fallback.getItem(r.id); }
          sh.load("left,top,width,height" + (state.api110 ? ",rotation" : ""));
          var adjRes = (r.adj !== undefined && state.api110) ? sh.adjustments.get(0) : null;
          return { r: r, sh: sh, adjRes: adjRes };
        });
        await ctx.sync();
        var back = rows.map(function (x) {
          return { slideId: x.r.slideId, id: x.r.id, left: x.sh.left, top: x.sh.top, width: x.sh.width, height: x.sh.height,
            rotation: state.api110 ? x.sh.rotation : null, adj: x.adjRes ? toFrac(x.adjRes.value) : undefined };
        });
        rows.forEach(function (x) { applyChange(x.sh, x.r); });
        await ctx.sync();
        to.push({ label: entry.label, snap: back });
        setStatus(word + ": " + entry.label, "ok");
      });
    } catch (e) {
      setStatus(word + " 실패 (오브젝트가 삭제됐거나 다른 슬라이드일 수 있어요)", "err");
    } finally {
      busy = false;
      refreshSelection();
    }
  }
  function undo() { return restore(history, future, "되돌리기"); }
  function redo() { return restore(future, history, "다시 실행"); }

  // ---------- 작업 정의 ----------
  function movesToChanges(res) {
    if (res.error) return res;
    return { changes: res.moves, msg: Object.keys(res.moves).length + "개 이동 · " + refLabel(res.ref) + " 기준" };
  }

  function opAlign(edge) {
    return mutate("정렬", function (shapes) { return movesToChanges(Core.align(shapes, edge, coreOpts())); });
  }
  function opDistribute(axis, gapPt) {
    return mutate(gapPt === null ? "간격 똑같이" : "간격 지정", function (shapes) {
      return movesToChanges(Core.distribute(shapes, axis, coreOpts({ gap: gapPt })));
    });
  }
  function opSize(dim, pt) {
    if (!(pt > 0)) return setStatus("0보다 큰 값을 넣으세요.", "err");
    return mutate("크기 변경", function (shapes) {
      var ch = {};
      shapes.forEach(function (s) {
        var w = s.width, h = s.height;
        if (dim === "w") { if (state.lock && w > 0) h = h * pt / w; w = pt; }
        else { if (state.lock && h > 0) w = w * pt / h; h = pt; }
        if (Math.abs(w - s.width) > 0.005 || Math.abs(h - s.height) > 0.005) ch[s.id] = { width: w, height: h };
      });
      return { changes: ch, msg: "크기 변경" };
    });
  }
  // X/Y: 선택 전체(바운딩 박스)의 왼쪽 위 기준으로 이동
  function opPos(axis, pt) {
    return mutate("위치 이동", function (shapes) {
      var box = Core.bbox(shapes);
      var d = pt - (axis === "x" ? box.left : box.top);
      var ch = {};
      if (Math.abs(d) > 0.005) shapes.forEach(function (s) {
        ch[s.id] = axis === "x" ? { left: s.left + d } : { top: s.top + d };
      });
      return { changes: ch, msg: "위치 이동" };
    });
  }

  function opRotate(deg) {
    if (!state.api110) return setStatus("이 PowerPoint 버전은 회전을 지원하지 않아요.", "err");
    return mutate("회전", function (shapes) {
      var ch = {};
      shapes.forEach(function (s) { if (Math.abs(normRot(s.rotation || 0) - normRot(deg)) > 0.005) ch[s.id] = { rotation: deg }; });
      return { changes: ch, msg: "회전 " + deg + "°" };
    });
  }
  function opRadius(pt) {
    if (!state.api110) return setStatus("이 PowerPoint 버전은 라운드 조절을 지원하지 않아요.", "err");
    if (!(pt >= 0)) return;
    return mutate("라운드", function (shapes) {
      var targets = shapes.filter(function (s) { return s.adj !== undefined; });
      if (!targets.length) return { error: "모서리 조절점이 있는 도형(둥근 사각형 등)을 선택하세요." };
      var ch = {};
      targets.forEach(function (s) {
        var m = Math.min(s.width, s.height);
        if (m > 0) ch[s.id] = { adj: Math.min(0.5, pt / m) };
      });
      return { changes: ch, msg: "라운드 " + targets.length + "개 적용" };
    }, true);
  }

  function refLabel(kind) { return { selection: "선택 영역", key: "기준 오브젝트", slide: "슬라이드" }[kind] || ""; }

  // ---------- UI ----------
  var statusTimer = null;
  function setStatus(msg, cls) {
    var el = $("#status");
    el.textContent = msg; el.className = cls || "";
    clearTimeout(statusTimer);
    statusTimer = setTimeout(function () { el.textContent = ""; el.className = ""; }, 4000);
  }

  function setMode(mode) { state.mode = mode; save("mode", mode); render(); }

  function setField(el, valuePt, formatter, mixedPlaceholder, basePlaceholder, disabled) {
    el.disabled = !!disabled;
    if (document.activeElement === el) return;   // 입력 중엔 덮어쓰지 않음
    var v = valuePt === null || valuePt === undefined ? "" : (formatter ? formatter(valuePt) : fmt(valuePt));
    el.value = v;
    el.placeholder = (v === "" && mixedPlaceholder && state.shapes.length) ? mixedPlaceholder : (basePlaceholder || "");
  }

  function render() {
    var shapes = state.shapes, n = shapes.length;
    $$(".seg button").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.mode === state.mode)); });
    $("#count").textContent = "선택 " + n + "개";
    $("#unit").value = state.unit;

    var hint = "";
    if (n === 1) hint = "1개 → 슬라이드 기준";
    else if (state.mode === "key" && n > 1) hint = "★ 오브젝트는 고정";
    $("#refHint").textContent = hint;

    // 간격
    var gh = Core.measureGap(shapes, "h"), gv = Core.measureGap(shapes, "v");
    setField($("#gapH"), gh && !gh.mixed ? gh.value : null, null, null, "간격", n < 2);
    setField($("#gapV"), gv && !gv.mixed ? gv.value : null, null, null, "간격", n < 2);

    // 모양
    var box = n ? Core.bbox(shapes) : null;
    setField($("#posX"), box ? box.left : null, null, "", "", !n);
    setField($("#posY"), box ? box.top : null, null, "", "", !n);
    setField($("#sizeW"), Core.common(shapes.map(function (s) { return s.width; })), null, "여러 값", "", !n);
    setField($("#sizeH"), Core.common(shapes.map(function (s) { return s.height; })), null, "여러 값", "", !n);
    $("#lock").setAttribute("aria-pressed", String(state.lock));
    var rot = Core.common(shapes.map(function (s) { return s.rotation === null ? null : normRot(s.rotation); }));
    setField($("#rot"), rot, function (v) { return String(parseFloat(v.toFixed(1))); }, "여러 값", "", !n || !state.api110);
    $("#rotPreset").disabled = !n || !state.api110;

    var rs = shapes.filter(function (s) { return s.adj !== undefined; });
    var radii = rs.map(function (s) { return toFrac(s.adj) * Math.min(s.width, s.height); });
    setField($("#radius"), Core.common(radii), null, "여러 값", "", !rs.length || !state.api110);
    $("#roundNote").textContent = !state.api110 ? "이 PowerPoint 버전은 지원하지 않아요."
      : !n ? "도형을 선택하세요."
      : rs.length ? "둥근 모서리 도형 " + rs.length + "개에 적용돼요."
      : "라운드를 넣을 수 없는 도형이에요.";

    renderKeys();

    // 선택 순서
    var key = effectiveKey();
    var byId = {};
    shapes.forEach(function (s) { byId[s.id] = s; });
    var list = $("#list");
    list.innerHTML = "";
    var ordered = state.order.filter(function (id) { return byId[id]; });
    if (!ordered.length) {
      var li0 = document.createElement("li");
      li0.className = "empty";
      li0.textContent = "PowerPoint에서 오브젝트를 선택하세요.";
      list.appendChild(li0);
    }
    ordered.forEach(function (id, i) {
      var s = byId[id];
      var li = document.createElement("li");
      var isKey = id === key && n > 1;
      if (isKey) li.className = "key";
      li.innerHTML = '<span class="n"></span><span class="nm"></span>' + (isKey ? '<span class="tag">★ 기준</span>' : "");
      li.querySelector(".n").textContent = i + 1;
      li.querySelector(".nm").textContent = s.name || "도형";
      li.title = "이 오브젝트를 기준으로";
      li.addEventListener("click", function () {
        state.manualKey = id;
        if (state.mode !== "key") setMode("key"); else render();
        setStatus("'" + (s.name || "도형") + "'을(를) 기준으로 지정했어요.", "ok");
      });
      list.appendChild(li);
    });

    $("#slidePreset").value = state.slidePreset;
    $("#customSize").hidden = state.slidePreset !== "custom";
    var sz = slideSize();
    $("#slideInfo").textContent = (state.slidePreset === "auto" && !state.detected ? "자동 감지 불가 → 16:9로 계산 · " : "") +
      "현재 " + Core.fromPt(sz.width, "cm").toFixed(2) + " × " + Core.fromPt(sz.height, "cm").toFixed(2) + " cm";
  }

  function commitInput(el) {
    var v = readNum(el);
    if (v === null) { render(); return; }
    var kind = el.dataset.kind;
    if (kind === "gap") return opDistribute(el.dataset.axis, Core.toPt(v, state.unit));
    if (kind === "size") return opSize(el.dataset.dim, Core.toPt(v, state.unit));
    if (kind === "rot") return opRotate(v);
    if (kind === "pos") return opPos(el.dataset.axis, Core.toPt(v, state.unit));
    if (kind === "radius") return opRadius(Core.toPt(v, state.unit));
  }

  function bindUI() {
    $$(".seg button").forEach(function (b) { b.addEventListener("click", function () { setMode(b.dataset.mode); }); });
    $$("[data-align]").forEach(function (b) { b.addEventListener("click", function () { opAlign(b.dataset.align); }); });
    $$("[data-dist]").forEach(function (b) { b.addEventListener("click", function () { opDistribute(b.dataset.dist, null); }); });

    $$("input[data-kind]").forEach(function (el) {
      el.addEventListener("focus", function () { el.dataset.dirty = "0"; el.dataset.orig = el.value; });
      el.addEventListener("input", function () { el.dataset.dirty = "1"; });
      el.addEventListener("keydown", function (e) {
        if (e.key === "Enter") { e.preventDefault(); el.blur(); }
        else if (e.key === "Escape") { el.value = el.dataset.orig || ""; el.dataset.dirty = "0"; el.blur(); }
      });
      el.addEventListener("change", function () {
        if (el.dataset.dirty !== "1" && document.activeElement === el) return;
        el.dataset.dirty = "0";
        commitInput(el);
      });
    });

    $("#lock").addEventListener("click", function () { state.lock = !state.lock; save("lock", state.lock ? "1" : "0"); render(); });
    $("#rotPreset").addEventListener("change", function (e) {
      var v = e.target.value; e.target.value = "";
      if (v !== "") opRotate(+v);
    });
    $("#roundDetails").open = load("roundOpen", "0") === "1";
    $("#roundDetails").addEventListener("toggle", function (e) { save("roundOpen", e.target.open ? "1" : "0"); });
    $("#keysOn").addEventListener("change", function (e) {
      state.keysOn = e.target.checked; save("keysOn", state.keysOn ? "1" : "0"); cancelCapture(); renderKeys();
      setStatus(state.keysOn ? "단축키를 켰어요." : "단축키를 껐어요.", "ok");
    });
    $("#unit").addEventListener("change", function (e) { state.unit = e.target.value; save("unit", state.unit); render(); });
    $("#slidePreset").addEventListener("change", function (e) { state.slidePreset = e.target.value; save("slidePreset", state.slidePreset); render(); });
    $("#sw").value = state.customW; $("#sh").value = state.customH;
    $("#sw").addEventListener("input", function (e) { state.customW = e.target.value; save("customW", state.customW); render(); });
    $("#sh").addEventListener("input", function (e) { state.customH = e.target.value; save("customH", state.customH); render(); });

    // 리모컨을 눌러 포커스가 패널에 있어도 ⌘Z / Ctrl+Z 가 바로 먹도록
    document.addEventListener("keydown", function (e) {
      if (capture) { onCaptureKey(e); return; }
      // 패널에 포커스가 있을 때도 설정한 단축키가 먹도록
      var a0 = document.activeElement;
      if (state.keysOn && !(a0 && a0.tagName === "INPUT")) {
        var combo = comboFromEvent(e);
        if (combo.ok) {
          for (var id in state.keys) if (state.keys[id] === combo.text) { e.preventDefault(); runKeyAction(id); return; }
        }
      }
      var mod = e.metaKey || e.ctrlKey;
      if (!mod) return;
      var k = e.key.toLowerCase();
      if (k !== "z" && k !== "y") return;
      var a = document.activeElement;
      if (a && a.tagName === "INPUT" && a.dataset.dirty === "1") return; // 입력 중인 글자는 기본 되돌리기
      e.preventDefault();
      if (k === "y" || e.shiftKey) redo(); else undo();
    }, true);

    // 버튼을 누른 뒤 포커스가 버튼에 남아 있게 해서 단축키가 패널로 오게
    $$("button").forEach(function (b) { b.addEventListener("mousedown", function () { if (document.activeElement && document.activeElement.tagName === "INPUT") document.activeElement.blur(); }); });
  }

  // ---------- 리본 버튼 (공유 런타임) ----------
  function registerRibbon() {
    if (!Office.actions || !Office.actions.associate) return;
    function wrap(fn) { return function (event) { Promise.resolve(fn()).finally(function () { event.completed(); }); }; }
    ["left", "hcenter", "right", "top", "vcenter", "bottom"].forEach(function (edge) {
      Office.actions.associate("align_" + edge, wrap(function () { return opAlign(edge); }));
    });
    Office.actions.associate("dist_h", wrap(function () { return opDistribute("h", null); }));
    Office.actions.associate("dist_v", wrap(function () { return opDistribute("v", null); }));
    Office.actions.associate("ref_selection", wrap(function () { setMode("selection"); }));
    Office.actions.associate("ref_key", wrap(function () { setMode("key"); }));
    Office.actions.associate("ref_slide", wrap(function () { setMode("slide"); }));
    KEY_ACTIONS.forEach(function (k) {
      Office.actions.associate(k.id, wrap(function () { return runKeyAction(k.id); }));
    });
  }

  // ---------- 미리보기 모드 (PowerPoint 밖에서 열었을 때) ----------
  function demoMutate(label, planner) {
    var plan = planner(state.shapes);
    if (plan.error) return setStatus(plan.error, "err");
    var snap = [];
    state.shapes.forEach(function (s) {
      var c = plan.changes[s.id];
      if (!c) return;
      snap.push(snapshotOf(s, null));
      demoApply(s, c);
    });
    pushHistory({ label: label, snap: snap });
    render();
    setStatus("미리보기: " + (plan.msg || label), "ok");
  }
  function demoApply(s, c) {
    ["left", "top", "width", "height", "rotation"].forEach(function (k) { if (c[k] !== undefined && c[k] !== null) s[k] = c[k]; });
    if (c.adj !== undefined && c.adj !== null) s.adj = fromFrac(c.adj);
  }
  function demoRestore(entry, to, word) {
    var back = [];
    entry.snap.forEach(function (r) {
      var s = state.shapes.filter(function (x) { return x.id === r.id; })[0];
      if (!s) return;
      back.push(snapshotOf(s, null));
      demoApply(s, r);
    });
    to.push({ label: entry.label, snap: back });
    render();
    setStatus(word + ": " + entry.label, "ok");
  }

  // ---------- 단축키 ----------
  var isMac = /Mac/i.test(navigator.platform || navigator.userAgent);
  var KEY_ACTIONS = [
    { id: "key_left", edge: "left", name: "왼쪽 맞춤", n: 1 },
    { id: "key_hcenter", edge: "hcenter", name: "가로 가운데", n: 2 },
    { id: "key_right", edge: "right", name: "오른쪽 맞춤", n: 3 },
    { id: "key_vcenter", edge: "vcenter", name: "세로 가운데", n: 4 },
    { id: "key_top", edge: "top", name: "위쪽 맞춤", n: 5 },
    { id: "key_bottom", edge: "bottom", name: "아래쪽 맞춤", n: 6 },
  ];
  function defaultCombo(k) { return (isMac ? "Command+Option+" : "Ctrl+Alt+") + k.n; }
  var capture = null; // { id, text }

  function runKeyAction(id) {
    if (!state.keysOn) return;
    var k = KEY_ACTIONS.filter(function (x) { return x.id === id; })[0];
    if (k) return opAlign(k.edge);
  }

  function initKeys() {
    var saved = {};
    try { saved = JSON.parse(load("keys", "{}")) || {}; } catch (e) {}
    KEY_ACTIONS.forEach(function (k) { state.keys[k.id] = saved[k.id] || defaultCombo(k); });
    state.keysApi = !!(typeof Office !== "undefined" && Office.actions && Office.actions.replaceShortcuts &&
      Office.context && Office.context.requirements && Office.context.requirements.isSetSupported("KeyboardShortcuts", "1.1"));
    if (state.keysApi && Office.actions.getShortcuts) {
      Office.actions.getShortcuts().then(function (m) {
        KEY_ACTIONS.forEach(function (k) { if (m && m[k.id]) state.keys[k.id] = m[k.id]; });
        renderKeys();
      }).catch(function () {});
    }
    renderKeys();
  }

  // 키 이벤트 → "Command+Option+1" 같은 조합 문자열
  function comboFromEvent(e) {
    var parts = [];
    if (isMac) {
      if (e.metaKey) parts.push("Command");
      if (e.altKey) parts.push("Option");
    } else {
      if (e.ctrlKey) parts.push("Ctrl");
      if (e.altKey) parts.push("Alt");
    }
    var hasMain = parts.length > 0;
    if (e.shiftKey) parts.push("Shift");
    var code = e.code || "", key = null, m;
    if ((m = /^Key([A-Z])$/.exec(code))) key = m[1];
    else if ((m = /^Digit(\d)$/.exec(code))) key = m[1];
    else if (code === "Minus") key = "-";
    var modOnly = /^(Meta|Alt|Control|Shift|OS)/.test(code);
    var text = parts.concat(key ? [key] : []).join("+");
    var err = null;
    if (modOnly) err = "";
    else if (!key) err = "영문자·숫자·- 키만 쓸 수 있어요. (F키는 PowerPoint가 막아 둬서 불가)";
    else if (!hasMain) err = isMac ? "⌘ 또는 ⌥ 키를 같이 눌러 주세요." : "Ctrl 또는 Alt 키를 같이 눌러 주세요.";
    return { ok: !err && err !== "", text: text, err: err, partial: modOnly };
  }

  function pretty(combo) {
    if (!combo) return "없음";
    if (!isMac) return combo;
    return combo.replace(/Command\+?/g, "⌘").replace(/Option\+?/g, "⌥").replace(/Shift\+?/g, "⇧")
      .replace(/Ctrl\+?/g, "⌃").replace(/Alt\+?/g, "⌥");
  }

  function startCapture(id) { capture = { id: id, text: "", err: "" }; renderKeys(); }
  function cancelCapture() { capture = null; }

  function onCaptureKey(e) {
    e.preventDefault(); e.stopPropagation();
    if (e.key === "Escape") { cancelCapture(); renderKeys(); return; }
    var c = comboFromEvent(e);
    if (c.partial) { capture.text = c.text; capture.err = ""; }
    else if (c.ok) { capture.text = c.text; capture.err = ""; capture.ready = true; }
    else { capture.text = c.text; capture.err = c.err; capture.ready = false; }
    renderKeys();
  }

  function confirmCapture() {
    if (!capture || !capture.ready) return;
    var id = capture.id, combo = capture.text;
    var dup = KEY_ACTIONS.filter(function (k) { return k.id !== id && state.keys[k.id] === combo; })[0];
    if (dup) { capture.err = "'" + dup.name + "'에 이미 쓰는 조합이에요."; capture.ready = false; renderKeys(); return; }
    var done = function () {
      state.keys[id] = combo;
      save("keys", JSON.stringify(state.keys));
      cancelCapture(); renderKeys();
      setStatus("단축키를 " + pretty(combo) + "(으)로 바꿨어요.", "ok");
    };
    if (!state.keysApi) { done(); return; }
    var map = {}; map[id] = combo;
    Office.actions.replaceShortcuts(map).then(done).catch(function () {
      if (capture) { capture.err = "PowerPoint가 이 조합을 받아주지 않았어요. 다른 조합을 눌러 주세요."; capture.ready = false; renderKeys(); }
    });
  }

  function renderKeys() {
    var ul = $("#keyList");
    if (!ul) return;
    $("#keysOn").checked = state.keysOn;
    ul.className = "keys" + (state.keysOn ? "" : " off");
    ul.innerHTML = "";
    KEY_ACTIONS.forEach(function (k) {
      var li = document.createElement("li");
      var icon = document.querySelector('[data-align="' + k.edge + '"] svg');
      var rec = capture && capture.id === k.id;
      li.innerHTML = (icon ? icon.outerHTML : "<span></span>") + '<span class="kn"></span><span class="kbd"></span><span class="acts"></span>';
      li.querySelector(".kn").textContent = k.name;
      var kbd = li.querySelector(".kbd");
      kbd.textContent = rec ? (capture.text ? pretty(capture.text) : "키 입력…") : pretty(state.keys[k.id]);
      if (rec) kbd.className = "kbd rec";
      var acts = li.querySelector(".acts");
      if (rec) {
        var ok = document.createElement("button"); ok.className = "sbtn pri"; ok.textContent = "확인";
        ok.disabled = !capture.ready; ok.style.opacity = capture.ready ? "1" : ".4";
        ok.addEventListener("click", confirmCapture);
        var no = document.createElement("button"); no.className = "sbtn"; no.textContent = "취소"; no.style.marginLeft = "4px";
        no.addEventListener("click", function () { cancelCapture(); renderKeys(); });
        acts.appendChild(ok); acts.appendChild(no);
      } else {
        var set = document.createElement("button"); set.className = "sbtn"; set.textContent = "단축키 설정";
        set.disabled = !state.keysOn;
        set.addEventListener("click", function () { startCapture(k.id); });
        acts.appendChild(set);
      }
      ul.appendChild(li);
    });
    var note = "";
    if (capture && capture.err) note = capture.err;
    else if (capture) note = "원하는 조합을 누르고 '확인'. Esc로 취소.";
    else if (!state.keysOn) note = "단축키가 꺼져 있어요.";
    else if (state.inOffice && !state.keysApi) note = "이 PowerPoint 버전은 단축키 변경을 지원하지 않아 기본 단축키만 동작할 수 있어요.";
    $("#keysNote").textContent = note;
  }

  // ---------- 시작 ----------
  bindUI();
  render();

  if (typeof Office === "undefined") startDemo();
  else {
    Office.onReady(function (info) {
      if (info.host === Office.HostType.PowerPoint) {
        state.inOffice = true;
        state.api110 = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
        registerRibbon();
        initKeys();
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, scheduleRefresh);
        refreshSelection();
        // 드래그로 크기·위치를 바꿔도 값이 따라오도록 주기적으로 갱신 (입력 중엔 멈춤)
        setInterval(function () {
          var a = document.activeElement;
          if (document.visibilityState === "visible" && !(a && a.tagName === "INPUT")) refreshSelection();
        }, 1000);
      } else startDemo();
    });
  }

  function startDemo() {
    initKeys();
    state.api110 = true;
    state.adjScale = "frac";
    applySelection([
      { id: "d1", name: "타이틀 박스", left: 80, top: 60, width: 300, height: 60, rotation: 0, adj: 0.1667 },
      { id: "d2", name: "제품 이미지", left: 420, top: 150, width: 200, height: 200, rotation: 0 },
      { id: "d3", name: "CTA 버튼", left: 700, top: 380, width: 140, height: 44, rotation: 15, adj: 0.5 },
    ]);
    setStatus("PowerPoint 밖에서 열려 미리보기 모드로 동작 중이에요.", "");
  }
})();
