/* 정렬 리모컨 — PowerPoint 연동 + 패널 UI */
(function () {
  "use strict";
  var Core = window.AlignCore;

  // ---------- 상태 ----------
  var state = {
    mode: load("mode", "selection"),   // selection | key | slide
    unit: load("unit", "cm"),
    slidePreset: load("slidePreset", "auto"),
    customW: load("customW", ""),
    customH: load("customH", ""),
    order: [],        // 선택한 순서대로 쌓이는 id 목록
    manualKey: null,  // 목록에서 직접 찍은 기준 오브젝트
    shapes: [],       // 현재 선택된 도형 {id,name,left,top,width,height}
    detected: null,   // 자동 감지된 슬라이드 크기 (pt)
    inOffice: false,
  };

  function load(k, d) { try { var v = localStorage.getItem("alignRemote." + k); return v === null ? d : v; } catch (e) { return d; } }
  function save(k, v) { try { localStorage.setItem("alignRemote." + k, v); } catch (e) {} }

  var $ = function (s) { return document.querySelector(s); };
  var $$ = function (s) { return Array.prototype.slice.call(document.querySelectorAll(s)); };

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

  function gapPt() {
    var raw = $("#gap").value.trim();
    if (raw === "") return null;
    var v = parseFloat(raw);
    return isFinite(v) ? Core.toPt(v, state.unit) : null;
  }

  // ---------- PowerPoint ----------
  function readSelection(ctx) {
    var coll = ctx.presentation.getSelectedShapes();
    coll.load("items/id,items/name,items/left,items/top,items/width,items/height");
    return coll;
  }

  function toPlain(items) {
    return items.map(function (s) {
      return { id: s.id, name: s.name, left: s.left, top: s.top, width: s.width, height: s.height };
    });
  }

  async function detectSlideSize(ctx) {
    if (state.detected) return;
    try {
      if (Office.context.requirements.isSetSupported("PowerPointApi", "1.10") && ctx.presentation.pageSetup) {
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
    if (!state.inOffice) return;
    try {
      await PowerPoint.run(async function (ctx) {
        var coll = readSelection(ctx);
        await ctx.sync();
        await detectSlideSize(ctx);
        applySelection(toPlain(coll.items));
      });
    } catch (e) {
      applySelection([]);
    }
  }

  var refreshTimer = null;
  function scheduleRefresh() {
    clearTimeout(refreshTimer);
    refreshTimer = setTimeout(refreshSelection, 60);
  }

  // op: { type: 'align', edge } | { type: 'dist', axis }
  async function runOp(op) {
    if (!state.inOffice) return demoRun(op);
    try {
      await PowerPoint.run(async function (ctx) {
        var coll = readSelection(ctx);
        await ctx.sync();
        await detectSlideSize(ctx);
        var shapes = toPlain(coll.items);
        // 패널이 선택 변경 이벤트를 놓쳤을 때도 순서가 맞도록 한 번 더 반영
        state.shapes = shapes;
        state.order = Core.updateOrder(state.order, shapes.map(function (s) { return s.id; }));

        var result = compute(op, shapes);
        if (result.error) { setStatus(result.error, "err"); render(); return; }

        var n = 0;
        coll.items.forEach(function (s) {
          var m = result.moves[s.id];
          if (!m) return;
          s.left = m.left; s.top = m.top; n++;
        });
        await ctx.sync();
        shapes.forEach(function (s) { var m = result.moves[s.id]; if (m) { s.left = m.left; s.top = m.top; } });
        render();
        setStatus(n ? n + "개 이동 · " + refLabel(result.ref) + " 기준" : "이미 맞춰져 있어요.", "ok");
      });
    } catch (e) {
      setStatus("실행 실패: " + (e && e.message ? e.message : e), "err");
    }
  }

  function compute(op, shapes) {
    if (!shapes.length) return { error: "먼저 오브젝트를 선택하세요." };
    var opts = { mode: state.mode, keyId: effectiveKey(), slide: slideSize(), gap: gapPt() };
    return op.type === "align" ? Core.align(shapes, op.edge, opts) : Core.distribute(shapes, op.axis, opts);
  }

  function refLabel(kind) {
    return { selection: "선택 영역", key: "기준 오브젝트", slide: "슬라이드" }[kind] || "";
  }

  // ---------- UI ----------
  var statusTimer = null;
  function setStatus(msg, cls) {
    var el = $("#status");
    el.textContent = msg; el.className = cls || "";
    clearTimeout(statusTimer);
    statusTimer = setTimeout(function () { el.textContent = ""; el.className = ""; }, 4000);
  }

  function setMode(mode) {
    state.mode = mode; save("mode", mode); render();
  }

  function render() {
    $$(".seg button").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.mode === state.mode)); });
    $("#count").textContent = "선택 " + state.shapes.length + "개";

    var hint = "";
    if (state.shapes.length === 1) hint = "1개 → 슬라이드 기준";
    else if (state.mode === "key" && state.shapes.length > 1) hint = "★ 오브젝트는 고정";
    $("#refHint").textContent = hint;

    var key = effectiveKey();
    var byId = {};
    state.shapes.forEach(function (s) { byId[s.id] = s; });
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
      var isKey = id === key && state.shapes.length > 1;
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

    $("#unit").value = state.unit;
    $("#slidePreset").value = state.slidePreset;
    $("#customSize").hidden = state.slidePreset !== "custom";
    var sz = slideSize();
    var auto = state.slidePreset === "auto";
    $("#slideInfo").textContent = (auto && !state.detected ? "자동 감지 불가 → 16:9로 계산 · " : "") +
      "현재 " + Core.fromPt(sz.width, "cm").toFixed(2) + " × " + Core.fromPt(sz.height, "cm").toFixed(2) + " cm";
  }

  function bindUI() {
    $$(".seg button").forEach(function (b) { b.addEventListener("click", function () { setMode(b.dataset.mode); }); });
    $$("[data-align]").forEach(function (b) { b.addEventListener("click", function () { runOp({ type: "align", edge: b.dataset.align }); }); });
    $$("[data-dist]").forEach(function (b) { b.addEventListener("click", function () { runOp({ type: "dist", axis: b.dataset.dist }); }); });
    $("#unit").addEventListener("change", function (e) { state.unit = e.target.value; save("unit", state.unit); });
    $("#slidePreset").addEventListener("change", function (e) { state.slidePreset = e.target.value; save("slidePreset", state.slidePreset); render(); });
    $("#sw").value = state.customW; $("#sh").value = state.customH;
    $("#sw").addEventListener("input", function (e) { state.customW = e.target.value; save("customW", state.customW); render(); });
    $("#sh").addEventListener("input", function (e) { state.customH = e.target.value; save("customH", state.customH); render(); });
  }

  // ---------- 리본 버튼 (공유 런타임) ----------
  function registerRibbon() {
    if (!Office.actions || !Office.actions.associate) return;
    function wrap(fn) { return function (event) { Promise.resolve(fn()).finally(function () { event.completed(); }); }; }
    ["left", "hcenter", "right", "top", "vcenter", "bottom"].forEach(function (edge) {
      Office.actions.associate("align_" + edge, wrap(function () { return runOp({ type: "align", edge: edge }); }));
    });
    Office.actions.associate("dist_h", wrap(function () { return runOp({ type: "dist", axis: "h" }); }));
    Office.actions.associate("dist_v", wrap(function () { return runOp({ type: "dist", axis: "v" }); }));
    Office.actions.associate("ref_selection", wrap(function () { setMode("selection"); }));
    Office.actions.associate("ref_key", wrap(function () { setMode("key"); }));
    Office.actions.associate("ref_slide", wrap(function () { setMode("slide"); }));
  }

  // ---------- 미리보기 모드 (PowerPoint 밖에서 열었을 때) ----------
  function demoRun(op) {
    var r = compute(op, state.shapes);
    if (r.error) return setStatus(r.error, "err");
    state.shapes.forEach(function (s) { var m = r.moves[s.id]; if (m) { s.left = m.left; s.top = m.top; } });
    render();
    setStatus("미리보기: " + Object.keys(r.moves).length + "개 이동 · " + refLabel(r.ref) + " 기준", "ok");
  }

  // ---------- 시작 ----------
  bindUI();
  render();

  if (typeof Office === "undefined") {
    startDemo();
  } else {
    Office.onReady(function (info) {
      if (info.host === Office.HostType.PowerPoint) {
        state.inOffice = true;
        registerRibbon();
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, scheduleRefresh);
        refreshSelection();
      } else {
        startDemo();
      }
    });
  }

  function startDemo() {
    applySelection([
      { id: "d1", name: "타이틀 박스", left: 80, top: 60, width: 300, height: 60 },
      { id: "d2", name: "제품 이미지", left: 420, top: 150, width: 200, height: 200 },
      { id: "d3", name: "CTA 버튼", left: 700, top: 380, width: 140, height: 44 },
    ]);
    setStatus("PowerPoint 밖에서 열려 미리보기 모드로 동작 중이에요.", "");
  }
})();
