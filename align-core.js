/*
 * 정렬 리모컨 — 계산 로직 (PowerPoint와 무관한 순수 함수)
 * 모든 좌표/크기 단위는 pt(포인트). PowerPoint 도형의 left/top/width/height와 같은 단위.
 *
 * shape: { id, left, top, width, height }
 * opts : { mode: 'selection' | 'key' | 'slide', keyId, slide: { width, height }, gap: number | null }
 */
(function (root) {
  "use strict";

  var EPS = 0.005;

  function round(v) {
    return Math.round(v * 100) / 100;
  }

  function bbox(list) {
    var l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
    list.forEach(function (s) {
      l = Math.min(l, s.left);
      t = Math.min(t, s.top);
      r = Math.max(r, s.left + s.width);
      b = Math.max(b, s.top + s.height);
    });
    return { left: l, top: t, width: r - l, height: b - t };
  }

  // 기준 박스 결정
  // - 슬라이드 모드이거나 오브젝트가 1개뿐이면 슬라이드 기준
  // - 기준(키) 모드면 키 오브젝트 (키 오브젝트 자신은 움직이지 않음)
  // - 그 외엔 선택 영역 전체의 바운딩 박스
  function resolveRef(shapes, opts) {
    if (opts.mode === "slide" || shapes.length === 1) {
      return {
        kind: "slide",
        box: { left: 0, top: 0, width: opts.slide.width, height: opts.slide.height },
      };
    }
    if (opts.mode === "key") {
      var key = null;
      for (var i = 0; i < shapes.length; i++) if (shapes[i].id === opts.keyId) key = shapes[i];
      if (!key) key = shapes[0];
      return { kind: "key", box: key, keyId: key.id };
    }
    return { kind: "selection", box: bbox(shapes) };
  }

  function align(shapes, edge, opts) {
    if (!shapes.length) return { error: "선택된 오브젝트가 없어요." };
    var ref = resolveRef(shapes, opts);
    var B = ref.box;
    var moves = {};
    shapes.forEach(function (s) {
      if (ref.keyId === s.id) return;
      var left = s.left, top = s.top;
      switch (edge) {
        case "left": left = B.left; break;
        case "hcenter": left = B.left + B.width / 2 - s.width / 2; break;
        case "right": left = B.left + B.width - s.width; break;
        case "top": top = B.top; break;
        case "vcenter": top = B.top + B.height / 2 - s.height / 2; break;
        case "bottom": top = B.top + B.height - s.height; break;
        default: return;
      }
      left = round(left); top = round(top);
      if (Math.abs(left - s.left) > EPS || Math.abs(top - s.top) > EPS) {
        moves[s.id] = { left: left, top: top };
      }
    });
    return { moves: moves, ref: ref.kind, keyId: ref.keyId || null };
  }

  // 간격 분배
  // axis: 'h'(가로) | 'v'(세로)
  // - gap 값이 있으면: 그 간격으로 나란히 배치. 기준(키) 모드면 키 오브젝트를 고정하고 양옆으로, 아니면 가장 앞(왼쪽/위) 오브젝트 고정
  // - gap 값이 없으면: 양 끝 오브젝트는 그대로 두고 사이 간격을 똑같이 (3개 이상)
  //   슬라이드 모드면 슬라이드 폭/높이 안에서 바깥 여백까지 똑같이
  function distribute(shapes, axis, opts) {
    var P = axis === "h" ? { pos: "left", size: "width", slide: "width" } : { pos: "top", size: "height", slide: "height" };
    var n = shapes.length;
    var hasGap = typeof opts.gap === "number" && isFinite(opts.gap);

    var sorted = shapes.slice().sort(function (a, b) {
      var d = (a[P.pos] + a[P.size] / 2) - (b[P.pos] + b[P.size] / 2);
      return Math.abs(d) > EPS ? d : 0;
    });

    var pos = new Array(n);
    var i;

    if (hasGap) {
      if (n < 2) return { error: "간격을 지정하려면 오브젝트를 2개 이상 선택하세요." };
      var ai = 0;
      if (opts.mode === "key") {
        for (i = 0; i < n; i++) if (sorted[i].id === opts.keyId) ai = i;
      }
      pos[ai] = sorted[ai][P.pos];
      for (i = ai + 1; i < n; i++) pos[i] = pos[i - 1] + sorted[i - 1][P.size] + opts.gap;
      for (i = ai - 1; i >= 0; i--) pos[i] = pos[i + 1] - opts.gap - sorted[i][P.size];
    } else {
      var sum = 0;
      sorted.forEach(function (s) { sum += s[P.size]; });
      if (opts.mode === "slide") {
        var g0 = (opts.slide[P.slide] - sum) / (n + 1);
        var p = g0;
        for (i = 0; i < n; i++) { pos[i] = p; p += sorted[i][P.size] + g0; }
      } else {
        if (n < 3) return { error: "간격을 똑같이 나누려면 3개 이상 선택하세요. 2개면 간격 칸에 값을 입력하세요." };
        var box = bbox(sorted);
        var start = axis === "h" ? box.left : box.top;
        var span = axis === "h" ? box.width : box.height;
        var g = (span - sum) / (n - 1);
        var q = start;
        for (i = 0; i < n; i++) { pos[i] = q; q += sorted[i][P.size] + g; }
      }
    }

    var moves = {};
    for (i = 0; i < n; i++) {
      var s = sorted[i];
      var v = round(pos[i]);
      if (Math.abs(v - s[P.pos]) > EPS) {
        moves[s.id] = axis === "h" ? { left: v, top: s.top } : { left: s.left, top: v };
      }
    }
    return { moves: moves, ref: opts.mode };
  }

  // 선택 순서 추적: 이전 순서 + 현재 선택 id 목록 → 새 순서
  function updateOrder(prevOrder, currentIds) {
    var next = prevOrder.filter(function (id) { return currentIds.indexOf(id) !== -1; });
    currentIds.forEach(function (id) { if (next.indexOf(id) === -1) next.push(id); });
    return next;
  }

  // 현재 간격 측정: 위치 순으로 정렬한 뒤 이웃 간 간격. 모두 같으면 값, 다르면 mixed
  function measureGap(shapes, axis) {
    if (shapes.length < 2) return null;
    var P = axis === "h" ? { pos: "left", size: "width" } : { pos: "top", size: "height" };
    var sorted = shapes.slice().sort(function (a, b) {
      return (a[P.pos] + a[P.size] / 2) - (b[P.pos] + b[P.size] / 2);
    });
    var gaps = [];
    for (var i = 1; i < sorted.length; i++) {
      gaps.push(sorted[i][P.pos] - (sorted[i - 1][P.pos] + sorted[i - 1][P.size]));
    }
    var first = gaps[0];
    var same = gaps.every(function (g) { return Math.abs(g - first) < 0.05; });
    return same ? { value: first, mixed: false } : { value: null, mixed: true };
  }

  // 여러 오브젝트의 같은 속성이 모두 같으면 그 값, 아니면 null
  function common(values) {
    if (!values.length) return null;
    var f = values[0];
    if (f === null || f === undefined) return null;
    return values.every(function (v) { return v !== null && v !== undefined && Math.abs(v - f) < 0.005; }) ? f : null;
  }

  var UNITS = { pt: 1, cm: 72 / 2.54, mm: 72 / 25.4, px: 0.75 };
  function toPt(value, unit) { return value * (UNITS[unit] || 1); }
  function fromPt(value, unit) { return value / (UNITS[unit] || 1); }

  var api = { align: align, distribute: distribute, bbox: bbox, updateOrder: updateOrder, measureGap: measureGap, common: common, toPt: toPt, fromPt: fromPt };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.AlignCore = api;
})(this);
