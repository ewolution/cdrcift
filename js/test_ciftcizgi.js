/*
  cdrcift - JavaScript cekirdek testleri (Node)
  ---------------------------------------------
  CorelDRAW'in COM nesne modelini taklit eden bir mock ile tum tarama
  boru hattini calistirir. VBA'da yapilamayan sey: algoritmayi gercekten
  kosturmak.

  Calistirma:  node js/test_ciftcizgi.js
*/

"use strict";

var path = require("path");
var CC = require(path.join(__dirname, "CiftCizgi.js"));
var T = CC._test;

var gecen = 0, kalan = 0, rapor = [];

function kontrol(kosul, aciklama) {
  if (kosul) { gecen++; rapor.push("[OK]   " + aciklama); }
  else { kalan++; rapor.push("[HATA] " + aciklama); }
}

// ---------------------------------------------------------------------
// CorelDRAW mock
// ---------------------------------------------------------------------

function coll(dizi) {
  return {
    Count: dizi.length,
    Item: function (i) { return dizi[i - 1] || null; }
  };
}

function liveColl(dizi) {
  return {
    get Count() { return dizi.length; },
    Item: function (i) { return dizi[i - 1] || null; }
  };
}

function seg(x1, y1, x2, y2, tip) {
  return {
    Type: tip === undefined ? 0 : tip,
    StartNode: { PositionX: x1, PositionY: y1 },
    EndNode: { PositionX: x2, PositionY: y2 },
    StartingControlPointX: x1, StartingControlPointY: y1,
    EndingControlPointX: x2, EndingControlPointY: y2
  };
}

function shape(id, segs, kapali, katman, z) {
  var sp = { Closed: !!kapali, Segments: coll(segs) };
  var crv = { SubPaths: coll([sp]), Segments: coll(segs), Nodes: coll([]) };
  return {
    Type: 3,
    StaticID: id,
    Name: "Nesne " + id,
    Layer: { Name: katman || "Layer 1" },
    Outline: { Width: 0, Color: { HexValue: "#000000" } },
    ZOrder: z === undefined ? 1 : z,
    DisplayCurve: crv,
    Curve: crv,
    PowerClip: null,
    Fill: { Type: 0 },
    _deleted: false,
    Delete: function () { this._deleted = true; }
  };
}

// Geometri kurma (CreateCurve / CreateSubPath / AppendLineSegment) mock'u
function curveBuilder() {
  var sps = [];

  function subPath(x, y) {
    var segs = [];
    var sp = {
      _closed: false,
      SegmentList: segs,
      Segments: liveColl(segs),
      AppendLineSegment: function (ex, ey) {
        var prev = segs.length ? segs[segs.length - 1].EndNode : { PositionX: x, PositionY: y };
        var s = {
          Type: 0,
          StartNode: { PositionX: prev.PositionX, PositionY: prev.PositionY },
          EndNode: { PositionX: ex, PositionY: ey }
        };
        segs.push(s);
        return s;
      },
      AppendCurveSegment: function (c1x, c1y, c2x, c2y, ex, ey) {
        var prev = segs.length ? segs[segs.length - 1].EndNode : { PositionX: x, PositionY: y };
        var s = {
          Type: 1,
          StartNode: { PositionX: prev.PositionX, PositionY: prev.PositionY },
          EndNode: { PositionX: ex, PositionY: ey },
          StartingControlPointX: c1x, StartingControlPointY: c1y,
          EndingControlPointX: c2x, EndingControlPointY: c2y
        };
        segs.push(s);
        return s;
      }
    };
    Object.defineProperty(sp, "Closed", {
      get: function () { return sp._closed; },
      set: function (v) { sp._closed = !!v; }
    });
    sps.push(sp);
    return sp;
  }

  return {
    SubPaths: liveColl(sps),
    CreateSubPath: subPath
  };
}
function cizgi(id, x1, y1, x2, y2, katman, z) {
  return shape(id, [seg(x1, y1, x2, y2, 0)], false, katman, z);
}

function dikdortgen(id, l, b, r, t, katman, z) {
  var segs = [
    seg(l, b, l, t, 0),
    seg(l, t, r, t, 0),
    seg(r, t, r, b, 0),
    seg(r, b, l, b, 0)
  ];
  return shape(id, segs, true, katman, z);
}

function kur(shapes) {
  var yeni = [];
  var katmanlar = {};

  function katman(ad) {
    if (!katmanlar[ad]) {
      katmanlar[ad] = {
        Name: ad,
        CreateCurve: function (crv) {
          var sh = {
            Type: 3, StaticID: 1000 + yeni.length, Name: "Yeni " + yeni.length,
            Layer: katmanlar[ad], Outline: { Width: 0, Color: { HexValue: "#000000" } },
            ZOrder: 0, DisplayCurve: crv, Curve: crv, PowerClip: null,
            _deleted: false,
            Delete: function () { this._deleted = true; }
          };
          yeni.push(sh);
          return sh;
        }
      };
    }
    return katmanlar[ad];
  }

  var i;
  for (i = 0; i < shapes.length; i++) {
    shapes[i].Layer = katman(shapes[i].Layer.Name);
  }

  var page = {
    Shapes: coll(shapes),
    Layers: coll([]),
    CreateLayer: function () { return null; }
  };

  var doc = {
    Unit: 3, ActivePage: page, FileName: "test.cdr", Dirty: false,
    CreateCurve: function () { return curveBuilder(); },
    BeginCommandGroup: function () { return true; },
    EndCommandGroup: function () { return true; },
    ClearSelection: function () { return true; },
    Undo: function () { return true; }
  };

  global.host = {
    ActiveDocument: doc,
    ActiveSelectionRange: null,
    CreateObject: undefined,
    CorelScriptTools: undefined
  };
  global.alert = function () {};
  global.confirm = function () { return false; };

  T.reset();
  var hata = T.tara();
  shapes._yeni = yeni;          // yeniden kurulan nesneler
  return hata;
}

// ---------------------------------------------------------------------
// Testler
// ---------------------------------------------------------------------

function testFixture() {
  var shapes = [
    cizgi(1, 10, 10, 10, 60, "Layer 1", 4),      // A
    cizgi(2, 10, 10, 10, 60, "Layer 1", 3),      // B: birebir kopya
    cizgi(3, 110, 10, 110, 60, "Layer 1", 2),    // C1
    cizgi(4, 110, 60, 110, 10, "Layer 1", 1),    // C2: ters yon
    dikdortgen(5, 200, 10, 240, 50, "Layer 1", 5),
    dikdortgen(6, 240, 10, 280, 50, "Layer 1", 6),   // ortak kenar x=240
    cizgi(7, 300, 10, 300, 60, "Layer 1", 7),
    cizgi(8, 300.3, 10, 300.3, 60, "Layer 1", 8),    // 0,3 mm kaymis
    cizgi(9, 320, 10, 320, 60, "Layer 1", 9),
    cizgi(10, 320.1, 10, 320.1, 60, "Layer 1", 10)   // 0,1 mm kaymis
  ];

  CC.AYAR.tolerans = 0.05;
  CC.AYAR.yakinTolerans = 0.5;
  CC.AYAR.minOrtusme = 0.5;
  CC.AYAR.kapsam = "sayfa";
  CC.AYAR.otomatikSil = true;
  CC.AYAR.boya = false;

  var hata = kur(shapes);
  kontrol(hata === "", "Tarama hatasiz calisti (" + hata + ")");
  kontrol(T.g.nesneler.length === 10, "10 nesne tarandi, gelen " + T.g.nesneler.length);

  var bul = T.g.bulgular;

  function bulVar(tur, x, y, pay) {
    pay = pay === undefined ? 1 : pay;
    var i;
    for (i = 0; i < bul.length; i++) {
      var f = bul[i];
      if (f.tur === tur && Math.abs(f.x - x) <= pay && Math.abs(f.y - y) <= pay) return f;
    }
    return null;
  }

  var tamAB = bulVar("TAM", 10, 35);
  kontrol(!!tamAB, "Birebir kopya (10,35) TAM bulundu");
  kontrol(!!tamAB && tamAB.eylem === "sil", "Birebir kopya otomatik silinebilir");

  kontrol(!!bulVar("TAM", 110, 35), "Ters yondeki kopya (110,35) TAM bulundu");

  var kismi = bulVar("KISMI", 240, 30, 2);
  kontrol(!!kismi, "Ortak kenar (240,30) KISMI bulundu");
  kontrol(!!kismi && kismi.eylem === "elle", "KISMI asla otomatik silinmez");

  kontrol(!!bulVar("YAKIN", 300.15, 35, 1.2), "0,3 mm kaymis cizgi YAKIN bulundu");
  kontrol(!bulVar("TAM", 320.05, 35, 2), "0,1 mm kaymis cizgi TAM sayilmadi");

  var silIhlali = bul.filter(function (f) {
    return (f.tur === "KISMI" || f.tur === "YAKIN") && f.eylem === "sil";
  });
  kontrol(silIhlali.length === 0, "KISMI/YAKIN icinde 'sil' eylemi yok");

  var sayim = {};
  bul.forEach(function (f) { sayim[f.tur] = (sayim[f.tur] || 0) + 1; });
  rapor.push("       sayim: " + JSON.stringify(sayim));

  kontrol(sayim.TAM === 2, "Tam olarak 2 TAM beklendi, gelen " + (sayim.TAM || 0));
  kontrol((sayim.YAKIN || 0) >= 2, "En az 2 YAKIN beklendi, gelen " + (sayim.YAKIN || 0));
}

function testImza() {
  CC.AYAR.tolerans = 0.05;
  CC.AYAR.yakinTolerans = 0.5;
  CC.AYAR.minOrtusme = 0.5;
  CC.AYAR.kapsam = "sayfa";
  CC.AYAR.otomatikSil = false;

  // ayni dikdortgen: ters yon + farkli baslangic dugumu
  var a = dikdortgen(1, 0, 0, 100, 50, "L", 1);
  var b = shape(2, [
    seg(100, 0, 100, 50, 0),
    seg(100, 50, 0, 50, 0),
    seg(0, 50, 0, 0, 0),
    seg(0, 0, 100, 0, 0)
  ], true, "L", 2);

  var crvA = a.Curve.SubPaths.Item(1);
  var ia = T.imza(crvA);
  kontrol(typeof ia === "string" && ia.length > 10, "Kapali yol imzasi uretildi");
  // Tam imza icin segment listesi gerektiginden dogrudan tara() uzerinden test edilir
  kur([a, b]);
  var tam = T.g.bulgular.filter(function (f) { return f.tur === "TAM"; });
  kontrol(tam.length === 1, "Kapali yol: ters yon + farkli baslangic TAM (" + tam.length + ")");

  // cizgi: ters yon
  kur([cizgi(1, 0, 0, 50, 0, "L", 1), cizgi(2, 50, 0, 0, 0, "L", 2)]);
  tam = T.g.bulgular.filter(function (f) { return f.tur === "TAM"; });
  kontrol(tam.length === 1, "Acik cizgi: ters yon TAM (" + tam.length + ")");

  // 0,01 mm kaymis dikdortgen TAM olmamali
  kur([dikdortgen(1, 0, 0, 100, 50, "L", 1), dikdortgen(2, 0.01, 0, 100.01, 50, "L", 2)]);
  tam = T.g.bulgular.filter(function (f) { return f.tur === "TAM"; });
  kontrol(tam.length === 0, "0,01 mm kaymis dikdortgen TAM sayilmadi (" + tam.length + ")");
}

function testBirim() {
  // inc belgesi: 1 inc = 25.4 mm
  CC.AYAR.tolerans = 0.05;
  CC.AYAR.yakinTolerans = 0.5;
  CC.AYAR.otomatikSil = true;
  var s1 = cizgi(1, 1, 1, 1, 2, "L", 2);   // 1 inc cizgi
  var s2 = cizgi(2, 1, 1, 1, 2, "L", 1);
  var page = { Shapes: coll([s1, s2]), Layers: coll([]), CreateLayer: function () { return null; } };
  global.host = { ActiveDocument: { Unit: 1, ActivePage: page, FileName: "inc.cdr" } };
  T.reset();
  var hata = T.tara();
  kontrol(hata === "", "Inc belgesi tarandi");
  kontrol(Math.abs(T.g.altYollar[0].uzunluk - 25.4) < 0.01,
    "1 inc = 25,4 mm donusumu (" + (T.g.altYollar[0] && T.g.altYollar[0].uzunluk.toFixed(3)) + ")");
}

function testKatmanFarki() {
  CC.AYAR.tolerans = 0.05;
  CC.AYAR.otomatikSil = true;
  kur([
    cizgi(1, 10, 10, 10, 60, "Kesim", 2),
    cizgi(2, 10, 10, 10, 60, "Baska", 1)
  ]);
  var tam = T.g.bulgular.filter(function (f) { return f.tur === "TAM"; });
  kontrol(tam.length === 1, "Farkli katmandaki kopya TAM bulundu");
  kontrol(tam.length === 1 && tam[0].eylem === "elle",
    "Farkli katmanda otomatik silinmez (eylem=" + (tam[0] && tam[0].eylem) + ")");
}

function testIzgara() {
  // rastgele segmentler: izgara ile kaba kuvvet ayni ciftleri bulmali
  CC.AYAR.tolerans = 0.05;
  CC.AYAR.yakinTolerans = 0.5;
  CC.AYAR.minOrtusme = 0.5;
  CC.AYAR.kapsam = "sayfa";
  CC.AYAR.otomatikSil = true;

  var rnd = (function () { var s = 12345; return function () { s = (s * 1103515245 + 12345) % 2147483648; return s / 2147483648; }; })();
  var shapes = [], id = 1, i;
  for (i = 0; i < 60; i++) {
    var x = rnd() * 200, y = rnd() * 200;
    var aci = [0, 30, 45, 90, 12.5][Math.floor(rnd() * 5)];
    var boy = [5, 20, 60, 150][Math.floor(rnd() * 4)];
    var r = aci * Math.PI / 180;
    shapes.push(cizgi(id++, x, y, x + Math.cos(r) * boy, y + Math.sin(r) * boy, "L", i));
  }
  // kopyalar + kaymis kopyalar
  for (i = 0; i < 15; i++) {
    var k = shapes[Math.floor(rnd() * 60)];
    var c = k.Curve.SubPaths.Item(1).Segments.Item(1);
    shapes.push(cizgi(id++, c.StartNode.PositionX, c.StartNode.PositionY,
      c.EndNode.PositionX, c.EndNode.PositionY, "L", 100 + i));
  }

  var hata = kur(shapes);
  kontrol(hata === "", "Rastgele set tarandi (" + shapes.length + " nesne)");

  // En az 15 birebir kopya bulunmali
  var tam = T.g.bulgular.filter(function (f) { return f.tur === "TAM"; });
  kontrol(tam.length >= 15, "Rastgele sette >=15 TAM beklendi, gelen " + tam.length);
}

function testOrtakKenar() {
  // --- zincir mantigi (saf) ---
  var z = T.zincirler([false, true, false, false], true);
  kontrol(z.length === 1 && z[0].idx.join(",") === "2,3,0" && !z[0].closed,
    "Kapali: 1 kenar cikinca 3'lu acik zincir");

  z = T.zincirler([true, false, true, false], true);
  kontrol(z.length === 2 && z[0].idx.join(",") === "1" && z[1].idx.join(",") === "3",
    "Kapali: karsi 2 kenar cikinca 2 ayri zincir");

  z = T.zincirler([true, true, true, true], true);
  kontrol(z.length === 0, "Kapali: tum kenarlar cikinca zincir yok");

  z = T.zincirler([false, false, false, false], true);
  kontrol(z.length === 1 && z[0].closed, "Kapali: cikarma yoksa kapali kalir");

  z = T.zincirler([false, true, false], false);
  kontrol(z.length === 2 && z[0].idx.join(",") === "0" && z[1].idx.join(",") === "2",
    "Acik: ortadan cikinca 2 zincir");

  // --- 2x2 izgara: 4 kare, 4 ortak kenar ---
  CC.AYAR.tolerans = 0.05;
  CC.AYAR.yakinTolerans = 0.5;
  CC.AYAR.minOrtusme = 0.5;
  CC.AYAR.kapsam = "sayfa";
  CC.AYAR.otomatikSil = false;
  CC.AYAR.ortakKenarSil = true;

  var shapes = [
    dikdortgen(1, 0, 0, 10, 10, "L", 1),
    dikdortgen(2, 10, 0, 20, 10, "L", 2),
    dikdortgen(3, 0, 10, 10, 20, "L", 3),
    dikdortgen(4, 10, 10, 20, 20, "L", 4)
  ];

  var hata = kur(shapes);
  kontrol(hata === "", "Izgara tarandi (" + hata + ")");
  kontrol(T.g.tamSegmentler.length === 4,
    "4 ortak kenar cifti beklendi, gelen " + T.g.tamSegmentler.length);

  var s = T.sil();
  kontrol(s.ortakSegment === 4, "4 ortak kenar tek cizgiye indi, gelen " + s.ortakSegment);
  kontrol(s.ortakNesne === 3, "3 nesne yeniden kuruldu, gelen " + s.ortakNesne);

  // Kalan nesnelerde toplam segment: 4 + 3 + 3 + 2 = 12
  var canli = [];
  var i, j;
  for (i = 0; i < shapes.length; i++) if (!shapes[i]._deleted) canli.push(shapes[i]);
  for (i = 0; i < (shapes._yeni || []).length; i++) if (!shapes._yeni[i]._deleted) canli.push(shapes._yeni[i]);

  var toplam = 0, altYol = 0;
  for (i = 0; i < canli.length; i++) {
    var sps = canli[i].Curve.SubPaths;
    altYol += sps.Count;
    for (j = 1; j <= sps.Count; j++) toplam += sps.Item(j).Segments.Count;
  }
  kontrol(toplam === 12, "Toplam segment 12 beklendi, gelen " + toplam);
  kontrol(canli.length === 4, "4 nesne kaldi, gelen " + canli.length);
}

function testLogoKorunur() {
  CC.AYAR.tolerans = 0.05;
  CC.AYAR.yakinTolerans = 0.5;
  CC.AYAR.kapsam = "sayfa";
  CC.AYAR.otomatikSil = false;
  CC.AYAR.ortakKenarSil = true;

  // 1) Ayni nesnenin iki alt yolu ortak kenar paylasir -> DOKUNULMAZ
  var s1 = [seg(0, 0, 10, 0), seg(10, 0, 10, 10), seg(10, 10, 0, 10), seg(0, 10, 0, 0)];
  var s2 = [seg(10, 0, 20, 0), seg(20, 0, 20, 10), seg(20, 10, 10, 10), seg(10, 10, 10, 0)];
  var birlesik = shape(1, s1, true, "L", 1);
  birlesik.DisplayCurve = {
    SubPaths: liveColl([
      { Closed: true, Segments: coll(s1) },
      { Closed: true, Segments: coll(s2) }
    ])
  };
  birlesik.Curve = birlesik.DisplayCurve;

  kur([birlesik]);
  var a = T.sil();
  kontrol(a.ortakNesne === 0 && a.ortakSegment === 0,
    "Ayni nesnenin alt yollari: dokunulmadi (nesne=" + a.ortakNesne + ")");

  // 2) 8 kenarli harf + kare ortak kenar -> harf DOKUNULMAZ
  var harfSegs = [
    seg(10, 0, 20, 0), seg(20, 0, 20, 10), seg(20, 10, 15, 10),
    seg(15, 10, 15, 5), seg(15, 5, 10, 5), seg(10, 5, 10, 10),
    seg(10, 10, 5, 10), seg(5, 10, 10, 0)
  ];
  var harf = shape(2, harfSegs, true, "L", 2);
  var kare = dikdortgen(1, 0, 0, 10, 10, "L", 1);

  kur([kare, harf]);
  var b = T.sil();
  kontrol(b.ortakNesne === 0 && b.ortakSegment === 0,
    "8 kenarli harf korundu (nesne=" + b.ortakNesne + ", segment=" + b.ortakSegment + ")");
}

function testUstUsteKare() {
  CC.AYAR.tolerans = 0.05;
  CC.AYAR.yakinTolerans = 0.5;
  CC.AYAR.kapsam = "sayfa";
  CC.AYAR.ortakKenarSil = true;

  // A) TAM silme KAPALI: iki birebir kare -> sadece rapor, dokunma yok
  CC.AYAR.otomatikSil = false;
  var a = [dikdortgen(1, 0, 0, 10, 10, "L", 2), dikdortgen(2, 0, 0, 10, 10, "L", 1)];
  kur(a);
  var tam = T.g.bulgular.filter(function (f) { return f.tur === "TAM"; });
  kontrol(tam.length === 1, "Ust uste 2 kare: 1 TAM bulundu (" + tam.length + ")");

  var s1 = T.sil();
  kontrol(s1.ortakNesne === 0 && s1.ortakSegment === 0,
    "Silme kapaliyken ortak kenar mudahalesi yok");
  kontrol(a[0]._deleted === false && a[1]._deleted === false,
    "Silme kapaliyken iki kare de yerinde");

  // B) TAM silme ACIK: biri kalir, digeri silinir
  CC.AYAR.otomatikSil = true;
  var b = [dikdortgen(1, 0, 0, 10, 10, "L", 2), dikdortgen(2, 0, 0, 10, 10, "L", 1)];
  kur(b);
  var s2 = T.sil();
  var kalan = (b[0]._deleted ? 0 : 1) + (b[1]._deleted ? 0 : 1);
  kontrol(kalan === 1, "Ust uste 2 kare: tam 1 kare kaldi (" + kalan + ")");
  kontrol(s2.nesne === 1, "Tam 1 nesne silindi (" + s2.nesne + ")");
  kontrol(s2.ortakNesne === 0, "Silinen kopyaya ortak kenar mudahalesi yok");
}

// ---------------------------------------------------------------------

testFixture();
testImza();
testBirim();
testKatmanFarki();
testIzgara();
testOrtakKenar();
testLogoKorunur();
testUstUsteKare();

console.log(rapor.join("\n"));
console.log("\nGecen: " + gecen + "   Kalan: " + kalan);
process.exit(kalan === 0 ? 0 : 1);
