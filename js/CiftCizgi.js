/*
  =====================================================================
   cdrcift - Cift Cizgi Denetleyicisi  (CorelDRAW JavaScript surumu)
   Surum 1.0.0        CorelDRAW 2024 / 2025 / 2026
  =====================================================================

   NE YAPAR
   Plotter kesim dosyalarinda ust uste binen hatlari bulur, isaretler ve
   yalnizca kanitlanmis birebir kopyalari siler.

     TAM    : birebir kopya (kopyala-yapistir, ters yon, farkli baslangic)
     KISMI  : iki hattin bir bolumu ortak (bitisik parcalarin ortak kenari)
     YAKIN  : tolerans kadar kaymis kopya (cift kesim riski)
     NOKTAS : sifir uzunluklu segment / ust uste dugum

   KULLANIM
   1. Dosyayi acin.
   2. Bu scripti calistirin.
   3. Rapor: %TEMP%\cdrcift\rapor-*.txt (+ .csv, dosya yazma mumkunse)
   4. Bulgular DUP_CHECK katmaninda numarali isaretlenir (yazdirilamaz).
   5. Onay verirseniz TAM kopyalar tek geri-alma adiminda silinir.

   AYARLAR: asagidaki AYAR nesnesi. Scripti duzenleyip kaydedin.
   GUVENLIK: KISMI / YAKIN / NOKTAS asla otomatik silinmez.
             TAM silme yalniz ayni katman + ayni renk + ayni kalinlikta.
  =====================================================================
*/

var AYAR = {
  // --- toleranslar (mm) ---
  tolerans: 0.05,        // birebir sayilma esigi
  yakinTolerans: 0.5,    // "kaymis kopya" esigi
  minOrtusme: 0.5,       // raporlanacak en kisa ortusme (mm)

  // --- kapsam ---
  kapsam: "sayfa",       // "sayfa" | "secim"

  // --- temizlik ---
  otomatikSil: true,     // true: TAM kopyalari siler (tek Ctrl+Z ile geri alinir)
  silSonrasiTemizle: true, // silme basariliysa isaret katmanini da kaldir (ayni Ctrl+Z)
  boya: false,           // true: silinecek nesneleri kirmiziya boya

  // --- ortak kenar (izgara/karosel tasarimlari) ---
  ortakKenarSil: true,   // true: ortak kenarlari tek cizgiye indir
  ortakKenarMax: 200,    // tek calistirmada en fazla kac nesne yeniden kurulur
  maxOrtakCift: 5000,    // toplanacak en fazla ortak segment cifti
  ortakKenarAzamiKenar: 4, // yalniz bu kadar kenarli kapali hucreler (kare=4)
  ortakKenarTekAltYol: true, // yalniz tek alt yollu nesneler (harfler korunur)

  // --- isaretleme ---
  isaretKatmani: "DUP_CHECK",
  isaretSiniri: 150,     // en fazla kac bulgu isaretlenir

  // --- asama (hata ayiklama) ---
  // "tara"   : yalniz tara + ozet (isaret/rapor yok)
  // "isaret" : tara + isaretle
  // "rapor"  : tara + isaretle + rapor metni
  // "hepsi"  : normal calisma (varsayilan)
  asama: "hepsi",

  // --- tarama ---
  yoksayKatmanlar: [],   // ornek: ["Regmark", "Notlar"]
  hucreMm: 2.0,          // uzamsal izgara hucre boyu
  maxBulgu: 5000,
  powerClipTara: false,  // PowerClip icerigi taransin mi (bazi surumlerde cokmeye yol acabilir)
  maxOrnek: 4000,        // segment basina en fazla ornek (cok uzun cizgilerde bellek korumasi)
  maxSegment: 300000,    // toplam segment siniri

  // --- rapor ---
  raporKlasoru: "cdrcift", // %TEMP% altinda
  ondalik: ",",            // CSV/TXT ondalik ayirici (Turkce Excel: ",")
  raporMetni: false,       // true: raporu belgeye metin olarak da yaz (parcali)
  raporParca: 3500,        // metin parcasi basina karakter (32k siniri icin)
  raporParcaSayisi: 6      // en fazla kac parca yazilir
};

// Node testlerinden erisim icin (CorelDRAW icinde zararsiz)
var CDRCIFT = { AYAR: AYAR };

(function () {
  "use strict";

  var SURUM = "1.0.0";
  var MM = 0.005;                 // imza izgarasi
  var PARALEL_SIN = Math.sin(Math.PI / 180);

  // ------------------------------------------------------------------
  // Kucuk yardimcilar
  // ------------------------------------------------------------------

  function yuvarla(v) { return Math.floor(v / MM + 0.5); }
  function kucuk(a, b) { return a < b ? a : b; }
  function buyuk(a, b) { return a > b ? a : b; }

  function sayi(v, basamak) {
    var n = basamak === undefined ? 2 : basamak;
    var s = Number(v).toFixed(n);
    if (AYAR.ondalik === ",") s = s.replace(".", ",");
    return s;
  }

  function dene(fn, varsayilan) {
    try { return fn(); } catch (e) { return varsayilan; }
  }

  // COM proxy: bazi uyeler property, bazi parametreli uyeler fonksiyon olabilir.
  // DIKKAT: Bu host'ta metodlar icin `typeof obj.Metod` "undefined" donebiliyor
  // (CorelDRAW 2026'da olculdu) ama cagri calisiyor. Bu yuzden typeof'a guvenmeyiz.
  function P(obj, ad) {
    if (!obj) return null;
    var v = dene(function () { return obj[ad]; }, null);
    if (typeof v === "function") {
      var c = dene(function () { return v.call(obj); }, null);
      if (c !== null && c !== undefined) return c;
    }
    return v;
  }

  function M(obj, ad, a1, a2, a3, a4, a5, a6) {
    // Bu host'ta obj.Metod okumasi undefined donebiliyor; yalnizca
    // dogrudan cagri konumu calisiyor. Bu yuzden arity'ye gore cagiririz.
    if (!obj) return null;
    var self = obj;
    switch (arguments.length - 2) {
      case 0: return dene(function () { return self[ad](); }, null);
      case 1: return dene(function () { return self[ad](a1); }, null);
      case 2: return dene(function () { return self[ad](a1, a2); }, null);
      case 3: return dene(function () { return self[ad](a1, a2, a3); }, null);
      case 4: return dene(function () { return self[ad](a1, a2, a3, a4); }, null);
      case 5: return dene(function () { return self[ad](a1, a2, a3, a4, a5); }, null);
      case 6: return dene(function () { return self[ad](a1, a2, a3, a4, a5, a6); }, null);
      default: return null;
    }
  }

  // Koleksiyon elemani: .Item(i), koleksiyon(i), [i-1]
  function it(coll, i) {
    if (!coll) return null;
    var v = dene(function () { return coll.Item(i); }, null);
    if (v) return v;
    v = dene(function () { return coll(i); }, null);
    if (v) return v;
    v = dene(function () { return coll[i - 1]; }, null);
    return v || null;
  }

  function adet(coll) {
    var n = dene(function () { return coll.Count; }, null);
    if (typeof n === "number") return n;
    n = dene(function () { return coll.length; }, 0);
    return n || 0;
  }

  // ------------------------------------------------------------------
  // Birim
  // ------------------------------------------------------------------

  function birimKatsayi(kod) {
    // cdrUnit: 0 TenthMicron, 1 Inch, 2 Foot, 3 Millimeter, 4 Centimeter,
    //          5 Pixel, 6 Mile, 7 Meter, 8 Kilometer, 9 Didots, 10 Agate,
    //          11 Yard, 12 Pica, 13 Cicero, 14 Point, 15 Q, 16 H
    var tablo = {
      0: [0.0001, "0.1 um"], 1: [25.4, "inc"], 2: [304.8, "foot"],
      3: [1, "mm"], 4: [10, "cm"], 5: [1, "piksel"], 6: [1609344, "mil"],
      7: [1000, "m"], 8: [1000000, "km"], 9: [0.375972, "didot"],
      10: [4.233333, "agate"], 11: [914.4, "yarda"], 12: [4.233333, "pika"],
      13: [4.511278, "cicero"], 14: [0.352778, "pt"], 15: [0.25, "Q"],
      16: [0.1, "H"]
    };
    var t = tablo[kod];
    return t ? t : [1, "bilinmiyor"];
  }

  // ------------------------------------------------------------------
  // Durum
  // ------------------------------------------------------------------

  var g = {
    katsayi: 1,
    birimAdi: "mm",
    nesneler: [],        // {ad, katman, renk, kalinlik, z, altYolSayisi, uzunluk, ref}
    altYollar: [],       // {nesne, no, imza, uzunluk, kapali, x, y, segs}
    segmentler: [],      // {x1,y1,x2,y2,slot,si,tip,ref}
    tamSegmentler: [],   // ortak kenar ciftleri {a,b}
    bulgular: [],
    atlanan: 0,
    efekt: 0,
    powerclip: 0,
    uyarilar: [],
    kesildi: false,
    cokBuyuk: false
  };

  function uyari(metin) {
    if (g.uyarilar.length < 20) g.uyarilar.push(metin);
  }

  // ------------------------------------------------------------------
  // Yapisal imza (dugum duzeyinde birebir kopya)
  // ------------------------------------------------------------------

  function segToken(seg, ters) {
    var t1 = P(seg, "Type");
    var sn = P(seg, "StartNode");
    var en = P(seg, "EndNode");
    var x1 = P(sn, "PositionX"), y1 = P(sn, "PositionY");
    var x2 = P(en, "PositionX"), y2 = P(en, "PositionY");
    var k = g.katsayi;
    x1 *= k; y1 *= k; x2 *= k; y2 *= k;

    var s;
    if (!ters) {
      s = t1 + "|" + yuvarla(x1) + "|" + yuvarla(y1) + "|" + yuvarla(x2) + "|" + yuvarla(y2);
    } else {
      s = t1 + "|" + yuvarla(x2) + "|" + yuvarla(y2) + "|" + yuvarla(x1) + "|" + yuvarla(y1);
    }

    if (t1 === 1) {
      var c1x = P(seg, "StartingControlPointX") * k, c1y = P(seg, "StartingControlPointY") * k;
      var c2x = P(seg, "EndingControlPointX") * k, c2y = P(seg, "EndingControlPointY") * k;
      if (!ters) {
        s += "|" + yuvarla(c1x) + "|" + yuvarla(c1y) + "|" + yuvarla(c2x) + "|" + yuvarla(c2y);
      } else {
        s += "|" + yuvarla(c2x) + "|" + yuvarla(c2y) + "|" + yuvarla(c1x) + "|" + yuvarla(c1y);
      }
    }
    return s;
  }

  function kapaliKanonik(tok) {
    var n = tok.length, en = 0, i;
    for (i = 1; i < n; i++) { if (tok[i] < tok[en]) en = i; }
    var dizi = [];
    for (i = 0; i < n; i++) dizi.push(tok[(en + i) % n]);
    return dizi.join(";");
  }

  function imza(sp) {
    var segColl = P(sp, "Segments");
    var n = adet(segColl);
    if (!n) return "";
    var f = [], r = [], i;
    for (i = 1; i <= n; i++) {
      var sg = it(segColl, i);
      if (!sg) return "";
      f.push(segToken(sg, false));
      r.push(segToken(sg, true));
    }
    var kapali = !!P(sp, "Closed");
    var sf, sr;
    if (kapali) {
      sf = kapaliKanonik(f);
      sr = kapaliKanonik(r);
    } else {
      sf = f.join(";");
      sr = r.join(";");
    }
    return (kapali ? "C" : "O") + n + "|" + (sr < sf ? sr : sf);
  }

  // ------------------------------------------------------------------
  // Siniflandirma (VBA Karsilastir ile ayni)
  // ------------------------------------------------------------------

  var sonOrtusme = 0, sonOran = 0, sonX = 0, sonY = 0;

  function karsilastir(i, j) {
    var A = g.segmentler[i], B = g.segmentler[j];
    var l1 = Math.sqrt((A.x2 - A.x1) * (A.x2 - A.x1) + (A.y2 - A.y1) * (A.y2 - A.y1));
    var l2 = Math.sqrt((B.x2 - B.x1) * (B.x2 - B.x1) + (B.y2 - B.y1) * (B.y2 - B.y1));
    if (l1 < 0.001 || l2 < 0.001) return "";

    var ux = (A.x2 - A.x1) / l1, uy = (A.y2 - A.y1) / l1;
    var vx = (B.x2 - B.x1) / l2, vy = (B.y2 - B.y1) / l2;

    if (Math.abs(ux * vy - uy * vx) > PARALEL_SIN) return "";

    var d1 = Math.abs((B.x1 - A.x1) * (-uy) + (B.y1 - A.y1) * ux);
    var d2 = Math.abs((B.x2 - A.x1) * (-uy) + (B.y2 - A.y1) * ux);
    var off = buyuk(d1, d2);

    var t1 = (B.x1 - A.x1) * ux + (B.y1 - A.y1) * uy;
    var t2 = (B.x2 - A.x1) * ux + (B.y2 - A.y1) * uy;
    var lo = buyuk(0, kucuk(t1, t2));
    var hi = kucuk(l1, buyuk(t1, t2));
    var ov = hi - lo;
    if (ov < 0.001) return "";

    var mn = kucuk(l1, l2);
    var tur = "";

    if (off <= AYAR.tolerans) {
      if (ov >= mn - 0.001) tur = "TAM";
      else if (ov >= AYAR.minOrtusme) tur = "KISMI";
      else return "";
    } else if (off <= AYAR.yakinTolerans && ov >= 0.8 * mn) {
      tur = "YAKIN";
    } else {
      return "";
    }

    sonOrtusme = ov;
    sonOran = ov / mn;
    sonX = A.x1 + ux * (lo + ov / 2);
    sonY = A.y1 + uy * (lo + ov / 2);
    return tur;
  }

  // ------------------------------------------------------------------
  // Tarama
  // ------------------------------------------------------------------

  function katmanYoksay(ad) {
    var i;
    for (i = 0; i < AYAR.yoksayKatmanlar.length; i++) {
      if (String(AYAR.yoksayKatmanlar[i]).toLowerCase() === String(ad).toLowerCase()) return true;
    }
    return false;
  }

  function geometri(shape) {
    // Parametrik sekillerde Shape.Curve bos; DisplayCurve birincil.
    var crv = P(shape, "DisplayCurve");
    if (!crv) crv = P(shape, "Curve");
    if (crv && adet(P(crv, "SubPaths")) === 0) {
      var alt = P(shape, "Curve");
      if (alt && adet(P(alt, "SubPaths")) > 0) crv = alt;
    }
    return crv;
  }

  function nesneTara(shape, derinlik) {
    if (!shape || derinlik > 12) return;

    var tur = P(shape, "Type");

    if (tur === 7) {                       // grup
      var cocuklar = P(shape, "Shapes");
      var n = adet(cocuklar), i;
      for (i = 1; i <= n; i++) {
        try { nesneTara(it(cocuklar, i), derinlik + 1); } catch (e) { g.atlanan++; }
      }
      return;
    }

    // blend / extrude / contour / bevel / drop shadow / artistic media / custom effect
    if (tur === 10 || tur === 11 || tur === 13 || tur === 15 || tur === 16 || tur === 18 || tur === 22) {
      g.efekt++;
      return;
    }

    // Yalnizca kesimle ilgili vektor tipleri: dikdortgen, elips, egri, poligon, custom
    if (!(tur === 1 || tur === 2 || tur === 3 || tur === 4 || tur === 21)) { g.atlanan++; return; }

    var katman = P(P(shape, "Layer"), "Name") || "";
    if (katman.toUpperCase() === AYAR.isaretKatmani.toUpperCase()) return;
    if (katman.toUpperCase() === (AYAR.isaretKatmani + "_RAPOR").toUpperCase()) return;
    if (katmanYoksay(katman)) return;

    // PowerClip icerigi (varsayilan KAPALI: bazi surumlerde cokme riski)
    if (AYAR.powerClipTara) {
      var pc = P(shape, "PowerClip");
      if (pc) {
        var pcs = P(pc, "Shapes");
        var pn = adet(pcs), pi;
        for (pi = 1; pi <= pn; pi++) {
          try { nesneTara(it(pcs, pi), derinlik + 1); } catch (e) { g.atlanan++; }
        }
      }
    }

    curveTara(shape, katman);
  }

  function curveTara(shape, katman) {
    var crv = geometri(shape);
    if (!crv) { g.atlanan++; return; }

    var sps = P(crv, "SubPaths");
    if (!sps || adet(sps) === 0) { g.atlanan++; return; }

    var outline = P(shape, "Outline");
    var renk = dene(function () { return P(P(outline, "Color"), "HexValue"); }, "?");

    var nesne = {
      ad: dene(function () { return P(shape, "Name"); }, "") || ("Nesne " + P(shape, "StaticID")),
      katman: katman,
      renk: renk,
      kalinlik: (P(outline, "Width") || 0) * g.katsayi,
      z: dene(function () { return P(shape, "ZOrder"); }, 0),
      altYolSayisi: adet(sps),
      uzunluk: 0,
      ref: shape
    };
    var nesneNo = g.nesneler.length;
    g.nesneler.push(nesne);

    var n = adet(sps), i;
    for (i = 1; i <= n; i++) {
      altYolTara(nesneNo, i, it(sps, i));
    }
  }

  function altYolTara(nesneNo, spNo, sp) {
    if (!sp) return;

    var segs = [], noktalar = [];

    function ekle(x1, y1, x2, y2, si, tip, ref) {
      if (g.segmentler.length >= AYAR.maxSegment) { g.cokBuyuk = true; return false; }
      var boy = Math.sqrt((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1));
      if (boy < 0.005) return false;
      segs.push({ x1: x1, y1: y1, x2: x2, y2: y2, si: si, tip: tip, ref: ref });
      noktalar.push(x1, y1, x2, y2);
      return true;
    }

    var segColl = P(sp, "Segments");
    var sn = adet(segColl);
    if (!sn) return;

    var k = g.katsayi;
    var sifir = 0;

    var ilk = it(segColl, 1);
    var px = P(P(ilk, "StartNode"), "PositionX") * k;
    var py = P(P(ilk, "StartNode"), "PositionY") * k;

    var i;
    for (i = 1; i <= sn; i++) {
      var seg = it(segColl, i);
      if (!seg) continue;
      var tip = P(seg, "Type");
      var x2 = P(P(seg, "EndNode"), "PositionX") * k;
      var y2 = P(P(seg, "EndNode"), "PositionY") * k;

      if (tip === 0) {
        if (!ekle(px, py, x2, y2, i - 1, 0, seg)) sifir++;
        px = x2; py = y2;
      } else {
        var c1x = P(seg, "StartingControlPointX") * k;
        var c1y = P(seg, "StartingControlPointY") * k;
        var c2x = P(seg, "EndingControlPointX") * k;
        var c2y = P(seg, "EndingControlPointY") * k;
        var boy = Math.sqrt((x2 - px) * (x2 - px) + (y2 - py) * (y2 - py));
        var adim = Math.min(48, Math.max(6, Math.round(boy / 0.7) + 4));
        var j;
        for (j = 1; j <= adim; j++) {
          var t = j / adim, u = 1 - t;
          var a = u * u * u, b = 3 * u * u * t, c = 3 * u * t * t, d = t * t * t;
          var qx = a * px + b * c1x + c * c2x + d * x2;
          var qy = a * py + b * c1y + c * c2y + d * y2;
          if (!ekle(px, py, qx, qy, i - 1, 1, seg)) sifir++;
          px = qx; py = qy;
        }
      }
    }

    // agirlik merkezi (isaret konumu)
    var sx = 0, sy = 0, sc = 0, q;
    for (q = 0; q < noktalar.length; q += 2) { sx += noktalar[q]; sy += noktalar[q + 1]; sc++; }
    var cx = sc ? sx / sc : 0;
    var cy = sc ? sy / sc : 0;

    var uzunluk = 0, m;
    for (m = 0; m < segs.length; m++) {
      uzunluk += Math.sqrt(
        (segs[m].x2 - segs[m].x1) * (segs[m].x2 - segs[m].x1) +
        (segs[m].y2 - segs[m].y1) * (segs[m].y2 - segs[m].y1));
    }

    var slot = g.altYollar.length;
    g.altYollar.push({
      nesne: nesneNo, no: spNo, uzunluk: uzunluk,
      kapali: !!P(sp, "Closed"), x: cx, y: cy, imza: imza(sp),
      segSayisi: sn
    });
    g.nesneler[nesneNo].uzunluk += uzunluk;

    var p;
    for (p = 0; p < segs.length; p++) {
      g.segmentler.push({
        x1: segs[p].x1, y1: segs[p].y1, x2: segs[p].x2, y2: segs[p].y2,
        slot: slot, si: segs[p].si, tip: segs[p].tip, ref: segs[p].ref
      });
    }

    if (sifir > 0) {
      bulguEkle("NOKTAS", nesneNo, spNo, -1, 0, 0, 0, cx, cy, "elle", sifir + " adet sifir uzunluklu segment");
    }
  }

  // ------------------------------------------------------------------
  // Bulgular
  // ------------------------------------------------------------------

  function bulguEkle(tur, n1, a1, n2, a2, uzunluk, oran, x, y, eylem, not) {
    if (g.bulgular.length >= AYAR.maxBulgu) { g.kesildi = true; return -1; }
    g.bulgular.push({
      tur: tur, n1: n1, a1: a1, n2: n2, a2: a2,
      uzunluk: uzunluk, oran: oran, x: x, y: y, eylem: eylem, not: not || ""
    });
    return g.bulgular.length - 1;
  }

  function ayniOzellik(n1, n2) {
    var a = g.nesneler[n1], b = g.nesneler[n2];
    if (!a || !b) return "nesne yok";
    if (a.katman !== b.katman) return "farkli katman";
    if (Math.abs(a.kalinlik - b.kalinlik) > 0.001) return "farkli kalinlik";
    if (String(a.renk).toUpperCase() !== String(b.renk).toUpperCase()) return "farkli renk";
    return "";
  }

  function yapisalTam() {
    var gruplar = {}, si;
    for (si = 0; si < g.altYollar.length; si++) {
      var im = g.altYollar[si].imza;
      if (!im) continue;
      if (!gruplar[im]) gruplar[im] = [];
      gruplar[im].push(si);
    }

    var tamCift = {};   // "n|a|n|a" -> 1

    for (var anahtar in gruplar) {
      if (!Object.prototype.hasOwnProperty.call(gruplar, anahtar)) continue;
      var uyeler = gruplar[anahtar];
      if (uyeler.length < 2) continue;

      // korunan: en alttaki (buyuk ZOrder)
      var tutucu = uyeler[0];
      var j;
      for (j = 1; j < uyeler.length; j++) {
        var zj = g.nesneler[g.altYollar[uyeler[j]].nesne].z;
        var zt = g.nesneler[g.altYollar[tutucu].nesne].z;
        if (zj > zt) tutucu = uyeler[j];
      }

      var t = g.altYollar[tutucu];
      for (j = 0; j < uyeler.length; j++) {
        if (uyeler[j] === tutucu) continue;
        var d = g.altYollar[uyeler[j]];
        var eylem, neden = "";

        if (d.nesne === t.nesne) {
          eylem = "sil";
          neden = "ayni nesne icinde tekrar";
        } else {
          neden = ayniOzellik(d.nesne, t.nesne);
          eylem = (neden === "" && AYAR.otomatikSil) ? "sil" : "elle";
        }

        bulguEkle("TAM", d.nesne, d.no, t.nesne, t.no, d.uzunluk, 1,
          d.x, d.y, eylem, "birebir kopya" + (neden ? " (" + neden + ")" : ""));

        tamCift[d.nesne + "|" + d.no + "|" + t.nesne + "|" + t.no] = 1;
        tamCift[t.nesne + "|" + t.no + "|" + d.nesne + "|" + d.no] = 1;
      }
    }
    return tamCift;
  }

  function ciftAnahtar(n1, a1, n2, a2) {
    if (n1 < n2 || (n1 === n2 && a1 <= a2)) return n1 + "|" + a1 + "|" + n2 + "|" + a2;
    return n2 + "|" + a2 + "|" + n1 + "|" + a1;
  }

  function hucre(x, y) {
    return Math.floor(x / AYAR.hucreMm) + "|" + Math.floor(y / AYAR.hucreMm);
  }

  function ornekle(s, adim, fn) {
    var boy = Math.sqrt((s.x2 - s.x1) * (s.x2 - s.x1) + (s.y2 - s.y1) * (s.y2 - s.y1));
    var n = Math.round(boy / adim);
    if (!isFinite(n) || n < 1) n = 1;
    if (n > AYAR.maxOrnek) n = AYAR.maxOrnek;
    var i;
    for (i = 0; i <= n; i++) {
      var t = i / n;
      fn(s.x1 + (s.x2 - s.x1) * t, s.y1 + (s.y2 - s.y1) * t);
    }
  }

  function segmentTarama(tamCift) {
    var izgara = {}, i;
    var adimEkle = Math.max(AYAR.hucreMm / 2, 0.25);

    for (i = 0; i < g.segmentler.length; i++) {
      (function (idx) {
        var son = null;
        ornekle(g.segmentler[idx], adimEkle, function (x, y) {
          var k = hucre(x, y);
          if (k === son) return;
          son = k;
          if (!izgara[k]) izgara[k] = [];
          izgara[k].push(idx);
        });
      })(i);
    }

    var gorulen = {}, birlesik = {};

    for (i = 0; i < g.segmentler.length; i++) {
      (function (idx) {
        var aday = {}, hucreler = {}, son = null;

        ornekle(g.segmentler[idx], AYAR.hucreMm, function (x, y) {
          var cx = Math.floor(x / AYAR.hucreMm), cy = Math.floor(y / AYAR.hucreMm);
          var dx, dy;
          for (dx = -1; dx <= 1; dx++) {
            for (dy = -1; dy <= 1; dy++) {
              var k = (cx + dx) + "|" + (cy + dy);
              if (hucreler[k]) continue;
              hucreler[k] = 1;
              var liste = izgara[k];
              if (!liste) continue;
              var q;
              for (q = 0; q < liste.length; q++) aday[liste[q]] = 1;
            }
          }
        });

        for (var kk in aday) {
          if (!Object.prototype.hasOwnProperty.call(aday, kk)) continue;
          var j = parseInt(kk, 10);
          if (j <= idx) continue;
          var sa = g.segmentler[idx].slot, sb = g.segmentler[j].slot;
          if (sa === sb && j - idx === 1) continue;

          var tur = karsilastir(idx, j);
          if (!tur) continue;

          // Tam cakisan segmentler = ortak kenar adayi (yalniz duz cizgi)
          if (g.segmentler[idx].tip === 0 && g.segmentler[j].tip === 0) {
            var lA = Math.sqrt(Math.pow(g.segmentler[idx].x2 - g.segmentler[idx].x1, 2) +
                                Math.pow(g.segmentler[idx].y2 - g.segmentler[idx].y1, 2));
            var lB = Math.sqrt(Math.pow(g.segmentler[j].x2 - g.segmentler[j].x1, 2) +
                                Math.pow(g.segmentler[j].y2 - g.segmentler[j].y1, 2));
            if (sonOrtusme >= 0.98 * Math.max(lA, lB) && g.tamSegmentler.length < AYAR.maxOrtakCift) {
              g.tamSegmentler.push({ a: idx, b: j });
            }
          }

          var A = g.altYollar[sa], B = g.altYollar[sb];
          if (tamCift[A.nesne + "|" + A.no + "|" + B.nesne + "|" + B.no]) continue;

          var an = ciftAnahtar(A.nesne, A.no, B.nesne, B.no) + "|" + tur;
          if (birlesik[an] !== undefined) {
            g.bulgular[birlesik[an]].uzunluk += sonOrtusme;
          } else {
            var yeni = bulguEkle(tur, A.nesne, A.no, B.nesne, B.no,
              sonOrtusme, sonOran, sonX, sonY, "elle", "");
            if (yeni >= 0) birlesik[an] = yeni;
          }
        }
      })(i);
    }
  }

  function bulguTamamla() {
    var uzunluk = {}, i;
    for (i = 0; i < g.altYollar.length; i++) {
      var a = g.altYollar[i];
      uzunluk[a.nesne + "|" + a.no] = a.uzunluk;
    }

    for (i = 0; i < g.bulgular.length; i++) {
      var f = g.bulgular[i];

      if (f.tur !== "NOKTAS") {
        var b1 = uzunluk[f.n1 + "|" + f.a1] || 0;
        var b2 = f.n2 >= 0 ? (uzunluk[f.n2 + "|" + f.a2] || b1) : b1;
        var mn = kucuk(b1, b2);
        if (mn > 0) f.oran = kucuk(1, f.uzunluk / mn);
      }

      if (f.tur === "TAM" && f.oran < 0.98) f.tur = "KISMI";

      if (f.tur === "TAM") {
        if (f.eylem !== "sil" && AYAR.otomatikSil && f.n1 !== f.n2) {
          if (ayniOzellik(f.n1, f.n2) === "") f.eylem = "sil";
        }
        if (f.eylem === "sil" && !AYAR.otomatikSil) f.eylem = "elle";
      } else {
        f.eylem = "elle";
      }
    }
  }

  // ------------------------------------------------------------------
  // Isaretleme
  // ------------------------------------------------------------------

  function katmanBulAd(page, ad) {
    var layers = P(page, "Layers");
    var n = adet(layers), i;
    for (i = 1; i <= n; i++) {
      var l = it(layers, i);
      if (l && String(P(l, "Name")).toUpperCase() === String(ad).toUpperCase()) return l;
    }
    return null;
  }

  function katmanBul(page) {
    return katmanBulAd(page, AYAR.isaretKatmani);
  }

  function katmanAl(page) {
    var l = katmanBul(page);
    if (l) return l;
    l = M(page, "CreateLayer", AYAR.isaretKatmani);
    if (l) {
      dene(function () { l.Printable = false; });
      dene(function () { l.Editable = true; });
    }
    return l;
  }

  function isaretleriKaldir(page) {
    var l = katmanBul(page);
    var l2 = katmanBulAd(page, AYAR.isaretKatmani + "_RAPOR");
    if (!l && !l2) return false;

    var doc = host.ActiveDocument;
    dene(function () { doc.BeginCommandGroup("Isaretleri kaldir"); });
    if (l) dene(function () { M(l, "Delete"); });
    if (l2) dene(function () { M(l2, "Delete"); });
    dene(function () { doc.EndCommandGroup(); });
    return true;
  }

  function turRenk(tur, c) {
    var r = 0, gr = 0, b = 0;
    if (tur === "TAM") { r = 255; }
    else if (tur === "KISMI") { r = 255; gr = 128; }
    else if (tur === "YAKIN") { gr = 112; b = 255; }
    else { r = 128; b = 160; }
    dene(function () { M(c, "RGBAssign", r, gr, b); });
  }

  function isaretle(page) {
    isaretleriKaldir(page);          // onceki isaretler (tekrar calistirmalarda)
    var lyr = katmanAl(page);
    if (!lyr) { uyari("DUP_CHECK katmani olusturulamadi"); return 0; }

    // Yazdirilabilir kalirsa isaretler plotter'a gider: dogrula ve uyar
    dene(function () { lyr.Printable = false; });
    if (P(lyr, "Printable") === true) {
      uyari("DUP_CHECK katmani 'yazdirilamaz' yapilamadi. Cikti oncesi elle kapatın.");
    }

    var doc = host.ActiveDocument;
    dene(function () { doc.BeginCommandGroup("Cift cizgi isaretleri"); });

    var r = 1.5 / g.katsayi;
    var n = 0, i;

    for (i = 0; i < g.bulgular.length && n < AYAR.isaretSiniri; i++) {
      var f = g.bulgular[i];
      var x = f.x / g.katsayi, y = f.y / g.katsayi;

      var el = M(lyr, "CreateEllipse", x - r, y + r, x + r, y - r);
      if (el) {
        dene(function () { P(el, "Outline").Width = 0.5 / g.katsayi; });
        turRenk(f.tur, P(P(el, "Outline"), "Color"));
        dene(function () { M(P(el, "Fill"), "ApplyNoFill"); });
      }

      var tx = M(lyr, "CreateArtisticText", x + r * 1.6, y + r, String(i + 1));
      if (tx) turRenk(f.tur, P(P(tx, "Outline"), "Color"));
      n++;
    }

    if (AYAR.boya) {
      for (i = 0; i < g.bulgular.length; i++) {
        if (g.bulgular[i].tur === "TAM" && g.bulgular[i].eylem === "sil") {
          var nesne = g.nesneler[g.bulgular[i].n1];
          if (nesne && nesne.ref) {
            dene(function () { turRenk("TAM", P(P(nesne.ref, "Outline"), "Color")); });
          }
        }
      }
    }

    dene(function () { doc.EndCommandGroup(); });
    return n;
  }

  // --- Ortak kenar temizligi (izgara/karosel tasarimlari) -------------

  // Kaldirilan segmentlere gore kesintisiz zincirler uretir.
  // kaldir: [true/false,...]  kapali: subpath kapali mi
  // Donus: [{ idx: [segment indeksleri], closed: bool }, ...]
  function zincirler(kaldir, kapali) {
    var n = kaldir.length, i;
    var chain = [], chains = [];
    var start = 0;
    var herhangi = false;

    for (i = 0; i < n; i++) if (kaldir[i]) herhangi = true;
    if (!herhangi) {
      for (i = 0; i < n; i++) chain.push(i);
      return [{ idx: chain, closed: !!kapali }];
    }

    if (kapali) {
      for (i = 0; i < n; i++) {
        if (kaldir[i]) { start = (i + 1) % n; break; }
      }
    }

    for (i = 0; i < n; i++) {
      var idx = (start + i) % n;
      if (kaldir[idx]) {
        if (chain.length) { chains.push({ idx: chain, closed: false }); chain = []; }
      } else {
        chain.push(idx);
      }
    }
    if (chain.length) chains.push({ idx: chain, closed: false });
    return chains;
  }

  function segEkle(spY, ref) {
    if (!ref) return false;
    var tip = P(ref, "Type");
    var ex = P(P(ref, "EndNode"), "PositionX");
    var ey = P(P(ref, "EndNode"), "PositionY");
    if (ex === null || ey === null) return false;

    if (tip === 0) {
      M(spY, "AppendLineSegment", ex, ey);
      return true;
    }
    var c1x = P(ref, "StartingControlPointX"), c1y = P(ref, "StartingControlPointY");
    var c2x = P(ref, "EndingControlPointX"), c2y = P(ref, "EndingControlPointY");
    M(spY, "AppendCurveSegment", c1x, c1y, c2x, c2y, ex, ey);
    return true;
  }

  function yenidenKur(nesneNo, kaldirMap, slotSegs) {
    var nesne = g.nesneler[nesneNo];
    if (!nesne || !nesne.ref) return false;
    var shape = nesne.ref;
    var layer = P(shape, "Layer");
    var crvOld = geometri(shape);
    var doc = host.ActiveDocument;
    if (!layer || !crvOld) return false;

    var yeniCrv = M(doc, "CreateCurve");
    if (!yeniCrv) return false;

    var yazilan = 0, q;
    var beklenen = 0, uretilen = 0;
    for (q = 0; q < g.altYollar.length; q++) {
      var slot = g.altYollar[q];
      if (slot.nesne !== nesneNo) continue;

      var segList = slotSegs[q] || [];
      if (!segList.length) continue;

      var kaldir = [], t;
      for (t = 0; t < segList.length; t++) {
        var m = kaldirMap[slot.no];
        var git = !!(m && m[segList[t].si]);
        kaldir.push(git);
        if (!git) beklenen++;
      }

      var chains = zincirler(kaldir, slot.kapali);

      for (t = 0; t < chains.length; t++) {
        var c = chains[t];
        var ilk = segList[c.idx[0]].ref;
        if (!ilk) return false;
        var sx = P(P(ilk, "StartNode"), "PositionX");
        var sy = P(P(ilk, "StartNode"), "PositionY");
        if (sx === null || sy === null) return false;

        var spY = M(yeniCrv, "CreateSubPath", sx, sy);
        if (!spY) return false;

        var u;
        for (u = 0; u < c.idx.length; u++) {
          if (!segEkle(spY, segList[c.idx[u]].ref)) return false;
        }
        if (c.closed) dene(function () { spY.Closed = true; });
        uretilen += c.idx.length;
        yazilan++;
      }
    }

    if (!yazilan) return false;
    // Guvenlik: beklenen ile uretilen segment sayisi tutmuyorsa dokunma
    if (uretilen !== beklenen) return false;

    var yeni = M(layer, "CreateCurve", yeniCrv);
    if (!yeni) return false;

    // Cizgi ozelliklerini kopyala (renk = kalem secimi, kritik)
    dene(function () {
      var o = P(shape, "Outline"), n = P(yeni, "Outline");
      n.Width = P(o, "Width");
      M(P(n, "Color"), "CopyAssign", P(o, "Color"));
    });
    dene(function () { M(yeni, "OrderFrontOf", shape); });

    return dene(function () { M(shape, "Delete"); return true; }, false) === true;
  }

  function ortakKenarSil(page, silinen) {
    var sonuc = { segment: 0, nesne: 0, atlanan: 0 };
    if (!AYAR.ortakKenarSil || !g.tamSegmentler.length) return sonuc;
    silinen = silinen || {};

    // Slot -> segment listesi (si sirali)
    var slotSegs = {}, i;
    for (i = 0; i < g.segmentler.length; i++) {
      var sg = g.segmentler[i];
      if (!slotSegs[sg.slot]) slotSegs[sg.slot] = [];
      slotSegs[sg.slot].push(sg);
    }
    for (var sl in slotSegs) {
      if (Object.prototype.hasOwnProperty.call(slotSegs, sl)) {
        slotSegs[sl].sort(function (a, b) { return a.si - b.si; });
      }
    }

    function once(aSlot, aSi, bSlot, bSi) {
      var sa = g.altYollar[aSlot], sb = g.altYollar[bSlot];
      if (sa.nesne !== sb.nesne) return sa.nesne < sb.nesne;
      if (sa.no !== sb.no) return sa.no < sb.no;
      return aSi < bSi;
    }

    // Kim hangi segmenti kaybedecek
    var kaldir = {};
    for (i = 0; i < g.tamSegmentler.length; i++) {
      var A = g.segmentler[g.tamSegmentler[i].a];
      var B = g.segmentler[g.tamSegmentler[i].b];
      if (A.slot === B.slot) continue;

      var sa = g.altYollar[A.slot], sb = g.altYollar[B.slot];

      // --- GUVENLIK KURALLARI (harf/logo bozulmasini onler) ---
      if (silinen[sa.nesne] || silinen[sb.nesne]) continue; // silinmis nesneye dokunma
      if (sa.nesne === sb.nesne) continue;                 // ayni nesnenin alt yollari
      if (sa.imza && sa.imza === sb.imza) continue;        // birebir kopya -> TAM isi
      if (!sa.kapali || !sb.kapali) continue;              // ikisi de kapali hucre
      if (sa.segSayisi !== AYAR.ortakKenarAzamiKenar) continue;
      if (sb.segSayisi !== AYAR.ortakKenarAzamiKenar) continue;
      if (AYAR.ortakKenarTekAltYol) {
        if (g.nesneler[sa.nesne].altYolSayisi !== 1) continue;
        if (g.nesneler[sb.nesne].altYolSayisi !== 1) continue;
      }

      var kaybeden = once(A.slot, A.si, B.slot, B.si) ? B : A;
      var s = g.altYollar[kaybeden.slot];
      var anahtar = s.nesne + "|" + s.no;
      if (!kaldir[anahtar]) kaldir[anahtar] = {};
      kaldir[anahtar][kaybeden.si] = 1;
    }

    // Nesne bazinda grupla
    var nesneMap = {};
    for (var anahtar2 in kaldir) {
      if (!Object.prototype.hasOwnProperty.call(kaldir, anahtar2)) continue;
      var parca = anahtar2.split("|");
      var nn = parseInt(parca[0], 10), spno = parseInt(parca[1], 10);
      if (!nesneMap[nn]) nesneMap[nn] = {};
      nesneMap[nn][spno] = kaldir[anahtar2];
    }

    var islenen = 0;
    for (var ns in nesneMap) {
      if (!Object.prototype.hasOwnProperty.call(nesneMap, ns)) continue;
      if (islenen >= AYAR.ortakKenarMax) { sonuc.atlanan++; continue; }

      var nesneNo = parseInt(ns, 10);
      if (silinen[nesneNo]) continue;
      var nesne = g.nesneler[nesneNo];
      if (!nesne || !nesne.ref) continue;

      // Egri iceren nesneler elle birakilir
      var egriVar = false, q;
      for (q = 0; q < g.altYollar.length; q++) {
        if (g.altYollar[q].nesne !== nesneNo) continue;
        var lst = slotSegs[q] || [];
        var u;
        for (u = 0; u < lst.length; u++) if (lst[u].tip === 1) { egriVar = true; break; }
        if (egriVar) break;
      }
      if (egriVar) { sonuc.atlanan++; continue; }

      // Tum segmentler kalkiyor mu?
      var toplam = 0, kalkan = 0;
      for (q = 0; q < g.altYollar.length; q++) {
        if (g.altYollar[q].nesne !== nesneNo) continue;
        var l2 = slotSegs[q] || [];
        var v;
        for (v = 0; v < l2.length; v++) {
          toplam++;
          if (nesneMap[ns][g.altYollar[q].no] && nesneMap[ns][g.altYollar[q].no][l2[v].si]) kalkan++;
        }
      }

      if (toplam > 0 && kalkan >= toplam) {
        if (dene(function () { M(nesne.ref, "Delete"); return true; }, false)) {
          sonuc.nesne++; sonuc.segment += kalkan;
          silinen[nesneNo] = 1;
        }
        continue;
      }

      var ok = false;
      try { ok = yenidenKur(nesneNo, nesneMap[ns], slotSegs); } catch (e) { ok = false; }
      if (ok) { sonuc.nesne++; sonuc.segment += kalkan; silinen[nesneNo] = 1; }
      else sonuc.atlanan++;
      islenen++;
    }

    return sonuc;
  }

  // ------------------------------------------------------------------
  // Silme (TAM + ortak kenar)
  // ------------------------------------------------------------------

  function sil() {
    var doc = host.ActiveDocument;
    var silAltYol = {}, sayim = {}, kendKopya = {};
    var i;

    for (i = 0; i < g.bulgular.length; i++) {
      var f = g.bulgular[i];
      if (f.tur === "TAM" && f.eylem === "sil") {
        silAltYol[f.n1 + "|" + f.a1] = 1;
        if (f.n1 === f.n2) kendKopya[f.n1] = 1;
      }
    }

    var anahtarlar = Object.keys(silAltYol);
    if (!anahtarlar.length && !AYAR.ortakKenarSil) return { nesne: 0, altYol: 0, ortakSegment: 0, ortakNesne: 0, atlanan: 0 };

    for (i = 0; i < anahtarlar.length; i++) {
      var ns = anahtarlar[i].split("|")[0];
      sayim[ns] = (sayim[ns] || 0) + 1;
    }

    // tamamen silinecek nesneler
    var silNesne = {};
    for (i = 0; i < g.nesneler.length; i++) {
      if (sayim[i] && sayim[i] >= g.nesneler[i].altYolSayisi && !kendKopya[i]) silNesne[i] = 1;
    }

    var nesneSil = 0, altSil = 0;
    var silinenNesne = {};
    dene(function () { doc.BeginCommandGroup("Cift cizgi temizligi"); });

    // ondeki once (ZOrder 1 = en ustte)
    var liste = Object.keys(silNesne).map(Number);
    liste.sort(function (a, b) { return g.nesneler[a].z - g.nesneler[b].z; });

    for (i = 0; i < liste.length; i++) {
      var ref = g.nesneler[liste[i]].ref;
      if (!ref) continue;
      if (dene(function () { M(ref, "Delete"); return true; }, false)) {
        nesneSil++;
        silinenNesne[liste[i]] = 1;
      }
    }

    // kismi: yalniz ilgili alt yol (buyuk indeksten kucuge)
    for (i = 0; i < g.nesneler.length; i++) {
      if (silNesne[i]) continue;
      var noList = {};
      var j;
      for (j = 0; j < g.bulgular.length; j++) {
        var b = g.bulgular[j];
        if (b.tur === "TAM" && b.eylem === "sil" && b.n1 === i) noList[b.a1] = 1;
      }
      var numaralar = Object.keys(noList).map(Number).sort(function (a, b) { return b - a; });
      if (!numaralar.length) continue;

      var sr = g.nesneler[i].ref;
      for (j = 0; j < numaralar.length; j++) {
        (function (no) {
          var sp = dene(function () { return it(P(P(sr, "Curve"), "SubPaths"), no); }, null);
          if (!sp) sp = dene(function () { return it(P(P(sr, "DisplayCurve"), "SubPaths"), no); }, null);
          if (!sp) return;
          var toplam = dene(function () { return adet(P(P(sr, "Curve"), "SubPaths")); }, 0);
          if (toplam > 1) {
            if (dene(function () { M(sp, "Delete"); return true; }, false)) altSil++;
          }
        })(numaralar[j]);
      }
    }

    // Ortak kenarlari tek cizgiye indir
    var ortak = { segment: 0, nesne: 0, atlanan: 0 };
    if (AYAR.ortakKenarSil) {
      ortak = ortakKenarSil(P(doc, "ActivePage"), silinenNesne);
    }

    if (AYAR.silSonrasiTemizle && (nesneSil + altSil + ortak.nesne) > 0) {
      var pg = P(doc, "ActivePage");
      var ml = katmanBul(pg);
      if (ml) dene(function () { M(ml, "Delete"); });
      var rl = katmanBulAd(pg, AYAR.isaretKatmani + "_RAPOR");
      if (rl) dene(function () { M(rl, "Delete"); });
    }

    dene(function () { doc.EndCommandGroup(); });
    return {
      nesne: nesneSil, altYol: altSil,
      ortakSegment: ortak.segment, ortakNesne: ortak.nesne, atlanan: ortak.atlanan
    };
  }

  // ------------------------------------------------------------------
  // Rapor
  // ------------------------------------------------------------------

  var fso = null;

  function dosyaSistemi() {
    // Bu host'ta dosya erisimi genellikle yok (Teshis: FSO=yok).
    if (fso !== null) return fso || null;
    fso = dene(function () { return host.CreateObject("Scripting.FileSystemObject"); }, null);
    if (!fso) fso = dene(function () { return new ActiveXObject("Scripting.FileSystemObject"); }, null);
    return fso || null;
  }

  function tempKlasor() {
    var f = dosyaSistemi();
    if (f) {
      var t = dene(function () { return f.GetSpecialFolder(2).Path; }, null);
      if (t) return t;
    }
    var cst = dene(function () { return host.CorelScriptTools; }, null);
    if (cst) {
      var c = dene(function () { return M(cst, "GetTempFolder"); }, null);
      if (c) return c;
    }
    return null;
  }

  function raporMetni(baslikSonek) {
    var tam = 0, kismi = 0, yakin = 0, nokta = 0, silN = 0, i;
    for (i = 0; i < g.bulgular.length; i++) {
      var t = g.bulgular[i].tur;
      if (t === "TAM") tam++;
      else if (t === "KISMI") kismi++;
      else if (t === "YAKIN") yakin++;
      else nokta++;
      if (t === "TAM" && g.bulgular[i].eylem === "sil") silN++;
    }

    var L = [];
    L.push("CIFT CIZGI DENETLEYICISI - RAPOR");
    L.push("Surum      : " + SURUM);
    L.push("Tarih      : " + new Date().toLocaleString());
    L.push("Belge      : " + (dene(function () { return host.ActiveDocument.FileName; }, "") || "(kaydedilmemis)"));
    L.push("Birim      : " + g.birimAdi + " (1 birim = " + sayi(g.katsayi, 4) + " mm)");
    L.push("Kapsam     : " + AYAR.kapsam);
    L.push("Tolerans   : " + sayi(AYAR.tolerans, 3) + " mm");
    L.push("Yakin tol. : " + sayi(AYAR.yakinTolerans, 3) + " mm");
    L.push("Min ortusme: " + sayi(AYAR.minOrtusme, 2) + " mm");
    L.push("");
    L.push("OZET");
    L.push("  TAM (birebir)      : " + tam);
    L.push("  KISMI (ortak kenar): " + kismi);
    L.push("  YAKIN (kaymis)     : " + yakin);
    L.push("  NOKTAS             : " + nokta);
    L.push("  TOPLAM             : " + g.bulgular.length);
    L.push("  Silinebilir TAM    : " + silN);
    L.push("");
    L.push("TARANAN");
    L.push("  Nesne    : " + g.nesneler.length);
    L.push("  Alt yol  : " + g.altYollar.length);
    L.push("  Segment  : " + g.segmentler.length);
    L.push("  Atlanan nesne: " + g.atlanan + "   Efekt: " + g.efekt + "   PowerClip: " + g.powerclip);
    if (g.kesildi) L.push("  NOT: Bulgu listesi " + AYAR.maxBulgu + " kayitla sinirlandi.");
    if (g.uyarilar.length) {
      L.push("");
      L.push("UYARILAR");
      var u;
      for (u = 0; u < g.uyarilar.length; u++) L.push("  - " + g.uyarilar[u]);
    }
    L.push("");
    L.push("No | Tur | Katman | Nesne | AltYol | Uzunluk mm | Oran % | Konum X,Y | Eylem | Eslesen");
    L.push(new Array(120).join("-"));

    for (i = 0; i < g.bulgular.length; i++) {
      var f = g.bulgular[i];
      var nes = g.nesneler[f.n1];
      var es = f.n2 >= 0 ? g.nesneler[f.n2] : null;
      L.push(
        ("     " + (i + 1)).slice(-5) + " | " +
        (f.tur + "      ").slice(0, 6) + " | " +
        (nes.katman + "          ").slice(0, 10) + " | " +
        (nes.ad + "                    ").slice(0, 20) + " | " +
        ("   " + f.a1).slice(-3) + " | " +
        ("        " + sayi(f.uzunluk, 2)).slice(-8) + " | " +
        ("    " + sayi(f.oran * 100, 0)).slice(-4) + " | " +
        sayi(f.x, 2) + "," + sayi(f.y, 2) + " | " +
        (f.eylem + "    ").slice(0, 4) + " | " +
        (es ? es.ad : "-") +
        (f.not ? " | " + f.not : ""));
    }
    return L.join("\r\n");
  }

  function csvMetni() {
    var L = [];
    L.push("No;Tur;Katman;Nesne;AltYol;Uzunluk_mm;Oran_yuzde;X;Y;Eylem;Eslesen;Not");
    var i;
    for (i = 0; i < g.bulgular.length; i++) {
      var f = g.bulgular[i];
      var nes = g.nesneler[f.n1];
      var es = f.n2 >= 0 ? g.nesneler[f.n2] : null;
      L.push([
        (i + 1), f.tur, nes.katman, nes.ad, f.a1,
        sayi(f.uzunluk, 2), sayi(f.oran * 100, 0), sayi(f.x, 2), sayi(f.y, 2),
        f.eylem, (es ? es.ad : ""), (f.not || "").replace(/;/g, ",")
      ].join(";"));
    }
    return L.join("\r\n");
  }

  function raporYaz() {
    var f = dosyaSistemi();
    var t = tempKlasor();
    if (!f || !t) return null;
    var klasor = t + "\\" + AYAR.raporKlasoru;
    dene(function () { if (!f.FolderExists(klasor)) f.CreateFolder(klasor); });
    if (!dene(function () { return f.FolderExists(klasor); }, false)) return null;

    var d = new Date();
    function iki(v) { return (v < 10 ? "0" : "") + v; }
    var damga = "" + d.getFullYear() + iki(d.getMonth() + 1) + iki(d.getDate()) + "-" +
                iki(d.getHours()) + iki(d.getMinutes()) + iki(d.getSeconds());

    var txt = klasor + "\\rapor-" + damga + ".txt";
    var csv = klasor + "\\rapor-" + damga + ".csv";

    var ok1 = dene(function () {
      var w = f.CreateTextFile(txt, true);
      w.Write(raporMetni());
      w.Close();
      return true;
    }, false);

    var ok2 = dene(function () {
      var w = f.CreateTextFile(csv, true);
      w.Write(csvMetni());
      w.Close();
      return true;
    }, false);

    if (!ok1) return null;
    return { txt: txt, csv: ok2 ? csv : "" };
  }

  function raporNesnesi(page, metin) {
    // CorelDRAW sanat metni ~32.000 karakter siniri: parcALI yaz.
    var katAd = AYAR.isaretKatmani + "_RAPOR";

    var eski = katmanBulAd(page, katAd);
    if (eski) dene(function () { M(eski, "Delete"); });

    var lyr = M(page, "CreateLayer", katAd);
    if (!lyr) return null;

    dene(function () { lyr.Printable = false; });
    dene(function () { lyr.Editable = true; });

    var w = P(page, "SizeWidth") || 200;
    var h = P(page, "SizeHeight") || 200;

    // Satir bazinda parcala
    var satirlar = metin.split(/\r?\n/);
    var parcalar = [], tampon = "";
    var i;
    for (i = 0; i < satirlar.length; i++) {
      var s = satirlar[i];
      if (tampon.length + s.length + 2 > AYAR.raporParca) {
        if (tampon) parcalar.push(tampon);
        tampon = s;
      } else {
        tampon = tampon ? tampon + "\r\n" + s : s;
      }
    }
    if (tampon) parcalar.push(tampon);

    var yazilan = 0;
    for (i = 0; i < parcalar.length && i < AYAR.raporParcaSayisi; i++) {
      var x = w + (20 + i * 70) / g.katsayi;
      var tx = M(lyr, "CreateArtisticText", x, h, parcalar[i]);
      if (tx) yazilan++;
      else break;
    }

    if (!yazilan) return null;
    if (parcalar.length > yazilan) {
      return katAd + " (" + yazilan + "/" + parcalar.length + " parca)";
    }
    return katAd;
  }

  function sor(metin, varsayilan) {
    try {
      if (typeof confirm === "function") return !!confirm(metin);
    } catch (e) { /* yok */ }
    return !!varsayilan;
  }

  function ilkBulguSatirlari(adet_) {
    var L = [], i;
    for (i = 0; i < g.bulgular.length && i < adet_; i++) {
      var f = g.bulgular[i];
      var nes = g.nesneler[f.n1];
      L.push((i + 1) + ") " + f.tur + "  " + sayi(f.uzunluk, 1) + " mm  " +
             (nes ? nes.ad : "?") + "  @ " + sayi(f.x, 1) + "," + sayi(f.y, 1));
    }
    if (g.bulgular.length > adet_) L.push("... (" + (g.bulgular.length - adet_) + " bulgu daha)");
    return L.join("\n");
  }

  // ------------------------------------------------------------------
  // Ana akis
  // ------------------------------------------------------------------

  function taraBaslangic() {
    var doc = host.ActiveDocument;
    if (!doc) return "Acik belge yok. Once bir .cdr dosyasi acin.";

    var kod = P(doc, "Unit");
    var bk = birimKatsayi(kod);
    g.katsayi = bk[0];
    g.birimAdi = bk[1];
    if (g.birimAdi === "piksel" || g.birimAdi === "bilinmiyor") {
      uyari("Belge birimi (" + g.birimAdi + ") mm'ye kesin cevrilemiyor.");
    }

    var page = P(doc, "ActivePage");
    if (!page) return "Aktif sayfa yok.";

    var kok = null;
    if (String(AYAR.kapsam).toLowerCase() === "secim") {
      kok = dene(function () { return host.ActiveSelectionRange; }, null);
      if (!kok || adet(kok) === 0) return "Secim bos. Nesneleri secin veya kapsam='sayfa' yapin.";
    } else {
      kok = P(page, "Shapes");
    }
    if (!kok) return "Sayfa nesneleri okunamadi.";

    var n = adet(kok), i;
    for (i = 1; i <= n; i++) {
      try {
        nesneTara(it(kok, i), 0);
      } catch (e) {
        g.atlanan++;
        uyari("Nesne " + i + " islenemedi: " + (e && e.message ? e.message : e));
      }
      if (g.cokBuyuk) break;
    }

    if (g.cokBuyuk) {
      return "Belge cok buyuk (" + g.segmentler.length + " segment siniri). " +
             "Kapsam='secim' ile parca parca calistirin veya AYAR.maxSegment'i artirin.";
    }

    var tamCift = yapisalTam();
    segmentTarama(tamCift);
    bulguTamamla();
    return "";
  }

  function ozetMetni() {
    var tam = 0, kismi = 0, yakin = 0, nokta = 0, uygun = 0, i;
    for (i = 0; i < g.bulgular.length; i++) {
      var t = g.bulgular[i].tur;
      if (t === "TAM") {
        tam++;
        var f = g.bulgular[i];
        if (f.n1 === f.n2 || ayniOzellik(f.n1, f.n2) === "") uygun++;
      }
      else if (t === "KISMI") kismi++;
      else if (t === "YAKIN") yakin++;
      else nokta++;
    }
    var s = "TAM (birebir)      : " + tam + "\n" +
            "KISMI (ortak kenar): " + kismi + "\n" +
            "YAKIN (kaymis)     : " + yakin + "\n" +
            "NOKTAS             : " + nokta + "\n" +
            "TOPLAM             : " + g.bulgular.length + "\n" +
            "Ortak kenar cifti  : " + g.tamSegmentler.length +
            (AYAR.ortakKenarSil ? "  (tek cizgiye iner)" : "  (temizlik kapali)") + "\n" +
            "Otomatik silmeye uygun TAM : " + uygun +
            (AYAR.otomatikSil ? "  (SILME ACIK)" : "  (silme kapali)") + "\n\n" +
            "Taranan: " + g.nesneler.length + " nesne / " +
            g.altYollar.length + " alt yol / " + g.segmentler.length + " segment";
    if (g.uyarilar.length) s += "\n\nUyari: " + g.uyarilar.slice(0, 2).join(" | ");
    return s;
  }

  function calistir() {
    var hata = taraBaslangic();
    if (hata) { alert("Cift Cizgi\n\n" + hata); return; }

    var doc = host.ActiveDocument;
    var page = P(doc, "ActivePage");

    if (g.bulgular.length === 0) {
      // Eski isaretler kalmasin (onceki calistirmalardan)
      isaretleriKaldir(page);
      alert("Cift Cizgi " + SURUM + "\n\nCift cizgi bulunmadi.\n\n" + ozetMetni() +
            "\n\nEski isaret katmanlari temizlendi.");
      return;
    }

    if (AYAR.asama === "tara") {
      alert("Cift Cizgi " + SURUM + " - SADECE TARAMA\n\n" + ozetMetni() +
            "\n\n(Isaret ve rapor bu asamada atlandi.)");
      return;
    }

    var isaret = isaretle(page);

    if (AYAR.asama === "isaret") {
      alert("Cift Cizgi " + SURUM + " - TARAMA + ISARET\n\n" + ozetMetni() +
            "\n\nIsaretlenen: " + isaret);
      return;
    }

    var rapor = (AYAR.asama === "rapor" || AYAR.raporMetni) ? null : raporYaz();

    var metin = "Cift Cizgi " + SURUM + "\n\n" + ozetMetni() + "\n\n" +
                "Isaretlenen: " + isaret + " (" + AYAR.isaretKatmani + " katmani)\n";

    if (AYAR.asama === "rapor" || AYAR.raporMetni) {
      var deneme = raporNesnesi(page, raporMetni());
      metin += deneme ? "Rapor metni: " + deneme + " katmaninda\n"
                      : "Rapor metni olusturulamadi\n";
    } else if (rapor) {
      metin += "Rapor: " + rapor.txt + "\n";
    } else {
      metin += "Ayrintili rapor metni kapali (AYAR.raporMetni = true ile acilir)\n";
    }

    metin += "\n" + ilkBulguSatirlari(40);

    var silN = 0, i;
    for (i = 0; i < g.bulgular.length; i++) {
      if (g.bulgular[i].tur === "TAM" && g.bulgular[i].eylem === "sil") silN++;
    }
    var ortakVar = g.tamSegmentler.length;

    if ((AYAR.otomatikSil && silN > 0) || (AYAR.ortakKenarSil && ortakVar > 0)) {
      var onay = sor(silN + " birebir kopya silinsin mi?", true);
      if (onay) {
        var s2 = sil();
        metin += "\nSilinen nesne: " + s2.nesne + "   alt yol: " + s2.altYol + "\n";
        if (s2.ortakSegment > 0) {
          metin += "Ortak kenar: " + s2.ortakSegment + " segment tek cizgiye indirildi (" +
                   s2.ortakNesne + " nesne yeniden kuruldu)\n";
        }
        if (s2.atlanan > 0) metin += "Elle birakilan: " + s2.atlanan + " nesne (egri/limit)\n";
        if (AYAR.silSonrasiTemizle && (s2.nesne + s2.altYol + s2.ortakNesne) > 0) {
          metin += "Isaret katmani da kaldirildi.\n";
        }
        metin += "Geri almak icin Ctrl+Z (silme + isaretler birlikte).";
      }
    } else if (silN > 0 || ortakVar > 0) {
      metin += "\n" + silN + " birebir kopya silinebilir (AYAR.otomatikSil = true ile).";
      if (ortakVar) metin += "\n" + ortakVar + " ortak kenar cifti bulundu (AYAR.ortakKenarSil = true ile tek cizgiye iner).";
    }

    alert(metin);
  }

  // Node testleri icin ic API (CorelDRAW icinde kullanilmaz)
  CDRCIFT._test = {
    g: g,
    tara: taraBaslangic,
    karsilastir: karsilastir,
    imza: imza,
    yuvarla: yuvarla,
    sil: sil,
    zincirler: zincirler,
    ortakKenarSil: ortakKenarSil,
    bulguTamamla: bulguTamamla,
    yapisalTam: yapisalTam,
    segmentTarama: segmentTarama,
    reset: function () {
      g.nesneler = [];
      g.altYollar = [];
      g.segmentler = [];
      g.tamSegmentler = [];
      g.bulgular = [];
      g.atlanan = 0;
      g.efekt = 0;
      g.powerclip = 0;
      g.uyarilar = [];
      g.kesildi = false;
      g.katsayi = 1;
      g.birimAdi = "mm";
    }
  };

  var corelVar = (typeof host !== "undefined" && host !== null);
  if (corelVar) {
    try {
      calistir();
    } catch (e) {
      if (typeof alert === "function") {
        alert("Cift Cizgi hata verdi:\n\n" + (e && e.message ? e.message : e) +
              "\n\nOnce CiftCizgi-Teshis.js calistirip ciktisini iletin.");
      }
    }
  }
})();

if (typeof module !== "undefined" && module.exports) {
  module.exports = CDRCIFT;
}
