// ==UserScript==
// @name         MOReiter EasyView
// @namespace    https://github.com/libelle17/vb6wurzel
// @version      1.2
// @description  Medtrum EasyView Pro (Praxis): Patient nach Name (und Alter) suchen, Dashboard ueber Tage drucken (Auftrag von MOReiter, Strg+E)
// @match        https://easyview.medtrum.eu/*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @updateURL    https://raw.githubusercontent.com/libelle17/MOReiter/main/MOReiter-EasyView.user.js
// @downloadURL  https://raw.githubusercontent.com/libelle17/MOReiter/main/MOReiter-EasyView.user.js
// @run-at       document-start
// ==/UserScript==

// MOReiter oeffnet EasyView mit dem Auftrag hinter "#":
//     ...#mor=Nachname;Vorname;TT.MM.JJJJ;Kennung;Nr;PC-Name;Art;Tage   (Teile URL-kodiert)
// Die Seite lenkt ihre Unterseiten selbst ueber "#/..." (Vue), daher wird der Auftrag gleich beim Laden
// abgeholt und die Adresse auf "#/select" gesetzt (wie die Weiterleitung von www.easyviewpro.com). Ablauf: "NEW VERSION", Anmeldung (mit "Remember Me"),
// Patientenliste, Suche nach dem Nachnamen, Zeile mit dem Namen (Vor- und Nachname, auch ohne Leerzeichen
// dazwischen) und, wenn angegeben, dem Alter, Dashboard, Zeitraum Tage (7, 14, 30 oder 90), "Print". Den
// Druckdialog von Firefox und "Speichern unter" bedient MOReiter. Fortschritt im Seitentitel "MOR:Nr:Zustand":
//   passwort / passworttippen / anmelden / anmeldefehler   Anmeldung (hoechstens ein Versuch je Auftrag)
//   suche, patient, drucken|TT.MM.JJJJ-TT.MM.JJJJ (Druck ausgeloest), gedruckt, nichts, liste, fehler ...
(function () {
  'use strict';
  var schluessel = 'MOReiterEasyView';
  var takt = 400, hoechstens = 30 * 60 * 1000;

  // --- Auftrag ---------------------------------------------------------------------------------------
  function teil(s) {
    try { return decodeURIComponent(s.replace(/\+/g, ' ')).trim(); } catch (e) { return s.trim(); }
  }

  function auftragAusAdresse() {
    var m = location.hash.match(/^#mor=(.*)$/);
    if (!m) return false;
    var t = m[1].split(';').map(teil);
    var a = { name: t[0] || '', vor: t[1] || '', geb: t[2] || '', kennung: t[3] || '', nr: t[4] || '',
              pc: t[5] || '', art: t[6] || '', tage: t[7] || '', beginn: Date.now(), schritte: {} };
    merken(a);
    // wie die Weiterleitung von www.easyviewpro.com mit der Auswahl NEW/OLD VERSION beginnen ("#/" fuehrte
    // zur Anmeldung von EasyView Personal fuer Patienten)
    history.replaceState(history.state, '', location.pathname + location.search + '#/select');
    return true;
  }

  function auftrag() {
    var a = null;
    try { a = JSON.parse(sessionStorage.getItem(schluessel) || 'null'); } catch (e) {}
    if (a) return a;
    try { a = JSON.parse(GM_getValue(schluessel, 'null')); } catch (e) {}
    // nur einen frischen Auftrag uebernehmen, nicht in einem spaeter von Hand geoeffneten Tab
    if (a && Date.now() - (a.gemerkt || 0) < 2 * 60 * 1000) {
      sessionStorage.setItem(schluessel, JSON.stringify(a));
      return a;
    }
    return null;
  }
  function merken(a) {
    a.gemerkt = Date.now();
    sessionStorage.setItem(schluessel, JSON.stringify(a));
    GM_setValue(schluessel, JSON.stringify(a));
  }
  function erledigt() { sessionStorage.removeItem(schluessel); GM_deleteValue(schluessel); }

  function einmal(a, name) {
    if (a.schritte[name]) return false;
    a.schritte[name] = Date.now();
    merken(a);
    return true;
  }
  function seit(a, name) { return a.schritte[name] ? Date.now() - a.schritte[name] : -1; }

  var zustand = '';
  function titel(a, z) {
    zustand = z;
    var t = 'MOR:' + a.nr + ':' + z;
    if (document.title != t) document.title = t;
  }

  // --- Hilfen ----------------------------------------------------------------------------------------
  function $(sel) { return document.querySelector(sel); }

  // Wert so setzen, dass das Rahmenwerk (Vue) ihn mitbekommt
  function setze(el, v) {
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, v);
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }
  function enter(el) {
    ['keydown', 'keypress', 'keyup'].forEach(function (typ) {
      el.dispatchEvent(new KeyboardEvent(typ, { bubbles: true, key: 'Enter', code: 'Enter', keyCode: 13, which: 13 }));
    });
  }

  // Klick wie mit der Maus (ohne "view", s. MOReiter-YourLoops.user.js)
  function klicke(el) {
    try { el.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (e) {}
    var r = el.getBoundingClientRect(), x = r.left + r.width / 2, y = r.top + r.height / 2;
    var opt = { bubbles: true, cancelable: true, clientX: x, clientY: y, button: 0, buttons: 1 };
    if (window.PointerEvent) el.dispatchEvent(new PointerEvent('pointerdown', Object.assign({ pointerId: 1, isPrimary: true }, opt)));
    el.dispatchEvent(new MouseEvent('mousedown', opt));
    opt.buttons = 0;
    if (window.PointerEvent) el.dispatchEvent(new PointerEvent('pointerup', Object.assign({ pointerId: 1, isPrimary: true }, opt)));
    el.dispatchEvent(new MouseEvent('mouseup', opt));
    el.dispatchEvent(new MouseEvent('click', opt));
  }

  function sichtbar(el) { return !!(el && (el.offsetWidth || el.offsetHeight || el.getClientRects().length)); }
  function klein(s) { return (s || '').toLocaleLowerCase('de').replace(/\s+/g, ' ').trim(); }
  function mitText(sel, text) {
    var l = document.querySelectorAll(sel);
    for (var i = 0; i < l.length; i++) if (klein(l[i].textContent) == klein(text) && sichtbar(l[i])) return l[i];
    return null;
  }

  // Name nur aus Buchstaben, Umlaute ausgeschrieben: "Marko Ciric" und "MarkoCiric" sind gleich
  function buchstaben(s) {
    return klein(s).replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ß/g, 'ss')
      .replace(/[^a-z]/g, '');
  }
  function vornamen(a) {
    var l = [a.vor], t = klein(a.vor).split(/[\s-]+/).filter(function (x) { return x != ''; });
    if (t.length > 1) l.push(t[0]);
    return l;
  }
  function listenNamePasst(a, text) {
    var t = buchstaben(text), n = buchstaben(a.name);
    return vornamen(a).some(function (v) { v = buchstaben(v); return t == v + n || t == n + v; });
  }

  // Alter heute nach dem Geburtsdatum TT.MM.JJJJ
  function alter(geb) {
    var t = geb.split('.'), h = new Date();
    if (t.length != 3) return -1;
    var j = h.getFullYear() - (+t[2]);
    if (h.getMonth() + 1 < +t[1] || (h.getMonth() + 1 == +t[1] && h.getDate() < +t[0])) j--;
    return j;
  }

  // --- Seiten ----------------------------------------------------------------------------------------
  // Auswahl "NEW VERSION" / "OLD VERSION"
  function versionSeite(a) {
    var b = $('.middle-second .btn:not(.btn2)');
    if (!sichtbar(b)) return false;
    titel(a, 'anmelden');
    if (einmal(a, 'version')) klicke(b);
    else if (seit(a, 'version') > 5000 && einmal(a, 'version2')) klicke(b);
    return true;
  }

  // Anmeldung wie bei YourLoops: hoechstens ein Versuch je Auftrag; mit Kennung nur diese (Firefox-Liste ueber
  // "passworttippen"), sonst die erste gespeicherte Anmeldung ("passwort"); "Remember Me" wird angehakt
  var stand = '', standSeit = 0;
  function loginTeile() {
    // nur die Anmeldung von EasyView Pro (#/hcplogin/...), nicht die von EasyView Personal (#/login/...)
    if (!/^#\/hcplogin/i.test(location.hash)) return null;
    var em = $('input[placeholder="Email"]'), pw = $('input[type="password"]'), b = mitText('button', 'Sign in');
    return (em && pw && b && sichtbar(em)) ? { em: em, pw: pw, b: b } : null;
  }
  function anmelden(a) {
    var l = loginTeile();
    if (!l) return false;
    var em = l.em, pw = l.pw;
    if (seit(a, 'anmelden') >= 0) {
      titel(a, seit(a, 'anmelden') > 15000 ? 'anmeldefehler' : 'anmelden');
      if (seit(a, 'anmelden') > 15000) erledigt();
      return true;
    }
    if (seit(a, 'loginSeite') < 0) einmal(a, 'loginSeite');
    if (seit(a, 'loginSeite') < 1500) { titel(a, 'anmelden'); return true; }
    var merk = $('label.fourth-model-chechbox');
    if (merk && !/is-checked/.test(merk.className) && einmal(a, 'merken')) klicke(merk);
    if (a.kennung != '') {
      var ohneLetztes = a.kennung.slice(0, -1);
      if (klein(em.value) != klein(a.kennung)) {
        if (klein(em.value) == klein(ohneLetztes) && pw.value == '') {
          if (document.activeElement !== em) em.focus();
          titel(a, 'passworttippen');
          return true;
        }
        setze(em, ohneLetztes);
        if (pw.value != '') setze(pw, '');
        stand = '';
        titel(a, 'anmelden');
        return true;
      }
    }
    var jetzt = em.value + '\u0001' + pw.value;
    if (jetzt != stand) { stand = jetzt; standSeit = Date.now(); }
    if (em.value != '' && pw.value != '' && Date.now() - standSeit >= 1000) {
      einmal(a, 'anmelden');
      titel(a, 'anmelden');
      klicke(l.b);
      return true;
    }
    if (a.kennung != '') {
      if (pw.value == '' && Date.now() - standSeit > 3000) { setze(em, a.kennung.slice(0, -1)); stand = ''; titel(a, 'passworttippen'); }
      else titel(a, 'anmelden');
      return true;
    }
    if (pw.value == '' && einmal(a, 'passwortFrage')) em.focus();
    titel(a, pw.value == '' ? 'passwort' : 'anmelden');
    return true;
  }

  // Schutz vor ungewolltem Anmelden, solange ein Auftrag laeuft: nur mit Passwort (und ggf. Kennung), einmal
  function abschickenPruefen(ev) {
    var a = auftrag(), l = loginTeile(), t = ev.target;
    if (!a || !l || !t || !t.closest || t.closest('button') !== l.b) return;
    var ok = l.pw.value != '' && (a.kennung == '' || klein(l.em.value) == klein(a.kennung)) && seit(a, 'anmelden') < 1000;
    if (!ok) {
      ev.preventDefault();
      ev.stopImmediatePropagation();
      console.log('MOReiter EasyView: Anmelden ohne passende Kennung/Passwort oder zum zweiten Mal verhindert');
    }
  }
  window.addEventListener('click', abschickenPruefen, true);

  // Anmeldung von EasyView Personal (fuer Patienten): zurueck zur Auswahl NEW/OLD VERSION
  function personalSeite(a) {
    if (!/^#\/login\//i.test(location.hash)) return false;
    titel(a, 'anmelden');
    if (einmal(a, 'personal')) location.hash = '#/select';
    else if (seit(a, 'personal') > 5000 && einmal(a, 'personal2')) location.hash = '#/select';
    else if (seit(a, 'personal') > 15000) { titel(a, 'fehler EasyView Personal statt Pro'); erledigt(); }
    return true;
  }

  // Startseite nach der Anmeldung: weiter zur Patientenliste
  function startSeite(a) {
    if (!/^#\/HomePage/i.test(location.hash)) return false;
    titel(a, 'suche');
    if (einmal(a, 'zurListe')) location.hash = '#/patientManage';
    else if (seit(a, 'zurListe') > 5000 && einmal(a, 'zurListe2')) location.hash = '#/patientManage';
    return true;
  }

  // Patientenliste: passende Zeile anklicken; sonst nach dem Nachnamen suchen. Mehrere passende: Liste stehen
  // lassen ("liste"), keine: "nichts" (der Auftrag bleibt 3 min fuer eine Auswahl von Hand)
  function liste(a) {
    var feld = $('input[placeholder="Name/Account"]');
    if (!feld || !sichtbar(feld)) return false;
    if (seit(a, 'zeileKlick') < 0) titel(a, 'suche');
    if (seit(a, 'listeDa') < 0) { einmal(a, 'listeDa'); return true; }
    if (seit(a, 'listeDa') < 1500) return true;
    var zeilen = document.querySelectorAll('tr.el-table__row'), passend = [], soll = alter(a.geb);
    for (var i = 0; i < zeilen.length; i++) {
      var n = zeilen[i].querySelector('.name .span1'), zellen = zeilen[i].querySelectorAll('td .cell');
      if (!n || !listenNamePasst(a, n.textContent)) continue;
      // Spalten: Auswahl, Name, Geschlecht, Alter, ...; ein angegebenes Alter muss passen
      var al = zellen.length > 3 ? zellen[3].textContent.trim() : '';
      if (/^\d+$/.test(al) && soll >= 0 && +al != soll) continue;
      passend.push(n);
    }
    if (passend.length == 1) {
      titel(a, 'patient');
      if (einmal(a, 'zeileKlick')) klicke(passend[0]);
      else if (seit(a, 'zeileKlick') > 5000 && einmal(a, 'zeileKlick2')) klicke(passend[0].closest('tr') || passend[0]);
      return true;
    }
    if (passend.length > 1) { titel(a, 'liste'); erledigt(); return true; }
    if (einmal(a, 'suche')) {
      feld.focus();
      setze(feld, a.name);
      enter(feld);
      var lupe = feld.parentNode && feld.parentNode.querySelector('.el-icon-search');
      if (lupe) klicke(lupe);
      return true;
    }
    if (seit(a, 'suche') > 5000 && seit(a, 'zeileKlick') < 0) {
      einmal(a, 'nichts');
      titel(a, 'nichts');
      if (seit(a, 'nichts') > 180000) erledigt();
    }
    return true;
  }

  // Zeitraum aus den beiden Datumsfeldern ("21-07-2026") als "21.07.2026-18.10.2026"
  function zeitraum() {
    var l = document.querySelectorAll('.time-select input'), d = [];
    for (var i = 0; i < l.length; i++) {
      var m = (l[i].value || '').match(/^(\d\d)-(\d\d)-(\d{4})$/);
      if (m) d.push(m[1] + '.' + m[2] + '.' + m[3]);
    }
    return d.length == 2 ? d[0] + '-' + d[1] : '';
  }

  function laedt() {
    var l = document.querySelectorAll('.el-loading-mask');
    for (var i = 0; i < l.length; i++) if (sichtbar(l[i]) && getComputedStyle(l[i]).display != 'none') return true;
    return false;
  }

  // nach dem Druck ist der Tab nach 60 s wieder frei fuer MOReiter
  function freigeben() {
    setTimeout(function () { if (/^MOR:\d+:(gedruckt|drucken)/.test(document.title)) document.title = 'EasyView'; }, 60000);
  }

  // Dashboard des Patienten: Zeitraum waehlen, warten bis geladen, "drucken|Zeitraum" melden, dann "Print".
  // window.print() haelt die Seite an, bis der Druckdialog zu ist; danach "gedruckt"
  function dashboard(a) {
    var druck = $('.print-span');
    if (!/^#\/dashboard/i.test(location.hash) || !sichtbar(druck)) return false;
    if (seit(a, 'druck') >= 0) {
      titel(a, 'gedruckt' + (a.zeitraum ? '|' + a.zeitraum : ''));
      erledigt();
      freigeben();
      return true;
    }
    titel(a, 'patient');
    var tage = { '7': '7 days', '14': '14 days', '30': '30 days', '90': '90 days' }[a.tage] || '90 days';
    var chip = mitText('.select-title-box .select-span', tage);
    if (!chip) return true;
    if (!/selected-span/.test(chip.className)) {
      if (einmal(a, 'zeitraum')) { a.vorher = zeitraum(); merken(a); klicke(chip); }
      else if (seit(a, 'zeitraum') > 4000 && einmal(a, 'zeitraum2')) klicke(chip);
      return true;
    }
    // nach dem Wechsel: neue Daten abwarten
    var zr = zeitraum();
    if (!zr || laedt()) return true;
    if (seit(a, 'zeitraum') >= 0 && (seit(a, 'zeitraum') < 2500 || (zr == a.vorher && seit(a, 'zeitraum') < 15000))) return true;
    if (seit(a, 'geladen') < 0) { einmal(a, 'geladen'); return true; }
    if (seit(a, 'geladen') < 2000) return true;
    if (a.zeitraum != zr) { a.zeitraum = zr; merken(a); }
    if (seit(a, 'druckGemeldet') < 0) { einmal(a, 'druckGemeldet'); titel(a, 'drucken|' + zr); return true; }
    titel(a, 'drucken|' + zr);
    if (seit(a, 'druckGemeldet') < 1500) return true;
    // vor dem Klick vermerken: window.print() kehrt erst nach dem Druckdialog zurueck
    einmal(a, 'druck');
    klicke(druck);
    return true;
  }

  // --- Ablauf ----------------------------------------------------------------------------------------
  function schritt() {
    var a = auftrag();
    if (!a) return;
    if (Date.now() - a.beginn > hoechstens) { erledigt(); return; }
    if (!document.body) return;
    try {
      versionSeite(a) || personalSeite(a) || anmelden(a) || startSeite(a) || dashboard(a) || liste(a) || titel(a, zustand || 'suche');
    } catch (e) {
      titel(a, 'fehler ' + e.message);
      erledigt();
    }
  }

  auftragAusAdresse();
  window.addEventListener('hashchange', function () { if (auftragAusAdresse()) zustand = ''; });
  setInterval(schritt, takt);
})();
