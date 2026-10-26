// ==UserScript==
// @name         MOReiter Dexcom
// @namespace    https://github.com/libelle17/vb6wurzel
// @version      1.6
// @description  Dexcom Clarity (Praxis): Patient nach Name und Geburtsdatum suchen, dann Bericht speichern oder interaktive Berichte (Auftrag von MOReiter, Strg+D)
// @match        https://clarity.dexcom.eu/*
// @grant        none
// @updateURL    https://raw.githubusercontent.com/libelle17/MOReiter/main/MOReiter-Dexcom.user.js
// @downloadURL  https://raw.githubusercontent.com/libelle17/MOReiter/main/MOReiter-Dexcom.user.js
// @run-at       document-start
// ==/UserScript==

// MOReiter oeffnet die Patientenliste mit dem Auftrag hinter "#":
//     ...#mor=Nachname;Vorname;TT.MM.JJJJ;Kennung;Nr;PC-Name;Art;Tage   (Teile URL-kodiert)
// Art "speichern" (auf den PCs aus [Dexcom] SpeichernAuf=): "Bericht speichern oder drucken", Zeitraum Tage
// (7, 14, 30 oder 90), "Bericht speichern" (PDF-Download); sonst "Zu interaktiven Berichten wechseln".
// Fortschritt im Seitentitel "MOR:Nr:Zustand" wie beim LibreView-Skript:
//   passwort / passworttippen / anmelden / anmeldefehler   Anmeldung (hoechstens ein Versuch je Auftrag)
//   suche       angemeldet, Patient wird gesucht
//   patient     Patient gefunden, sein Fenster ist offen
//   fertig      interaktive Berichte geoeffnet und "Herunterladen" gedrueckt bzw. "Bericht speichern" gedrueckt
//               (MOReiter ist fertig);
//               beim Speichern wartet das Skript noch, bis der Bericht erstellt ist
//   gedruckt    Bericht erstellt, Download laeuft
//   nichts      Patient nicht gefunden
//   fehler ...  sonstiger Fehler
(function () {
  'use strict';
  var schluessel = 'MOReiterDexcom';
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
    sessionStorage.setItem(schluessel, JSON.stringify(a));
    history.replaceState(history.state, '', location.pathname + location.search);
    return true;
  }

  function auftrag() {
    try { return JSON.parse(sessionStorage.getItem(schluessel) || 'null'); } catch (e) { return null; }
  }
  function merken(a) { sessionStorage.setItem(schluessel, JSON.stringify(a)); }
  function erledigt() { sessionStorage.removeItem(schluessel); }

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

  // Wert so setzen, dass das Rahmenwerk (Ember) ihn mitbekommt
  function setze(el, v) {
    var proto = el.tagName == 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, v);
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'a' }));
  }

  // Klick wie mit der Maus
  function klicke(el) {
    try { el.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (e) {}
    var r = el.getBoundingClientRect(), x = r.left + r.width / 2, y = r.top + r.height / 2;
    var opt = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, button: 0, buttons: 1 };
    if (window.PointerEvent) el.dispatchEvent(new PointerEvent('pointerdown', Object.assign({ pointerId: 1, isPrimary: true }, opt)));
    el.dispatchEvent(new MouseEvent('mousedown', opt));
    opt.buttons = 0;
    if (window.PointerEvent) el.dispatchEvent(new PointerEvent('pointerup', Object.assign({ pointerId: 1, isPrimary: true }, opt)));
    el.dispatchEvent(new MouseEvent('mouseup', opt));
    el.dispatchEvent(new MouseEvent('click', opt));
  }

  function sichtbar(el) { return !!(el && (el.offsetWidth || el.offsetHeight || el.getClientRects().length)); }
  function klein(s) { return (s || '').toLocaleLowerCase('de').replace(/\s+/g, ' ').trim(); }
  function aktiv(b) { return b && !b.disabled && !b.hasAttribute('disabled'); }

  // Knopf/Link mit genau diesem Text (sichtbar)
  function mitText(sel, text) {
    var l = document.querySelectorAll(sel);
    for (var i = 0; i < l.length; i++) if (klein(l[i].textContent) == klein(text) && sichtbar(l[i])) return l[i];
    return null;
  }

  // Vornamen: Bindestrich wie Leerzeichen; es genuegt, dass alle Teile des kuerzeren im laengeren vorkommen
  // ("Hans-Peter" in MO passt zu "Hans Peter", "Hans" oder "Peter" im Portal); das Geburtsdatum entscheidet mit
  function teile(s) { return klein(s).split(/[\s-]+/).filter(function (x) { return x != ''; }); }
  function namePasst(a, nach, vor) {
    if (teile(a.name).join(' ') != teile(nach).join(' ')) return false;
    var v1 = teile(a.vor), v2 = teile(vor);
    if (!v1.length || !v2.length) return false;
    var kurz = v1.length <= v2.length ? v1 : v2, lang = kurz === v1 ? v2 : v1;
    return kurz.every(function (x) { return lang.indexOf(x) >= 0; });
  }

  var monate = ['januar', 'februar', 'maerz', 'april', 'mai', 'juni', 'juli', 'august', 'september',
                'oktober', 'november', 'dezember'];
  var monateEn = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  function monat(w) {
    for (var i = 0; i < 12; i++) if (monate[i].indexOf(w) == 0 || w.indexOf(monate[i].slice(0, 3)) == 0 ||
        w.indexOf(monateEn[i]) == 0) return i;
    return -1;
  }
  function zwei(n) { return (n < 10 ? '0' : '') + n; }

  // alle Datumsangaben eines Textes als TT.MM.JJJJ: T.M.JJJJ, T/M/JJJJ, "9. Aug. 1975", "9 Juni 1975"
  function daten(text) {
    var erg = [], m, re1 = /(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})/g;
    while ((m = re1.exec(text))) erg.push(zwei(+m[1]) + '.' + zwei(+m[2]) + '.' + m[3]);
    var t = klein(text).replace(/ä/g, 'ae'), re2 = /(\d{1,2})\.?\s+([a-z]{3,})\.?,?\s+(\d{4})/g;
    while ((m = re2.exec(t))) { var i = monat(m[2]); if (i >= 0) erg.push(zwei(+m[1]) + '.' + zwei(i + 1) + '.' + m[3]); }
    return erg;
  }

  // TT.MM.JJJJ als vergleichbare Zahl JJJJMMTT, 0 = keins
  function datumZahl(d) {
    if (!d) return 0;
    var t = d.split('.');
    return +(t[2] + t[1] + t[0]);
  }

  // --- Seiten ----------------------------------------------------------------------------------------
  // Cookie-Hinweis
  function cookies(a) {
    var b = mitText('button', 'Agree and Proceed') || mitText('button', 'Zustimmen und fortfahren');
    if (b && einmal(a, 'cookies')) klicke(b);
    return false;
  }

  // Anmeldung: hoechstens ein Versuch je Auftrag. Mit Kennung ([Dexcom] Kennung=) wie bei LibreView: nur sie,
  // Liste von Firefox ueber "passworttippen"; ohne Kennung die erste gespeicherte Anmeldung ("passwort")
  var stand = '', standSeit = 0;
  function anmelden(a) {
    var em = $('#username'), pw = $('#password'), b = $('[data-testid="login-form__submit"]');
    if (!em || !pw || !b) return false;
    if (seit(a, 'anmelden') >= 0) {
      // nach 15 s noch immer die Anmeldeseite: abgelehnt
      titel(a, seit(a, 'anmelden') > 15000 ? 'anmeldefehler' : 'anmelden');
      if (seit(a, 'anmelden') > 15000) erledigt();
      return true;
    }
    if (seit(a, 'loginSeite') < 0) einmal(a, 'loginSeite');
    if (seit(a, 'loginSeite') < 1500) { titel(a, 'anmelden'); return true; }
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
      klicke(b);
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

  // Schutz vor ungewolltem Abschicken, solange ein Auftrag laeuft: nur mit Passwort (und ggf. Kennung)
  function abschickenPruefen(ev) {
    var a = auftrag(), em = $('#username'), pw = $('#password'), t = ev.target;
    if (!a || !em || !pw || !t) return;
    var login = (t.closest && t.closest('[data-testid="login-form__submit"]')) ||
      (ev.type == 'submit' && t.querySelector && t.querySelector('#username'));
    var ok = pw.value != '' && (a.kennung == '' || klein(em.value) == klein(a.kennung)) && seit(a, 'anmelden') < 1000;
    if (login && !ok) {
      ev.preventDefault();
      ev.stopImmediatePropagation();
      console.log('MOReiter Dexcom: Anmelden ohne passende Kennung/Passwort oder zum zweiten Mal verhindert');
    }
  }
  window.addEventListener('submit', abschickenPruefen, true);
  window.addEventListener('click', abschickenPruefen, true);

  // Fenster "Bericht speichern oder drucken" (Patientenliste) bzw. "Herunterladen" (interaktive Berichte), beide
  // gleich aufgebaut: Zeitraum, dann "Bericht speichern"; danach warten, bis erstellt. Gesucht wird im Fenster,
  // die interaktiven Berichte haben dahinter eine eigene Zeitraumleiste
  function berichtFenster(a) {
    var h = mitText('h1', 'Bericht speichern oder drucken') || mitText('h1', 'Herunterladen');
    if (!h) return false;
    var fenster = h.closest('dialog-window, .dialog-window, .modal-dialog, [role="dialog"]') || h.parentElement || document;
    function imFenster(sel, text) {
      var l = fenster.querySelectorAll(sel);
      for (var i = 0; i < l.length; i++) if ((!text || klein(l[i].textContent) == klein(text)) && sichtbar(l[i])) return l[i];
      return null;
    }
    if (seit(a, 'gespeichert') >= 0) {
      titel(a, mitZeitraum(a, 'fertig'));
      if (/Ihr Bericht wurde erstellt/i.test(document.body.innerText)) { titel(a, mitZeitraum(a, 'gedruckt')); erledigt(); freigeben(); }
      return true;
    }
    titel(a, 'patient');
    var tage = a.tage;
    if (tage && !a.schritte.zeitraum) {
      var leiste = imFenster('.date-range-picker__bar');
      var knopf = imFenster('button.btn-sm', tage) || mitText('button.btn-sm', tage);
      if (knopf) { klicke(knopf); einmal(a, 'zeitraum'); return true; }
      if (leiste && einmal(a, 'zeitraumAuf')) { klicke(leiste); return true; }
      // Auswahl nicht zu finden: nach 5 s mit dem voreingestellten Zeitraum weiter
      if (seit(a, 'zeitraumAuf') < 5000) return true;
      einmal(a, 'zeitraum');
    }
    var sp = imFenster('button', 'Bericht speichern');
    if (aktiv(sp) && (!tage || seit(a, 'zeitraum') > 1000)) {
      // eingestellter Zeitraum, z.B. "Sa. 17. Apr. 2021 - Fr. 30. Apr. 2021"
      var zr = imFenster('.date-range-picker__date-range'), d = zr ? daten(zr.textContent) : [];
      if (d.length >= 2) a.zeitraum = d[0] + '-' + d[1];
      einmal(a, 'gespeichert');
      klicke(sp);
      titel(a, mitZeitraum(a, 'fertig'));
    }
    return true;
  }

  // Zeitraum des Berichts an die Fertigmeldung haengen ("fertig|TT.MM.JJJJ-TT.MM.JJJJ"), MOReiter setzt ihn in
  // den Dateinamen ein
  function mitZeitraum(a, z) { return a.zeitraum ? z + '|' + a.zeitraum : z; }

  // nach dem Download ist der Tab nach 60 s wieder frei fuer MOReiter
  function freigeben() {
    setTimeout(function () { if (/^MOR:\d+:(gedruckt|fertig)/.test(document.title)) document.title = 'Dexcom Clarity Clinic Portal'; }, 60000);
  }

  // Fenster des Patienten (nach Klick auf seine Zeile): "Bericht speichern oder drucken" bzw. interaktiv
  function patientFenster(a) {
    var bericht = $('[data-testid="patient-generate-report__button"]'), interaktiv = $('.patient-view-data__button');
    if (!sichtbar(bericht) && !sichtbar(interaktiv)) return false;
    titel(a, 'patient');
    if (a.art == 'speichern') {
      if (bericht && einmal(a, 'berichtKlick')) klicke(bericht);
    } else if (interaktiv && einmal(a, 'interaktivKlick')) {
      klicke(interaktiv);
    }
    return true;
  }

  // interaktive Berichte: in der Symbolleiste ueber dem Bericht "Herunterladen" druecken; oeffnet sich dann
  // "Bericht speichern oder drucken", erledigt das berichtFenster, sonst ist der Auftrag nach 30 s zu Ende
  function berichtSeite(a) {
    var b = $('.report-icon-bar__download-icon-button');
    if (!sichtbar(b) || seit(a, 'interaktivKlick') < 0) return false;
    titel(a, 'fertig');
    if (aktiv(b) && einmal(a, 'herunterladenKlick')) klicke(b);
    else if (seit(a, 'herunterladenKlick') > 30000) { erledigt(); freigeben(); }
    return true;
  }

  // Patientenliste: suchen, passende Zeile (Name "Nachname, Vorname", Geburtsdatum) anklicken
  // Suchtexte: Nachname, "Nachname, Vorname", "Vorname Nachname", bei mehrteiligen Vornamen (Hans-Peter)
  // noch "Nachname, Hans" und "Nachname, Peter"
  function suchTexte(a) {
    var l = [a.name, a.name + ', ' + a.vor, a.vor + ' ' + a.name], v = teile(a.vor);
    if (v.length > 1) v.forEach(function (x) { l.push(a.name + ', ' + x.charAt(0).toUpperCase() + x.slice(1)); });
    return l;
  }
  function liste(a) {
    var feld = $('input[type="search"]');
    var zeilen = document.querySelectorAll('tr.patient-list-row, tr[data-testid="patient-list-row"]');
    if (!feld && !zeilen.length) return false;
    if (seit(a, 'zeileKlick') < 0) titel(a, 'suche');
    // derselbe Patient steht oefter doppelt in der Liste: dann die Zeile mit dem juengsten "Letzter Upload"
    var bester = null, besterN = null, besterStand = -1;
    for (var i = 0; i < zeilen.length; i++) {
      var z = zeilen[i], n = z.querySelector('.patient-list__patient-name'), d = z.querySelector('.patient-list__dob');
      if (!n || !d) continue;
      var teile = n.textContent.split(',');
      if (teile.length < 2 || !namePasst(a, teile[0], teile.slice(1).join(','))) continue;
      if (daten(d.textContent).indexOf(a.geb) < 0) continue;
      var up = z.querySelector('.patient-list__last-upload'), st = up ? datumZahl(daten(up.textContent)[0]) : 0;
      if (st > besterStand) { bester = z; besterN = n; besterStand = st; }
    }
    if (bester) {
      titel(a, 'patient');
      if (einmal(a, 'zeileKlick')) klicke(besterN);
      else if (seit(a, 'zeileKlick') > 5000 && einmal(a, 'zeileKlick2')) klicke(bester);
      return true;
    }
    if (!feld) return true;
    if (seit(a, 'liste') < 0) { einmal(a, 'liste'); return true; }
    var k = a.schritte.suchNr || 0;
    if (k < suchTexte(a).length && (k == 0 || seit(a, 'suche' + (k - 1)) > 4000) && einmal(a, 'suche' + k)) {
      a.schritte.suchNr = k + 1;
      merken(a);
      feld.focus();
      setze(feld, suchTexte(a)[k]);
      return true;
    }
    if (k >= suchTexte(a).length && seit(a, 'suche' + (suchTexte(a).length - 1)) > 5000 && seit(a, 'zeileKlick') < 0) {
      titel(a, 'nichts');
      erledigt();
    }
    return true;
  }

  // --- Ablauf ----------------------------------------------------------------------------------------
  function schritt() {
    var a = auftrag();
    if (!a) return;
    if (Date.now() - a.beginn > hoechstens) { erledigt(); return; }
    if (!document.body) return;
    try {
      cookies(a);
      anmelden(a) || berichtFenster(a) || berichtSeite(a) || patientFenster(a) || liste(a) || titel(a, zustand || 'suche');
    } catch (e) {
      titel(a, 'fehler ' + e.message);
      erledigt();
    }
  }

  auftragAusAdresse();
  window.addEventListener('hashchange', function () { if (auftragAusAdresse()) zustand = ''; });
  setInterval(schritt, takt);
})();
