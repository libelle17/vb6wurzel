// ==UserScript==
// @name         MOReiter LibreView
// @namespace    https://github.com/libelle17/vb6wurzel
// @version      1.1
// @description  LibreView: Patient nach Name und Geburtsdatum suchen und seine Glukose-Berichte zeigen (Auftrag von MOReiter, Strg+L)
// @match        https://www.libreview.com/*
// @grant        none
// @run-at       document-start
// ==/UserScript==

// MOReiter oeffnet eine LibreView-Seite mit dem Auftrag hinter "#":
//     ...#mor=Nachname;Vorname;TT.MM.JJJJ;Kennung;Nr;PC-Name   (Teile URL-kodiert)
// Der Auftrag wird im sessionStorage des Tabs gemerkt (uebersteht Anmeldung und Seitenwechsel). LibreView
// wechselt die Seiten meist ohne Neuladen, daher schaut das Skript laufend nach, welche Seite gerade zu
// sehen ist, und erledigt dort den naechsten Schritt. Fortschritt im Seitentitel "MOR:Nr:Zustand":
//   land        Wohnsitzland/Sprache gewaehlt
//   passwort    Anmeldeseite ohne Passwort: E-Mail-Feld hat den Fokus (MOReiter waehlt die gespeicherte
//               Anmeldung aus der Firefox-Liste)
//   anmelden    Anmeldung abgeschickt (hoechstens einmal je Auftrag, sonst droht eine Sperre)
//   anmeldefehler  LibreView hat E-Mail/Passwort abgelehnt
//   zweifaktor  Zwei-Faktor-Anmeldung: Code per E-Mail angefordert
//   code        Code-Seite: "14 Tage merken" angehakt, PC-Name eingetragen, der Mensch tippt den Code
//   suche       angemeldet, Patient wird gesucht
//   patient     Patient gefunden, seine Seite wird geoeffnet
//   fertig      Glukose-Berichte werden geoeffnet
//   keinedaten  Patient ohne Glukosedaten
//   liste       kein Patient mit passendem Geburtsdatum: Suchergebnis bleibt stehen
//   fehler ...  sonstiger Fehler
(function () {
  'use strict';
  var schluessel = 'MOReiterLibreView';
  var takt = 400, hoechstens = 10 * 60 * 1000;

  // --- Auftrag ---------------------------------------------------------------------------------------
  function teil(s) {
    try { return decodeURIComponent(s.replace(/\+/g, ' ')).trim(); } catch (e) { return s.trim(); }
  }

  function auftragAusAdresse() {
    var m = location.hash.match(/^#mor=(.*)$/);
    if (!m) return false;
    var t = m[1].split(';').map(teil);
    var a = { name: t[0] || '', vor: t[1] || '', geb: t[2] || '', kennung: t[3] || '', nr: t[4] || '',
              pc: t[5] || '', beginn: Date.now(), schritte: {} };
    sessionStorage.setItem(schluessel, JSON.stringify(a));
    history.replaceState(history.state, '', location.pathname + location.search);
    return true;
  }

  function auftrag() {
    try { return JSON.parse(sessionStorage.getItem(schluessel) || 'null'); } catch (e) { return null; }
  }

  function merken(a) { sessionStorage.setItem(schluessel, JSON.stringify(a)); }
  function erledigt() { sessionStorage.removeItem(schluessel); }

  // einmalige Schritte (Klick auf Anmelden usw.) je Auftrag merken
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
  function $(id) { return document.getElementById(id); }

  // Wert so setzen, dass React ihn mitbekommt
  function setze(el, v) {
    var proto = el.tagName == 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, v);
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function sichtbar(el) { return !!(el && (el.offsetWidth || el.offsetHeight || el.getClientRects().length)); }

  function klein(s) { return (s || '').toLocaleLowerCase('de').replace(/\s+/g, ' ').trim(); }

  // Vorname: "Hans Peter" in MO passt zu "Hans" in LibreView und umgekehrt
  function namePasst(a, nach, vor) {
    var n1 = klein(a.name), n2 = klein(nach), v1 = klein(a.vor), v2 = klein(vor);
    if (n1 != n2) return false;
    return v1 == v2 || v1.indexOf(v2 + ' ') == 0 || v2.indexOf(v1 + ' ') == 0;
  }

  var monate = ['januar', 'februar', 'maerz', 'april', 'mai', 'juni', 'juli', 'august', 'september',
                'oktober', 'november', 'dezember'];
  function zwei(n) { return (n < 10 ? '0' : '') + n; }

  // alle Datumsangaben eines Textes als TT.MM.JJJJ: T/M/JJJJ (LibreView-Tabelle, Tag zuerst),
  // T.M.JJJJ und "1 Januar 1975"
  function daten(text) {
    var erg = [], m, re1 = /(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})/g;
    while ((m = re1.exec(text))) erg.push(zwei(+m[1]) + '.' + zwei(+m[2]) + '.' + m[3]);
    var t = klein(text).replace(/ä/g, 'ae'), re2 = /(\d{1,2})\.?\s+([a-z]+)\s+(\d{4})/g;
    while ((m = re2.exec(t))) {
      var i = monate.indexOf(m[2]);
      if (i >= 0) erg.push(zwei(+m[1]) + '.' + zwei(i + 1) + '.' + m[3]);
    }
    return erg;
  }

  var UUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;

  function zumPatienten(a, id) {
    titel(a, 'patient');
    location.assign('/patient/' + id + '/reports');
  }

  // --- Seiten ----------------------------------------------------------------------------------------
  // Wohnsitzland und Sprache (nur beim ersten Besuch bzw. ohne Cookies)
  function land(a) {
    var c = $('country-select'), l = $('language-select'), b = $('submit-button');
    if (!c || !b) return false;
    titel(a, 'land');
    if (einmal(a, 'land')) {
      if (c.value != 'DE') setze(c, 'DE');
      if (l && l.value != 'de-DE') setze(l, 'de-DE');
      setTimeout(function () { b.click(); }, 300);
    }
    return true;
  }

  // Anmeldung: hoechstens ein Versuch je Auftrag
  function anmelden(a) {
    var em = $('loginForm-email-input'), pw = $('loginForm-password-input'), b = $('loginForm-submit-button');
    if (!em || !pw || !b) return false;
    if (seit(a, 'anmelden') >= 0) {
      if (/Problem mit Ihrer E-Mail/i.test(document.body.innerText)) { titel(a, 'anmeldefehler'); erledigt(); }
      else titel(a, 'anmelden');
      return true;
    }
    // Kennung vorgegeben ([LibreView] Kennung=): hat Firefox einen anderen Benutzer eingetragen, ihn samt
    // Passwort ersetzen; die Firefox-Liste (Pfeil runter) zeigt dann nur die passende gespeicherte Anmeldung
    if (a.kennung != '' && klein(em.value) != klein(a.kennung) && einmal(a, 'kennung')) {
      setze(em, a.kennung);
      if (pw.value != '') setze(pw, '');
    }
    if (em.value != '' && pw.value != '') {
      einmal(a, 'anmelden');
      titel(a, 'anmelden');
      b.click();
      return true;
    }
    // Firefox traegt gespeicherte Daten kurz nach dem Laden ein; erst danach nach der Liste fragen
    if (seit(a, 'loginSeite') < 0) einmal(a, 'loginSeite');
    else if (seit(a, 'loginSeite') > 700 && einmal(a, 'passwortFrage')) { em.focus(); }
    titel(a, seit(a, 'passwortFrage') >= 0 ? 'passwort' : 'anmelden');
    return true;
  }

  // Zwei-Faktor, Schritt 1: Code per E-Mail anfordern
  function zweiFaktor(a) {
    var sel = $('2fa-method-select'), b = $('twoFactor-step1-next-button');
    if (!sel || !b) return false;
    titel(a, 'zweifaktor');
    if (einmal(a, 'codeAnfordern')) {
      if (sel.value != 'email') setze(sel, 'email');
      setTimeout(function () { b.click(); }, 500);
    }
    return true;
  }

  // Zwei-Faktor, Schritt 2: 14 Tage merken, PC-Name als Kurzname, Fokus ins Codefeld; den Code tippt der Mensch
  function codeSeite(a) {
    var code = $('twoFactor-step2-code-input'), cb = $('twoFactor-step2-rememberBrowser-checkbox');
    if (!code) return false;
    titel(a, 'code');
    if (cb && !cb.checked && einmal(a, 'merken')) {
      var lab = $('twoFactor-step2-rememberBrowser-checkbox-label');
      (lab || cb).click();
    }
    if (cb && cb.checked && a.pc && !a.schritte.kurzname) {
      // das Feld "Kurzname" erscheint erst nach dem Haken
      var karte = $('wizard-card') || document, felder = karte.querySelectorAll('input');
      for (var i = 0; i < felder.length; i++) {
        var f = felder[i];
        if (f === code || f.type == 'checkbox' || f.type == 'hidden' || !sichtbar(f)) continue;
        if (f.value == '') setze(f, a.pc);
        einmal(a, 'kurzname');
        break;
      }
    }
    if (einmal(a, 'codeFokus')) code.focus();
    return true;
  }

  // Dashboard: Patient in der Tabelle, sonst ueber die Patientensuche
  function dashboard(a) {
    if (!/^\/dashboard/.test(location.pathname)) return false;
    titel(a, 'suche');
    var zeilen = document.querySelectorAll('tr[id$="-table-cell-row"]');
    for (var i = 0; i < zeilen.length; i++) {
      var z = zeilen[i], id = (z.id.match(UUID) || [])[0];
      var zelle = function (c) { var e = z.querySelector('[data-col="' + c + '"]'); return e ? e.textContent : ''; };
      if (id && namePasst(a, zelle('lastName'), zelle('firstName')) && daten(zelle('dob')).indexOf(a.geb) >= 0) {
        zumPatienten(a, id);
        return true;
      }
    }
    // Tabelle erst laden lassen
    if (seit(a, 'dashboard') < 0) { einmal(a, 'dashboard'); return true; }
    if (seit(a, 'dashboard') < 2500 && !zeilen.length) return true;
    var huelle = $('main-header-patient-search-wrapper'), ein = $('main-header-patient-search-inner-input');
    if (!huelle || !ein) return true;
    if (einmal(a, 'suchen')) {
      huelle.click();
      // nur nach dem Nachnamen, das Geburtsdatum entscheidet unter den Treffern
      setTimeout(function () { ein.focus(); setze(ein, a.name); }, 300);
      return true;
    }
    // Treffer der Suche: Elemente mit Patientenkennung, deren Text Name und Geburtsdatum enthaelt
    var ergebnis = huelle.querySelector('[class*="searchResults"]') || huelle;
    var kand = ergebnis.querySelectorAll('[id], a[href]'), gesehen = 0;
    for (var j = 0; j < kand.length; j++) {
      var k = kand[j], kid = ((k.id || '') + ' ' + (k.getAttribute('href') || '')).match(UUID);
      if (!kid) continue;
      gesehen++;
      var txt = k.textContent;
      if (klein(txt).indexOf(klein(a.name)) >= 0 && daten(txt).indexOf(a.geb) >= 0) {
        zumPatienten(a, kid[0]);
        return true;
      }
    }
    if (seit(a, 'suchen') > 8000) {
      titel(a, gesehen ? 'liste' : 'nichts');
      erledigt();
    }
    return true;
  }

  // Patientenseite: zu "Glukoseverlaufsdaten", dort "Glukose-Berichte"
  function patient(a) {
    var m = location.pathname.match(/^\/patient\/([0-9a-f-]{36})(\/[a-z]+)?/i);
    if (!m) return false;
    if (m[2] != '/reports') { zumPatienten(a, m[1]); return true; }
    titel(a, 'patient');
    var b = $('pastGlucoseCard-report-button');
    if (b) {
      titel(a, 'fertig');
      erledigt();
      b.click();
      return true;
    }
    if (seit(a, 'patientSeite') < 0) einmal(a, 'patientSeite');
    else if (seit(a, 'patientSeite') > 8000 && ($('glucose-history-Page') || $('patient-reports-list-view'))) {
      titel(a, 'keinedaten');
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
      land(a) || anmelden(a) || zweiFaktor(a) || codeSeite(a) || patient(a) || dashboard(a) ||
        titel(a, zustand || 'suche');
    } catch (e) {
      titel(a, 'fehler ' + e.message);
      erledigt();
    }
  }

  auftragAusAdresse();
  // gleiche Seite mit neuem "#mor=" laedt nicht neu, sondern aendert nur den Anker
  window.addEventListener('hashchange', function () { if (auftragAusAdresse()) zustand = ''; });
  setInterval(schritt, takt);
})();
