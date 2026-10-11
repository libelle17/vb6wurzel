// ==UserScript==
// @name         MOReiter LibreView
// @namespace    https://github.com/libelle17/vb6wurzel
// @version      1.7
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
//   passworttippen  Anmeldeseite: im E-Mail-Feld (Fokus) steht die Kennung ohne ihr letztes Zeichen; MOReiter
//               tippt es nach und waehlt die gespeicherte Anmeldung aus der Firefox-Liste
//   anmelden    Anmeldung abgeschickt (hoechstens einmal je Auftrag, sonst droht eine Sperre)
//   anmeldefehler  LibreView hat E-Mail/Passwort abgelehnt
//   ohnekennung keine Kennung vorgegeben: nicht automatisch anmelden
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

  // Klick wie mit der Maus: Element in den sichtbaren Bereich holen, dann pointer-/mousedown, -up und click an
  // seiner Mitte (manche Listen reagieren nur auf mousedown, element.click() allein liefe dann ins Leere)
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

  // Vorname: "Hans Peter" in MO passt zu "Hans" in LibreView und umgekehrt
  function namePasst(a, nach, vor) {
    var n1 = klein(a.name), n2 = klein(nach), v1 = klein(a.vor), v2 = klein(vor);
    if (n1 != n2) return false;
    return v1 == v2 || v1.indexOf(v2 + ' ') == 0 || v2.indexOf(v1 + ' ') == 0;
  }

  var monate = ['januar', 'februar', 'maerz', 'april', 'mai', 'juni', 'juli', 'august', 'september',
                'oktober', 'november', 'dezember'];
  // Monat aus dem vollen oder abgekuerzten Namen (deutsch, sonst englisch), 0-11, -1 unbekannt
  var monateEn = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  function monat(w) {
    for (var i = 0; i < 12; i++) if (monate[i].indexOf(w) == 0 || w.indexOf(monate[i].slice(0, 3)) == 0 ||
        w.indexOf(monateEn[i]) == 0) return i;
    return -1;
  }
  function zwei(n) { return (n < 10 ? '0' : '') + n; }

  // alle Datumsangaben eines Textes als TT.MM.JJJJ: T/M/JJJJ (LibreView-Tabelle, Tag zuerst),
  // T.M.JJJJ und "1 Januar 1975"
  function daten(text) {
    var erg = [], m, re1 = /(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})/g;
    while ((m = re1.exec(text))) erg.push(zwei(+m[1]) + '.' + zwei(+m[2]) + '.' + m[3]);
    // "1 Januar 1975", "1. Jan. 1975", "Jan 1, 1975"
    var t = klein(text).replace(/ä/g, 'ae'), re2 = /(\d{1,2})\.?\s+([a-z]{3,})\.?,?\s+(\d{4})/g,
        re3 = /\b([a-z]{3,})\.?\s+(\d{1,2}),?\s+(\d{4})/g;
    while ((m = re2.exec(t))) { var i = monat(m[2]); if (i >= 0) erg.push(zwei(+m[1]) + '.' + zwei(i + 1) + '.' + m[3]); }
    while ((m = re3.exec(t))) { var k = monat(m[1]); if (k >= 0) erg.push(zwei(+m[2]) + '.' + zwei(k + 1) + '.' + m[3]); }
    return erg;
  }

  var UUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;

  // zum Patienten wie ein Mensch: auf den Treffer bzw. die Tabellenzeile $el klicken (LibreView leitet einen
  // direkten Aufruf von /patient/<id>/... sonst zum Dashboard zurueck); erst wenn sich nach 5 s nichts tut,
  // die Adresse laden
  function zumPatienten(a, id, el) {
    titel(a, 'patient');
    if (el && einmal(a, 'patientKlick')) {
      console.log('MOReiter LibreView: Klick auf <' + el.tagName.toLowerCase() + ' id="' +
        (el.id || '').replace(new RegExp(UUID.source, 'gi'), '<uuid>') + '" class="' + el.className + '">');
      klicke(el);
      return;
    }
    if ((!el || seit(a, 'patientKlick') > 5000) && einmal(a, 'patientAdresse')) {
      console.log('MOReiter LibreView: Klick ohne Wirkung, lade /patient/<uuid>/profile');
      location.assign('/patient/' + id + '/profile');
    }
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

  // Anmeldung: hoechstens ein Versuch je Auftrag, und nur mit der vorgegebenen Kennung ([LibreView] Kennung=).
  // Firefox traegt beim Laden (auch mehrmals) die erste gespeicherte Anmeldung ein, die eines anderen Benutzers
  // sein kann: dann E-Mail ersetzen und Passwort leeren, bei jedem Takt. Abgeschickt wird erst, wenn E-Mail
  // gleich der Kennung ist, ein Passwort drinsteht und sich beides 1 s nicht mehr geaendert hat.
  function anmelden(a) {
    var em = $('loginForm-email-input'), pw = $('loginForm-password-input'), b = $('loginForm-submit-button');
    if (!em || !pw || !b) return false;
    if (seit(a, 'anmelden') >= 0) {
      if (/Problem mit Ihrer E-Mail/i.test(document.body.innerText)) { titel(a, 'anmeldefehler'); erledigt(); }
      else titel(a, 'anmelden');
      return true;
    }
    if (a.kennung == '') { titel(a, 'ohnekennung'); erledigt(); return true; }
    if (seit(a, 'loginSeite') < 0) einmal(a, 'loginSeite');
    // Firefox erst eintragen lassen
    if (seit(a, 'loginSeite') < 1500) { titel(a, 'anmelden'); return true; }
    // Firefox oeffnet seine Liste der gespeicherten Anmeldungen nur bei echtem Tippen: daher die Kennung ohne
    // ihr letztes Zeichen einsetzen; MOReiter tippt es auf "passworttippen" nach und waehlt dann den ersten
    // (auf diese Kennung gefilterten) Eintrag der Liste
    var ohneLetztes = a.kennung.slice(0, -1);
    if (klein(em.value) != klein(a.kennung)) {
      if (klein(em.value) == klein(ohneLetztes) && pw.value == '') {
        // Fokus bei jedem Takt zurueckholen, sonst landet das nachgetippte Zeichen woanders
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
    var jetzt = em.value + '\u0001' + pw.value;
    if (jetzt != stand) { stand = jetzt; standSeit = Date.now(); }
    // Kennung vollstaendig, aber nach 3 s noch kein Passwort (Firefox-Liste nicht aufgegangen): wieder ohne
    // letztes Zeichen, MOReiter tippt erneut (hoechstens dreimal)
    if (pw.value == '' && Date.now() - standSeit > 3000) {
      setze(em, ohneLetztes);
      stand = '';
      titel(a, 'passworttippen');
      return true;
    }
    if (pw.value != '' && Date.now() - standSeit >= 1000 && klein(em.value) == klein(a.kennung)) {
      einmal(a, 'anmelden');
      titel(a, 'anmelden');
      b.click();
      return true;
    }
    // Kennung vollstaendig, aber (noch) kein Passwort: Firefox traegt es nach der Auswahl ein; kommt keins,
    // bleibt es dabei (MOReiter meldet dann "bitte selbst anmelden")
    titel(a, 'anmelden');
    return true;
  }
  var stand = '', standSeit = 0;

  // Schutz vor ungewolltem Abschicken (z.B. Enter von MOReiter, wenn die Firefox-Liste nicht aufging), solange
  // ein Auftrag laeuft: nur mit E-Mail = Kennung und Passwort; zaehlt sonst womoeglich als Fehlversuch
  function abschickenErlaubt() {
    var a = auftrag(), em = $('loginForm-email-input'), pw = $('loginForm-password-input');
    if (!a || !em || !pw) return true;
    return a.kennung != '' && klein(em.value) == klein(a.kennung) && pw.value != '';
  }
  function abschickenPruefen(ev) {
    var t = ev.target;
    var login = t && (t.id == 'loginForm-submit-button' || (t.closest && t.closest('#loginForm-submit-button')) ||
      (ev.type == 'submit' && t.querySelector && t.querySelector('#loginForm-email-input')));
    if (login && !abschickenErlaubt()) {
      ev.preventDefault();
      ev.stopImmediatePropagation();
      console.log('MOReiter LibreView: Anmelden ohne passende E-Mail/Passwort verhindert');
    }
  }
  window.addEventListener('submit', abschickenPruefen, true);
  window.addEventListener('click', abschickenPruefen, true);

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
        zumPatienten(a, id, z.querySelector('[data-col="lastName"] div') || z);
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
    // Treffer der Suche: das kleinste Element, dessen Text Nachname und Geburtsdatum enthaelt; traegt es
    // (oder ein umgebendes) eine Patientenkennung, direkt dorthin, sonst darauf klicken
    var ergebnis = document.querySelector('[class*="searchResults"]') || huelle;
    var alle = ergebnis.querySelectorAll('*'), best = null, gesehen = 0;
    for (var j = 0; j < alle.length; j++) {
      var txt = alle[j].textContent;
      if (klein(txt).indexOf(klein(a.name)) < 0) continue;
      gesehen++;
      if (daten(txt).indexOf(a.geb) < 0) continue;
      if (!best || txt.length < best.textContent.length) best = alle[j];
    }
    if (best) {
      for (var e = best; e && e !== ergebnis.parentNode; e = e.parentElement) {
        var kid = ((e.id || '') + ' ' + (e.getAttribute('href') || '')).match(UUID);
        // Treffer: <div id="patient-result-button-<id>" role="button"> mit Name und Geburtsdatum
        if (kid) { zumPatienten(a, kid[0], e); return true; }
      }
      var ziel = best.closest('a, button, li, [role="button"], [role="option"], [tabindex]') || best;
      titel(a, 'patient');
      if (einmal(a, 'trefferKlick')) ziel.click();
      return true;
    }
    if (seit(a, 'suchen') > 8000) {
      bericht(a, zeilen, ergebnis);
      titel(a, gesehen ? 'liste' : 'nichts');
      erledigt();
    }
    return true;
  }

  // fuer die Fehlersuche (Strg+Umsch+K in Firefox): Aufbau von Tabelle und Trefferliste, Namen und Ziffern
  // unkenntlich (Buchstaben -> a, Ziffern -> 9)
  function bericht(a, zeilen, ergebnis) {
    var unk = function (s) { return s.replace(/[^\W\d_]/g, 'a').replace(/\d/g, '9'); };
    var gleich = 0;
    for (var i = 0; i < zeilen.length; i++) {
      var n = zeilen[i].querySelector('[data-col="lastName"]');
      if (n && klein(n.textContent) == klein(a.name)) gleich++;
    }
    console.log('MOReiter LibreView: Tabelle ' + zeilen.length + ' Zeilen, davon ' + gleich +
      ' mit gleichem Nachnamen; gesuchtes Datum ' + unk(a.geb));
    var kopie = ergebnis.cloneNode(true), w = document.createTreeWalker(kopie, NodeFilter.SHOW_TEXT), k;
    while ((k = w.nextNode())) k.nodeValue = unk(k.nodeValue);
    console.log('MOReiter LibreView: Trefferliste\n' + kopie.outerHTML.replace(new RegExp(UUID.source, 'gi'), '<uuid>').slice(0, 20000));
  }

  // Patientenseite: zu "Glukoseverlaufsdaten", dort "Glukose-Berichte"
  function patient(a) {
    var m = location.pathname.match(/^\/patient\/([0-9a-f-]{36})(\/[a-z]+)?/i);
    if (!m) return false;
    titel(a, 'patient');
    if (m[2] != '/reports') {
      // Seitenleiste "Glukoseverlaufsdaten"
      var nav = $('reports-nav-button-container');
      if (nav && einmal(a, 'verlaufKlick')) klicke(nav);
      else if (seit(a, 'verlaufKlick') > 5000 && einmal(a, 'verlaufAdresse')) location.assign('/patient/' + m[1] + '/reports');
      return true;
    }
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
