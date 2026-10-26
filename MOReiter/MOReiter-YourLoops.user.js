// ==UserScript==
// @name         MOReiter YourLoops
// @namespace    https://github.com/libelle17/vb6wurzel
// @version      1.1
// @description  YourLoops (Diabeloop, Praxis): Patient nach Name und Geburtsdatum suchen, PDF-Bericht herunterladen (Auftrag von MOReiter, Strg+Y)
// @match        https://www.your-loops.com/*
// @match        https://auth.diabeloop.com/*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @updateURL    https://raw.githubusercontent.com/libelle17/MOReiter/main/MOReiter-YourLoops.user.js
// @downloadURL  https://raw.githubusercontent.com/libelle17/MOReiter/main/MOReiter-YourLoops.user.js
// @run-at       document-start
// ==/UserScript==

// MOReiter oeffnet YourLoops mit dem Auftrag hinter "#":
//     ...#mor=Nachname;Vorname;TT.MM.JJJJ;Kennung;Nr;PC-Name;Art;Tage   (Teile URL-kodiert)
// Die Anmeldung liegt auf auth.diabeloop.com, die Patienten auf www.your-loops.com: der sessionStorage gilt nur
// je Herkunft, daher wird der Auftrag (wie bei Glooko) zusaetzlich im Speicher von Tampermonkey abgelegt und auf
// der anderen Seite abgeholt (hoechstens 2 min seit dem letzten Schritt). Ablauf: Cookie-Hinweis "Nein danke",
// "Verbinden", Anmeldung, ggf. Code aus der Mail (den tippt der Mensch; "Dieses Geraet 30 Tage lang speichern"
// hakt das Skript an), Patientensuche, Zeile mit Name und Geburtsdatum, "Download-Bericht", Zeitraum Tage (7, 14,
// 28 oder 90), PDF, "Herunterladen". Fortschritt im Seitentitel "MOR:Nr:Zustand":
//   passwort / passworttippen / anmelden / anmeldefehler   Anmeldung (hoechstens ein Versuch je Auftrag)
//   code        Code-Seite: der Mensch tippt den Code aus der Mail
//   suche, patient, fertig (Bericht wird heruntergeladen), gedruckt (Fenster zu), nichts, fehler ...
(function () {
  'use strict';
  var schluessel = 'MOReiterYourLoops';
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
    history.replaceState(history.state, '', location.pathname + location.search);
    return true;
  }

  // im Tab (sessionStorage) und fuer den Wechsel zwischen www.your-loops.com und auth.diabeloop.com bei Tampermonkey
  function auftrag() {
    var a = null;
    try { a = JSON.parse(sessionStorage.getItem(schluessel) || 'null'); } catch (e) {}
    if (a) return a;
    try { a = JSON.parse(GM_getValue(schluessel, 'null')); } catch (e) {}
    // nur einen frischen (eben noch bearbeiteten) Auftrag uebernehmen, nicht in einem anderen, spaeter
    // von Hand geoeffneten Tab
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

  // Wert so setzen, dass das Rahmenwerk (React) ihn mitbekommt
  function setze(el, v) {
    var proto = el.tagName == 'SELECT' ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, v);
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'a' }));
  }

  // Klick wie mit der Maus; ohne "view": wegen @grant GM_* laeuft das Skript in Firefox abgeschottet, dort ist
  // "window" nicht das Fenster der Seite, und der MouseEvent-Konstruktor lehnt es ab
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
  function aktiv(b) { return b && !b.disabled && !b.hasAttribute('disabled'); }

  // Knopf/Link mit genau diesem Text (sichtbar), auch im Schatten-DOM des Cookie-Hinweises (Axeptio)
  function mitText(sel, text, wurzel) {
    var l = (wurzel || document).querySelectorAll(sel);
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

  // Name in der Liste "NACHNAME VORNAME" (z.B. "AL MULLA HARUN"), vorsichtshalber auch "Vorname Nachname"
  function listenNamePasst(a, text) {
    var t = teile(text), n = teile(a.name);
    if (t.length <= n.length) return false;
    if (t.slice(0, n.length).join(' ') == n.join(' ') && namePasst(a, a.name, t.slice(n.length).join(' '))) return true;
    return t.slice(t.length - n.length).join(' ') == n.join(' ') && namePasst(a, a.name, t.slice(0, t.length - n.length).join(' '));
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

  // alle Datumsangaben eines Textes als TT.MM.JJJJ: T.M.JJJJ, T/M/JJJJ, "31. Juli 2003", "18. Okt. 2026"
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
  // Cookie-Hinweis (Axeptio, im Schatten-DOM): "Nein danke"; stoert sonst nicht, daher kein eigener Zustand
  function cookies(a) {
    if (a.schritte.cookies) return false;
    var wurzeln = [document], l = document.querySelectorAll('[id^="axeptio"], .axeptio_mount');
    for (var i = 0; i < l.length; i++) if (l[i].shadowRoot) wurzeln.push(l[i].shadowRoot);
    for (var j = 0; j < wurzeln.length; j++) {
      var b = mitText('button', 'Nein danke', wurzeln[j]);
      if (b) { einmal(a, 'cookies'); klicke(b); break; }
    }
    return false;
  }

  // Startseite www.your-loops.com/login: "Verbinden" fuehrt zur Anmeldung bei Diabeloop
  function startSeite(a) {
    var b = mitText('button, a', 'Verbinden');
    if (!b || location.hostname != 'www.your-loops.com') return false;
    titel(a, 'anmelden');
    if (einmal(a, 'verbinden')) klicke(b);
    else if (seit(a, 'verbinden') > 8000 && einmal(a, 'verbinden2')) klicke(b);
    return true;
  }

  // Anmeldung (auth.diabeloop.com) wie bei Dexcom: hoechstens ein Versuch je Auftrag; mit Kennung nur diese
  // (Firefox-Liste ueber "passworttippen"), sonst die erste gespeicherte Anmeldung ("passwort")
  var stand = '', standSeit = 0;
  function loginKnopf(em) {
    var f = em && em.form;
    return f && f.querySelector('button[type="submit"][name="action"][value="default"], button[type="submit"]');
  }
  function anmelden(a) {
    var em = $('#username'), pw = $('#password'), b = loginKnopf(em);
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
    var b = loginKnopf(em);
    var login = (b && t.closest && t.closest('button') === b) || (ev.type == 'submit' && t === em.form);
    var ok = pw.value != '' && (a.kennung == '' || klein(em.value) == klein(a.kennung)) && seit(a, 'anmelden') < 1000;
    if (login && !ok) {
      ev.preventDefault();
      ev.stopImmediatePropagation();
      console.log('MOReiter YourLoops: Anmelden ohne passende Kennung/Passwort oder zum zweiten Mal verhindert');
    }
  }
  window.addEventListener('submit', abschickenPruefen, true);
  window.addEventListener('click', abschickenPruefen, true);

  // Code aus der Mail: "Dieses Geraet 30 Tage lang speichern" anhaken, Fokus ins Codefeld; den Code tippt der
  // Mensch und drueckt "Fortfahren"
  function codeSeite(a) {
    var code = $('#code'), cb = $('#rememberBrowser');
    if (!code || location.hostname != 'auth.diabeloop.com') return false;
    titel(a, 'code');
    if (cb && !cb.checked && einmal(a, 'merken')) cb.click();
    if (einmal(a, 'codeFokus')) code.focus();
    return true;
  }

  // Zeitraum des Berichts an die Fertigmeldung haengen ("fertig|TT.MM.JJJJ-TT.MM.JJJJ"), MOReiter setzt ihn in
  // den Dateinamen ein
  function mitZeitraum(a, z) { return a.zeitraum ? z + '|' + a.zeitraum : z; }

  // nach dem Download ist der Tab nach 60 s wieder frei fuer MOReiter
  function freigeben() {
    setTimeout(function () { if (/^MOR:\d+:(gedruckt|fertig)/.test(document.title)) document.title = 'YourLoops'; }, 60000);
  }

  // Fenster "Download-Bericht": fester Zeitraum nach Tage, PDF, "Herunterladen". Erst "fertig" melden und
  // MOReiter Zeit geben, es im Titel zu lesen: Firefox oeffnet das PDF in einem neuen Tab, dann ist dieser
  // Titel nicht mehr zu sehen (MOReiter hoert bei "fertig" auf zu lesen). Steht das Fenster 6 s nach dem Klick
  // noch da, ein zweiter Klick
  var zeitraeume = { '7': 'one-week', '14': 'two-weeks', '28': 'four-weeks', '30': 'four-weeks', '90': 'three-months' };
  function berichtFenster(a) {
    var b = $('[data-testid="pdf-options-button-download"]');
    if (!sichtbar(b)) return false;
    if (seit(a, 'herunterladen') < 0) titel(a, 'patient');
    var chip = zeitraeume[a.tage] && document.getElementById('pdf-options-button-' + zeitraeume[a.tage]);
    if (chip && chip.getAttribute('aria-selected') != 'true') {
      if (einmal(a, 'zeitraum')) klicke(chip);
      else if (seit(a, 'zeitraum') > 3000 && einmal(a, 'zeitraum2')) klicke(chip);
      return true;
    }
    var pdf = document.getElementById('dialog-pdf-options-selector-pdf');
    if (pdf && !pdf.checked) { if (einmal(a, 'pdfWahl')) pdf.click(); return true; }
    // eingestellter Zeitraum, z.B. "21. Juli 2026 → 18. Okt. 2026"
    var zr = document.getElementById('pdf-options-button-custom-range'), d = zr ? daten(zr.value) : [];
    if (d.length >= 2 && a.zeitraum != d[0] + '-' + d[1]) { a.zeitraum = d[0] + '-' + d[1]; merken(a); }
    if (!aktiv(b)) return true;
    if (seit(a, 'fertigGemeldet') < 0) { einmal(a, 'fertigGemeldet'); titel(a, mitZeitraum(a, 'fertig')); return true; }
    if (seit(a, 'fertigGemeldet') < 1500) return true;
    if (einmal(a, 'herunterladen')) { klicke(b); titel(a, mitZeitraum(a, 'fertig')); return true; }
    if (seit(a, 'herunterladen') > 6000 && einmal(a, 'herunterladen2')) klicke(b);
    return true;
  }

  // Patientenseite (Uebersicht): "Download-Bericht"; ist das Fenster danach zu, laeuft der Download
  function patientSeite(a) {
    var b = $('[data-testid="download-report"]');
    if (!sichtbar(b)) return false;
    if (seit(a, 'herunterladen') >= 0) {
      titel(a, mitZeitraum(a, 'gedruckt'));
      erledigt();
      freigeben();
      return true;
    }
    titel(a, 'patient');
    if (einmal(a, 'berichtKlick')) klicke(b);
    else if (seit(a, 'berichtKlick') > 5000 && einmal(a, 'berichtKlick2')) klicke(b);
    return true;
  }

  // Patientenliste: suchen, passende Zeile (Name, Geburtsdatum) anklicken; derselbe Patient doppelt eingetragen:
  // die Zeile mit dem juengsten "Letztes Daten-Update"
  function suchTexte(a) {
    var l = [a.name, a.name + ' ' + a.vor], v = teile(a.vor);
    if (v.length > 1) v.forEach(function (x) { l.push(a.name + ' ' + x); });
    return l;
  }
  function liste(a) {
    var feld = $('[data-testid="search-patient-bar"]');
    var zeilen = document.querySelectorAll('[data-testid="current-patient-list-grid"] [role="row"][data-id]');
    if (!feld && !zeilen.length) return false;
    if (seit(a, 'zeileKlick') < 0) titel(a, 'suche');
    var bester = null, besterStand = -1;
    for (var i = 0; i < zeilen.length; i++) {
      var z = zeilen[i], n = z.querySelector('[data-field="patient"]'), g = z.querySelector('[data-field="date-of-birth"]');
      if (!n || !g || !listenNamePasst(a, n.textContent) || daten(g.textContent).indexOf(a.geb) < 0) continue;
      var up = z.querySelector('[data-field="last-data-update"]'), st = up ? datumZahl(daten(up.textContent)[0]) : 0;
      if (st > besterStand) { bester = n; besterStand = st; }
    }
    if (bester) {
      titel(a, 'patient');
      if (einmal(a, 'zeileKlick')) klicke(bester);
      else if (seit(a, 'zeileKlick') > 5000 && einmal(a, 'zeileKlick2')) klicke(bester.closest('[role="row"]') || bester);
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
    // nichts Passendes: MOReiter meldet es; der Auftrag bleibt aber noch 3 min, damit es nach einer Auswahl von
    // Hand (z.B. anderer Schreibweise) mit dem Bericht weitergeht
    if (k >= suchTexte(a).length && seit(a, 'suche' + (suchTexte(a).length - 1)) > 5000 && seit(a, 'zeileKlick') < 0) {
      einmal(a, 'nichts');
      titel(a, 'nichts');
      if (seit(a, 'nichts') > 180000) erledigt();
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
      anmelden(a) || codeSeite(a) || startSeite(a) || berichtFenster(a) || patientSeite(a) || liste(a) ||
        titel(a, zustand || 'suche');
    } catch (e) {
      titel(a, 'fehler ' + e.message);
      erledigt();
    }
  }

  auftragAusAdresse();
  window.addEventListener('hashchange', function () { if (auftragAusAdresse()) zustand = ''; });
  setInterval(schritt, takt);
})();
