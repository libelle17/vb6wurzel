// ==UserScript==
// @name         MOReiter Glooko
// @namespace    https://github.com/libelle17/vb6wurzel
// @version      1.3
// @description  Glooko (Praxis): Patient nach Name und Geburtsdatum suchen, PDF-Bericht erstellen und herunterladen (Auftrag von MOReiter, Strg+G)
// @match        https://*.glooko.com/*
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @updateURL    https://raw.githubusercontent.com/libelle17/vb6wurzel/main/MOReiter/MOReiter-Glooko.user.js
// @downloadURL  https://raw.githubusercontent.com/libelle17/vb6wurzel/main/MOReiter/MOReiter-Glooko.user.js
// @run-at       document-start
// ==/UserScript==

// MOReiter oeffnet die Patientenliste mit dem Auftrag hinter "#":
//     ...#mor=Nachname;Vorname;TT.MM.JJJJ;Kennung;Nr;PC-Name;Art;Tage   (Teile URL-kodiert)
// Die Anmeldung liegt auf my.glooko.com, die Patienten auf de-fr.my.glooko.com: der sessionStorage gilt nur je
// Herkunft, daher wird der Auftrag zusaetzlich im Speicher von Tampermonkey abgelegt und auf der anderen Seite
// abgeholt (hoechstens 2 min seit dem letzten Schritt). Ablauf: Patientensuche, Zeile mit Name und Geburtsdatum, "PDF-Bericht
// erstellen", "PDF erstellen", das angezeigte PDF wird als VornameNachname_Glooko_TT-MM-JJJJ.pdf heruntergeladen
// (MOReiter verschiebt es mit [Glooko] Ablage=/Datei=). Fortschritt im Seitentitel "MOR:Nr:Zustand":
//   passwort / passworttippen / anmelden / anmeldefehler   Anmeldung (hoechstens ein Versuch je Auftrag)
//   suche, patient, fertig (PDF wird erstellt, MOReiter ist fertig), gedruckt (heruntergeladen), nichts, fehler ...
(function () {
  'use strict';
  var schluessel = 'MOReiterGlooko';
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

  // im Tab (sessionStorage) und fuer den Wechsel zwischen my.glooko.com und de-fr.my.glooko.com bei Tampermonkey
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
    // ohne "view": wegen @grant GM_* laeuft das Skript in Firefox abgeschottet, dort ist "window" nicht das
    // Fenster der Seite, und der MouseEvent-Konstruktor lehnt es ab
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

  // --- Seiten ----------------------------------------------------------------------------------------
  // Cookie-Hinweis (OneTrust): alle ablehnen
  function cookies(a) {
    var b = document.getElementById('onetrust-reject-all-handler') || $('.ot-pc-refuse-all-handler');
    if (sichtbar(b) && einmal(a, 'cookies')) klicke(b);
    return false;
  }

  // Anmeldung wie bei Dexcom: hoechstens ein Versuch je Auftrag; mit Kennung nur diese (Firefox-Liste ueber
  // "passworttippen"), sonst die erste gespeicherte Anmeldung ("passwort")
  var stand = '', standSeit = 0;
  function anmelden(a) {
    var em = $('#user_email'), pw = $('#user_password'), b = $('#sign-in-button');
    if (!em || !pw || !b) return false;
    if (seit(a, 'anmelden') >= 0) {
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

  // Schutz vor ungewolltem Abschicken, solange ein Auftrag laeuft
  function abschickenPruefen(ev) {
    var a = auftrag(), em = $('#user_email'), pw = $('#user_password'), t = ev.target;
    if (!a || !em || !pw || !t) return;
    var login = (t.closest && t.closest('#sign-in-button')) || (ev.type == 'submit' && t.querySelector && t.querySelector('#user_email'));
    var ok = pw.value != '' && (a.kennung == '' || klein(em.value) == klein(a.kennung)) && seit(a, 'anmelden') < 1000;
    if (login && !ok) {
      ev.preventDefault();
      ev.stopImmediatePropagation();
      console.log('MOReiter Glooko: Anmelden ohne passende Kennung/Passwort oder zum zweiten Mal verhindert');
    }
  }
  window.addEventListener('submit', abschickenPruefen, true);
  window.addEventListener('click', abschickenPruefen, true);

  // das angezeigte PDF (iframe) unter festem Namen herunterladen
  function pdfName(a) {
    var d = new Date(), z = function (n) { return (n < 10 ? '0' : '') + n; };
    return (a.vor + a.name).replace(/[\s\\\/:*?"<>|]+/g, '') + '_Glooko_' + z(d.getDate()) + '-' + z(d.getMonth() + 1) + '-' +
      d.getFullYear() + '.pdf';
  }
  function herunterladen(a, src) {
    var l = document.createElement('a');
    l.href = src;
    l.download = pdfName(a);
    l.style.display = 'none';
    document.body.appendChild(l);
    l.click();
    setTimeout(function () { l.remove(); }, 5000);
  }

  function freigeben() {
    setTimeout(function () { if (/^MOR:\d+:(gedruckt|fertig)/.test(document.title)) document.title = 'LogbookWeb'; }, 60000);
  }

  // PDF fertig: Rahmen mit dem PDF
  function pdfRahmen(a) {
    var f = $('iframe.PdfLoader_pdfContent, iframe[title="iframe-pdfdownload"]');
    if (!f || !f.src || seit(a, 'pdfErstellt') < 0) return false;
    titel(a, 'fertig');
    if (einmal(a, 'heruntergeladen')) {
      herunterladen(a, f.src);
      titel(a, 'gedruckt');
      erledigt();
      freigeben();
    }
    return true;
  }

  // Fenster "PDF erstellen": Knopf druecken, sobald aktiv; steht das Fenster 8 s spaeter noch ohne PDF da
  // (Klick zu frueh, bevor Glooko bereit war), noch hoechstens zweimal
  function pdfFenster(a) {
    var b = $('[data-testid="button-pdf-create-pdf"]');
    if (!sichtbar(b)) return false;
    titel(a, 'patient');
    if (aktiv(b) && einmal(a, 'pdfErstellt')) { klicke(b); titel(a, 'fertig'); return true; }
    if (seit(a, 'pdfErstellt') < 0) return true;
    var n = a.schritte.pdfKlicks || 1;
    if (aktiv(b) && n < 3 && seit(a, 'pdfErstellt') > 8000 * n) {
      a.schritte.pdfKlicks = n + 1;
      merken(a);
      klicke(b);
    }
    titel(a, 'fertig');
    return true;
  }

  // Patientenseite (summary?patient=...): "PDF-Bericht erstellen"
  function patientSeite(a) {
    var b = $('#create-pdf');
    if (!sichtbar(b)) return false;
    if (seit(a, 'pdfErstellt') >= 0) { titel(a, 'fertig'); return true; }
    titel(a, 'patient');
    if (einmal(a, 'berichtKlick')) klicke(b);
    else if (seit(a, 'berichtKlick') > 5000 && einmal(a, 'berichtKlick2')) klicke(b);
    return true;
  }

  // Name in der Liste "Vorname Nachname": endet auf den Nachnamen, davor der Vorname (s. namePasst)
  function listenNamePasst(a, text) {
    var t = teile(text.replace(/ /g, ' ')), n = teile(a.name);
    if (t.length <= n.length || t.slice(t.length - n.length).join(' ') != n.join(' ')) return false;
    return namePasst(a, a.name, t.slice(0, t.length - n.length).join(' '));
  }

  // juengste Uebertragung eines Geraets in der Zeile als JJJJMMTT, aus aria-label="Geraet, 02. Okt. 2025"; 0 = keine
  function letzteUebertragung(tr) {
    var st = 0, l = tr ? tr.querySelectorAll('[aria-label]') : [];
    for (var i = 0; i < l.length; i++) {
      daten(l[i].getAttribute('aria-label')).forEach(function (d) {
        var t = d.split('.'), n = +(t[2] + t[1] + t[0]);
        if (n > st) st = n;
      });
    }
    return st;
  }

  // Patientenliste: suchen, passende Zeile anklicken
  function suchTexte(a) {
    var l = [a.name, a.vor + ' ' + a.name], v = teile(a.vor);
    if (v.length > 1) v.forEach(function (x) { l.push(x.charAt(0).toUpperCase() + x.slice(1) + ' ' + a.name); });
    return l;
  }
  function liste(a) {
    var feld = $('input[type="search"][data-testid="textInput-input"]') || $('input[type="search"]');
    var namen = document.querySelectorAll('.PatientColumn_nameWrapper');
    if (!feld && !namen.length) return false;
    if (seit(a, 'zeileKlick') < 0) titel(a, 'suche');
    // derselbe Patient steht oefter doppelt in der Liste: dann die Zeile mit der juengsten Uebertragung nehmen
    var bester = null, besterStand = -1;
    for (var i = 0; i < namen.length; i++) {
      var knopf = namen[i], spalte = knopf.closest('.PatientColumn_columnWrapper') || knopf.parentElement;
      var geb = spalte && spalte.querySelector('.PatientColumn_leftColumnBottom');
      if (!geb || !listenNamePasst(a, knopf.textContent) || daten(geb.textContent).indexOf(a.geb) < 0) continue;
      var st = letzteUebertragung(knopf.closest('tr'));
      if (st > besterStand) { bester = knopf; besterStand = st; }
    }
    if (bester) {
      titel(a, 'patient');
      if (einmal(a, 'zeileKlick')) klicke(bester);
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
      // Glooko filtert die grosse Liste erst nach Enter (Abschicken des Suchformulars); die Vorschlagsliste
      // unter dem Feld wird nicht gebraucht
      setTimeout(function () {
        var f = feld.form;
        if (f && f.requestSubmit) f.requestSubmit();
        else {
          ['keydown', 'keypress', 'keyup'].forEach(function (t) {
            feld.dispatchEvent(new KeyboardEvent(t, { bubbles: true, cancelable: true, key: 'Enter', code: 'Enter', keyCode: 13, which: 13 }));
          });
        }
      }, 700);
      return true;
    }
    // nichts Passendes: MOReiter meldet es; der Auftrag bleibt aber noch 3 min, damit es nach einer Auswahl von
    // Hand (z.B. anderer Schreibweise) mit dem PDF weitergeht
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
      anmelden(a) || pdfRahmen(a) || pdfFenster(a) || patientSeite(a) || liste(a) || titel(a, zustand || 'suche');
    } catch (e) {
      titel(a, 'fehler ' + e.message);
      erledigt();
    }
  }

  auftragAusAdresse();
  window.addEventListener('hashchange', function () { if (auftragAusAdresse()) zustand = ''; });
  setInterval(schritt, takt);
})();
