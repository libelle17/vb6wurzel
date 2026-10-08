// ==UserScript==
// @name         MOReiter Labor
// @namespace    https://github.com/libelle17/vb6wurzel
// @version      1.0
// @description  Onlinebefunde Labor Staber: Patient nach Name und Geburtsdatum suchen und den ersten Befund zeigen (Auftrag von MOReiter, Strg+Alt+B)
// @match        https://onlinebefunde.labor-staber.de/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

// MOReiter oeffnet eine Portalseite mit dem Auftrag hinter "#":
//     ...#mor=Nachname;Vorname;TT.MM.JJJJ;Kennung;Nr   (Teile URL-kodiert)
// Der Auftrag wird im sessionStorage des Tabs gemerkt, damit er die Anmeldung uebersteht, und
// abgearbeitet; den Fortschritt meldet das Skript MOReiter ueber den Seitentitel "MOR:Nr:Zustand":
//   passwort   Anmeldeseite ohne Passwort: Kennung eingetragen, Feld hat den Fokus (MOReiter waehlt
//              dann die gespeicherte Anmeldung aus der Firefox-Liste, das Skript wartet darauf)
//   anmelden   Anmeldeformular abgeschickt
//   suche      angemeldet, Suche laeuft
//   fertig     erster Befund des passenden Patienten wird geoeffnet
//   liste      kein Patient mit passendem Geburtsdatum: Suchliste wird gezeigt
//   nichts     Suche ohne Treffer
//   fehler ... sonstiger Fehler
(function () {
  'use strict';
  var basis = '/onlinebefunde/index.php';
  var schluessel = 'MOReiterAuftrag';

  function teil(s) {
    try { return decodeURIComponent(s.replace(/\+/g, ' ')).trim(); } catch (e) { return s.trim(); }
  }

  // neuer Auftrag aus der Adresse uebernehmen und aus ihr entfernen
  function auftragAusAdresse() {
    var m = location.hash.match(/^#mor=(.*)$/);
    if (!m) return;
    sessionStorage.setItem(schluessel, m[1]);
    history.replaceState(null, '', location.pathname + location.search);
  }

  function auftrag() {
    var a = sessionStorage.getItem(schluessel);
    if (!a) return null;
    var t = a.split(';').map(teil);
    return { name: t[0] || '', vor: t[1] || '', geb: t[2] || '', kennung: t[3] || '', nr: t[4] || '' };
  }

  function erledigt() { sessionStorage.removeItem(schluessel); }

  var auf = null;
  function titel(s) { document.title = 'MOR:' + auf.nr + ':' + s; }

  function lies(r) {
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r.text().then(function (h) { return new DOMParser().parseFromString(h, 'text/html'); });
  }

  // Ziel aus onclick="window.location.href='...'" einer Tabellenzeile
  function ziel(tr) {
    var m = (tr.getAttribute('onclick') || '').match(/href\s*=\s*["']([^"']+)["']/);
    return m ? m[1] : '';
  }

  // Zeilen der Ergebnistabelle mit Patientenseite und Geburtsdatum (Spalte "Geb.Dat.")
  function treffer(doc) {
    var tab = doc.querySelector('table.obtable');
    if (!tab) return [];
    var sp = -1, th = tab.querySelectorAll('thead th');
    for (var i = 0; i < th.length; i++) if (/Geb/.test(th[i].textContent)) sp = i;
    var erg = [], tr = tab.querySelectorAll('tbody tr[onclick]');
    for (var j = 0; j < tr.length; j++) {
      var z = ziel(tr[j]);
      if (z.indexOf('pid=') < 0) continue;
      var c = tr[j].cells;
      erg.push({ ziel: z, geb: (sp >= 0 && c[sp]) ? c[sp].textContent.trim() : '' });
    }
    return erg;
  }

  // erste Befundzeile auf der Seite eines Patienten
  function befund(doc) {
    var tr = doc.querySelectorAll('table.obtable tbody tr[onclick]');
    for (var j = 0; j < tr.length; j++) {
      var z = ziel(tr[j]);
      if (/[?&]id=[1-9]/.test(z) && z.indexOf('offset=') >= 0) return z;
    }
    return '';
  }

  function gehe(z) { location.href = '/onlinebefunde/' + z; }

  // Suchliste normal anzeigen (POST wie das Suchformular)
  function liste() {
    titel('liste');
    var f = document.createElement('form');
    f.method = 'POST';
    f.action = basis + '?func=patienten';
    [['freitext', auf.name + ' ' + auf.vor], ['suchstart', '']].forEach(function (p) {
      var e = document.createElement('input');
      e.type = 'hidden'; e.name = p[0]; e.value = p[1];
      f.appendChild(e);
    });
    document.body.appendChild(f);
    f.submit();
  }

  function suchen() {
    titel('suche');
    var daten = new URLSearchParams();
    daten.append('freitext', auf.name + ' ' + auf.vor);
    daten.append('suchstart', '');
    fetch(basis + '?func=patienten', { method: 'POST', body: daten, credentials: 'same-origin' })
      .then(lies)
      .then(function (doc) {
        var tr = treffer(doc);
        if (tr.length == 0) { titel('nichts'); return; }
        var passend = tr.filter(function (x) { return x.geb == auf.geb; });
        if (passend.length == 0) { liste(); return; }
        // der Reihe nach, bis ein Patient einen Befund hat; sonst den ersten passenden zeigen
        var i = 0;
        function naechster() {
          if (i >= passend.length) { titel('fertig'); gehe(passend[0].ziel); return; }
          var p = passend[i++];
          fetch('/onlinebefunde/' + p.ziel, { credentials: 'same-origin' }).then(lies).then(function (d) {
            var z = befund(d);
            if (z == '') { naechster(); return; }
            titel('fertig');
            gehe(z);
          }).catch(function (e) { titel('fehler ' + e.message); });
        }
        naechster();
      })
      .catch(function (e) { titel('fehler ' + e.message); });
  }

  // Anmeldeseite: abschicken, sobald Firefox (bzw. die Auswahl aus seiner Liste) das Passwort
  // eingetragen hat; hoechstens 20 s warten
  function anmelden(lf) {
    if (lf.login.value == '' && auf.kennung != '') lf.login.value = auf.kennung;
    var beginn = Date.now(), gemeldet = false;
    (function pruefen() {
      if (lf.pwd.value != '' && lf.login.value != '') {
        titel('anmelden');
        if (lf.requestSubmit) lf.requestSubmit(lf.btnlogin); else lf.submit();
        return;
      }
      // Firefox traegt gespeicherte Daten kurz nach dem Laden ein; erst danach nach der Liste fragen
      if (!gemeldet && Date.now() - beginn > 700) {
        gemeldet = true;
        lf.login.focus();
        titel('passwort');
      }
      if (Date.now() - beginn > 20000) { erledigt(); return; }
      setTimeout(pruefen, 200);
    })();
  }

  function los() {
    auftragAusAdresse();
    auf = auftrag();
    if (!auf) return;
    var lf = document.forms.loginform;
    if (lf) { anmelden(lf); return; }
    // angemeldet: Auftrag erledigen (vorher austragen, damit die Zielseite ihn nicht erneut startet)
    erledigt();
    suchen();
  }

  // gleiche Seite mit neuem "#mor=" laedt nicht neu, sondern aendert nur den Anker
  window.addEventListener('hashchange', los);
  los();
})();
