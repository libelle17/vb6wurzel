/* MOReiter-Labor: Lesezeichen-Skript fuer onlinebefunde.labor-staber.de (Firefox)
   Wird als Lesezeichen mit dem Schluesselwort "morlabor" angelegt (MOReiter-Labor-Lesezeichen.html
   importieren) und von MOReiter (Strg+Alt+B) aufgerufen mit
       morlabor Nachname;Vorname;TT.MM.JJJJ;Kennung;Nr
   Jeder Aufruf macht einen Schritt und meldet ihn MOReiter ueber den Seitentitel "MOR:Nr:...",
   Nr unterscheidet die Aufrufe, damit MOReiter keinen alten Titel fuer die Antwort haelt:
     MOR:laden      nicht auf dem Portal: Patientenseite wird geladen (dann erneuter Aufruf)
     MOR:anmelden   Anmeldeseite, Passwort vom Firefox eingetragen: Formular abgeschickt
     MOR:passwort   Anmeldeseite ohne Passwort: Kennung eingetragen, Feld hat den Fokus
                    (MOReiter waehlt dann die gespeicherte Anmeldung aus der Firefox-Liste)
     MOR:suche      angemeldet: Suche laeuft im Hintergrund
     MOR:fertig     erster Befund des passenden Patienten wird geoeffnet
     MOR:liste      kein Patient mit passendem Geburtsdatum: Suchliste wird gezeigt
     MOR:nichts     Suche ohne Treffer
     MOR:fehler ... sonstiger Fehler
   Keine Zeilenkommentare und kein Prozentzeichen im Code: das Skript wird zu einer Zeile
   zusammengezogen, und Firefox entschluesselt %xx in javascript:-Adressen. */
(function (arg) {
  var basis = '/onlinebefunde/index.php';
  var a;
  try { a = decodeURIComponent(arg.replace(/\+/g, ' ')); } catch (e) { a = arg; }
  var t = a.split(';');
  var name = (t[0] || '').trim(), vor = (t[1] || '').trim(), geb = (t[2] || '').trim(), kennung = (t[3] || '').trim();
  var nr = (t[4] || '').trim();
  function titel(s) { document.title = 'MOR:' + nr + ':' + s; }
  if (location.hostname != 'onlinebefunde.labor-staber.de') {
    titel('laden');
    location.href = 'https://onlinebefunde.labor-staber.de' + basis + '?func=patienten&cache=delete';
    return;
  }
  var lf = document.forms.loginform;
  if (lf) {
    if (lf.login.value == '' && kennung != '') lf.login.value = kennung;
    if (lf.pwd.value == '') {
      lf.login.focus();
      titel('passwort');
      return;
    }
    titel('anmelden');
    if (lf.requestSubmit) lf.requestSubmit(lf.btnlogin); else lf.submit();
    return;
  }
  titel('suche');
  function lies(r) {
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r.text().then(function (h) { return new DOMParser().parseFromString(h, 'text/html'); });
  }
  function ziel(tr) {
    var m = (tr.getAttribute('onclick') || '').match(/href\s*=\s*["']([^"']+)["']/);
    return m ? m[1] : '';
  }
  /* Zeilen der Ergebnistabelle mit Patienten-Nr. und Geburtsdatum (Spalte "Geb.Dat.") */
  function treffer(doc) {
    var tab = doc.querySelector('table.obtable');
    if (!tab) return [];
    var sp = -1, th = tab.querySelectorAll('thead th');
    for (var i = 0; i < th.length; i++) if (/Geb/.test(th[i].textContent)) sp = i;
    var erg = [], tr = tab.querySelectorAll('tbody tr[onclick]');
    for (var j = 0; j < tr.length; j++) {
      var z = ziel(tr[j]);
      if (z.indexOf('pid=') < 0) continue;
      var c = tr[j].cells, g = (sp >= 0 && c[sp]) ? c[sp].textContent.trim() : '';
      erg.push({ ziel: z, geb: g });
    }
    return erg;
  }
  /* erste Befundzeile auf der Seite eines Patienten */
  function befund(doc) {
    var tr = doc.querySelectorAll('table.obtable tbody tr[onclick]');
    for (var j = 0; j < tr.length; j++) {
      var z = ziel(tr[j]);
      if (/[?&]id=[1-9]/.test(z) && z.indexOf('offset=') >= 0) return z;
    }
    return '';
  }
  function liste() {
    titel('liste');
    var f = document.createElement('form');
    f.method = 'POST';
    f.action = basis + '?func=patienten';
    var e = document.createElement('input');
    e.type = 'hidden'; e.name = 'freitext'; e.value = name + ' ' + vor;
    var b = document.createElement('input');
    b.type = 'hidden'; b.name = 'suchstart'; b.value = '';
    f.appendChild(e); f.appendChild(b);
    document.body.appendChild(f);
    f.submit();
  }
  var daten = new URLSearchParams();
  daten.append('freitext', name + ' ' + vor);
  daten.append('suchstart', '');
  fetch(basis + '?func=patienten', { method: 'POST', body: daten, credentials: 'same-origin' })
    .then(lies)
    .then(function (doc) {
      var tr = treffer(doc);
      if (tr.length == 0) { titel('nichts'); return; }
      var passend = tr.filter(function (x) { return x.geb == geb; });
      if (passend.length == 0) { liste(); return; }
      /* der Reihe nach, bis ein Patient einen Befund hat; sonst den ersten passenden zeigen */
      var i = 0;
      function naechster() {
        if (i >= passend.length) {
          titel('fertig');
          location.href = '/onlinebefunde/' + passend[0].ziel;
          return;
        }
        var p = passend[i++];
        fetch('/onlinebefunde/' + p.ziel, { credentials: 'same-origin' }).then(lies).then(function (d) {
          var z = befund(d);
          if (z == '') { naechster(); return; }
          titel('fertig');
          location.href = '/onlinebefunde/' + z;
        });
      }
      naechster();
    })
    .catch(function (e) { titel('fehler ' + e.message); });
})('%s');
