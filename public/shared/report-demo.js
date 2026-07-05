/* ============================================================================
   Swiggy Rider Recruitment CRM — Shared "Sample / Demo" Report Generator
   ----------------------------------------------------------------------------
   Produces a beautifully designed, Swiggy-styled PDF report populated with
   realistic dummy data. Used everywhere a "Download Sample Report" button is
   shown (agent Daily Report panel + admin/client Generated Reports Archive)
   so users always have something gorgeous to download even before the first
   real report is auto-generated at 6:30 PM IST.

   Fully namespaced under  window.SwiggyReport  to avoid any global clashes.
   Requires (loaded by the host page):
     - jsPDF            (window.jspdf)
     - jsPDF-AutoTable  (doc.autoTable)
     - Chart.js         (window.Chart)  — optional; charts gracefully skipped
   ========================================================================== */
(function () {
  'use strict';

  // ── Swiggy brand palette (RGB) ──────────────────────────────────────────────
  var SWIGGY        = [252, 128, 25];   // #FC8019
  var SWIGGY_DARK   = [232, 114, 15];   // #E8720F
  var SWIGGY_DEEP   = [193, 88, 6];     // deep amber for gradient tail
  var INK           = [18, 41, 63];     // #12293F
  var MUTED         = [100, 116, 139];

  var DISPO_META = {
    'Interested':     { color: [16, 185, 129], label: 'Interested' },
    'Followup':       { color: [59, 130, 246], label: 'Follow-up' },
    'Not Interested': { color: [148, 163, 184], label: 'Not Interested' },
    'CNR':            { color: [245, 158, 11], label: 'CNR (Not Received)' },
    'Switch Off':     { color: [249, 115, 22], label: 'Switch Off' },
    'CNC (Dead)':     { color: [239, 68, 68],  label: 'CNC (Dead)' },
    'DND':            { color: [190, 24, 93],  label: 'DND' }
  };

  function rgb(a) { return 'rgb(' + a[0] + ',' + a[1] + ',' + a[2] + ')'; }

  // ── Chart.js → PNG (offscreen, no animation) ────────────────────────────────
  function makeChartImage(config, w, h) {
    if (typeof window.Chart === 'undefined') return null;
    var holder = document.createElement('div');
    holder.style.cssText = 'position:fixed;left:-99999px;top:-99999px;width:' + w + 'px;height:' + h + 'px;';
    var canvas = document.createElement('canvas');
    canvas.width = w; canvas.height = h;
    holder.appendChild(canvas);
    document.body.appendChild(holder);
    var url = null;
    try {
      var chart = new window.Chart(canvas.getContext('2d'), Object.assign({}, config, {
        options: Object.assign({ animation: false, responsive: false, maintainAspectRatio: false, devicePixelRatio: 2 }, config.options || {})
      }));
      url = canvas.toDataURL('image/png', 1.0);
      chart.destroy();
    } catch (e) { url = null; }
    document.body.removeChild(holder);
    return url;
  }

  // ── Drawing helpers ─────────────────────────────────────────────────────────
  function gradientRect(doc, x, y, w, h, c1, c2, steps) {
    steps = steps || 60;
    var sw = w / steps;
    for (var i = 0; i < steps; i++) {
      var t = i / (steps - 1);
      doc.setFillColor(
        Math.round(c1[0] + (c2[0] - c1[0]) * t),
        Math.round(c1[1] + (c2[1] - c1[1]) * t),
        Math.round(c1[2] + (c2[2] - c1[2]) * t)
      );
      doc.rect(x + i * sw, y, sw + 0.6, h, 'F');
    }
  }

  function kpiCard(doc, x, y, w, h, label, value, color) {
    doc.setFillColor(255, 255, 255);
    doc.setDrawColor(226, 232, 240);
    doc.setLineWidth(0.3);
    doc.roundedRect(x, y, w, h, 2.5, 2.5, 'FD');
    doc.setFillColor(color[0], color[1], color[2]);
    doc.roundedRect(x, y, w, 2.3, 1, 1, 'F');
    var val = String(value);
    var fs = 17;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(fs);
    while (doc.getTextWidth(val) > w - 5 && fs > 8) { fs -= 1; doc.setFontSize(fs); }
    doc.setTextColor(color[0], color[1], color[2]);
    doc.text(val, x + w / 2, y + h / 2 + 1.5, { align: 'center' });
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7);
    doc.setTextColor(MUTED[0], MUTED[1], MUTED[2]);
    doc.text(String(label).toUpperCase(), x + w / 2, y + h - 3.5, { align: 'center' });
  }

  function sectionHeader(doc, y, title, color) {
    y = checkPage(doc, y, 18);
    color = color || SWIGGY;
    doc.setFillColor(color[0], color[1], color[2]);
    doc.roundedRect(14, y - 4.5, 3.2, 7, 1, 1, 'F');
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(color[0], color[1], color[2]);
    doc.text(title, 21, y + 1);
    doc.setTextColor(INK[0], INK[1], INK[2]);
    return y + 9;
  }

  function checkPage(doc, y, needed) {
    if (y + needed > 278) { doc.addPage(); return 18; }
    return y;
  }

  function addFooters(doc) {
    var total = doc.internal.getNumberOfPages();
    for (var i = 1; i <= total; i++) {
      doc.setPage(i);
      doc.setDrawColor(226, 232, 240);
      doc.setLineWidth(0.3);
      doc.line(14, 286, 196, 286);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5);
      doc.setTextColor(MUTED[0], MUTED[1], MUTED[2]);
      doc.text('Swiggy Rider Recruitment CRM  |  Sample / Demo Report', 14, 291);
      doc.text('Page ' + i + ' of ' + total, 196, 291, { align: 'right' });
    }
  }

  // ── Realistic demo dataset ──────────────────────────────────────────────────
  function demoData() {
    return {
      agentName: 'Demo Agent',
      counts: {
        'Interested': 18,
        'Followup': 12,
        'Not Interested': 21,
        'CNR': 15,
        'Switch Off': 7,
        'CNC (Dead)': 5,
        'DND': 4
      },
      interested: [
        ['98•••••210', 'Rahul Verma',   'Bike (2-Wheeler)', 'Koramangala'],
        ['97•••••884', 'Amit Sharma',   'Bike (2-Wheeler)', 'HSR Layout'],
        ['90•••••173', 'Sohel Khan',    'EV Scooter',       'Indiranagar'],
        ['88•••••452', 'Pradeep Nair',  'Bike (2-Wheeler)', 'Whitefield'],
        ['96•••••019', 'Vikas Yadav',   'Bicycle',          'BTM Layout'],
        ['70•••••638', 'Manish Gupta',  'Bike (2-Wheeler)', 'Marathahalli']
      ],
      followups: [
        ['98•••••771', 'Deepak Rao',    'Bike (2-Wheeler)', '05-07-2026', '11:30 AM'],
        ['91•••••205', 'Suresh Menon',  'EV Scooter',       '05-07-2026', '02:15 PM'],
        ['99•••••640', 'Arjun Singh',   'Bike (2-Wheeler)', '06-07-2026', '10:00 AM'],
        ['80•••••317', 'Farhan Ali',    'Bicycle',          '06-07-2026', '04:45 PM']
      ],
      converted: [
        ['98•••••210', 'Rahul Verma',   'Bike (2-Wheeler)', 'Koramangala', '04-07-2026'],
        ['90•••••173', 'Sohel Khan',    'EV Scooter',       'Indiranagar', '04-07-2026'],
        ['96•••••019', 'Vikas Yadav',   'Bicycle',          'BTM Layout',  '04-07-2026']
      ]
    };
  }

  // ── Master builder ───────────────────────────────────────────────────────────
  function build() {
    if (!window.jspdf || !window.jspdf.jsPDF) {
      alert('PDF engine is still loading — please try again in a moment.');
      return;
    }
    var jsPDF = window.jspdf.jsPDF;
    var doc = new jsPDF();
    var today = new Date();
    var dateStr = String(today.getDate()).padStart(2, '0') + '-' +
                  String(today.getMonth() + 1).padStart(2, '0') + '-' + today.getFullYear();
    var fileName = dateStr + '_Swiggy_Rider_Report_SAMPLE.pdf';

    var d = demoData();
    var counts = d.counts;
    var totalDialed = Object.keys(counts).reduce(function (s, k) { return s + counts[k]; }, 0);
    var interestedN = counts['Interested'];
    var followupN   = counts['Followup'];
    var convertedN  = d.converted.length;
    var convRate    = totalDialed ? Math.round((convertedN / totalDialed) * 100) : 0;

    // ===== HEADER (Swiggy gradient) =====
    gradientRect(doc, 0, 0, 210, 42, SWIGGY, SWIGGY_DEEP);
    doc.setFillColor(250, 204, 21);           // amber accent line
    doc.rect(0, 42, 210, 1.4, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(21);
    doc.text('SWIGGY RIDER RECRUITMENT', 14, 17);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(11);
    doc.setTextColor(255, 236, 214);
    doc.text('CRM', 14, 25);

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.setTextColor(255, 255, 255);
    doc.text('DAILY AGENT REPORT', 196, 14, { align: 'right' });
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    doc.setTextColor(255, 236, 214);
    doc.text(dateStr, 196, 21, { align: 'right' });
    doc.text('Agent: ' + d.agentName, 196, 27, { align: 'right' });

    // SAMPLE badge
    doc.setFillColor(250, 204, 21);
    doc.roundedRect(158, 31, 38, 6.5, 1.5, 1.5, 'F');
    doc.setTextColor(SWIGGY_DEEP[0], SWIGGY_DEEP[1], SWIGGY_DEEP[2]);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.5);
    doc.text('SAMPLE / DEMO DATA', 177, 35.3, { align: 'center' });

    doc.setTextColor(INK[0], INK[1], INK[2]);

    // ===== KPI CARDS =====
    var y = 52;
    var kpis = [
      { label: 'Total Dialed', value: totalDialed, color: SWIGGY },
      { label: 'Interested',   value: interestedN, color: [16, 185, 129] },
      { label: 'Follow-ups',   value: followupN,   color: [59, 130, 246] },
      { label: 'Onboarded',    value: convertedN,  color: [236, 72, 153] },
      { label: 'Deliveries',   value: convertedN,  color: [22, 163, 74] }
    ];
    var gap = 3, cardW = (182 - gap * 4) / 5, cardH = 24;
    kpis.forEach(function (k, i) { kpiCard(doc, 14 + i * (cardW + gap), y, cardW, cardH, k.label, k.value, k.color); });
    y += cardH + 6;

    // conversion-rate strip
    doc.setFillColor(255, 244, 233);
    doc.roundedRect(14, y, 182, 9, 2, 2, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(9);
    doc.setTextColor(SWIGGY_DARK[0], SWIGGY_DARK[1], SWIGGY_DARK[2]);
    doc.text('Onboarding Rate: ' + convRate + '%', 18, y + 5.9);
    doc.text('Interested Rate: ' + Math.round(interestedN / totalDialed * 100) + '%', 90, y + 5.9);
    doc.text('Contactable: ' + Math.round((interestedN + followupN) / totalDialed * 100) + '%', 150, y + 5.9);
    doc.setTextColor(INK[0], INK[1], INK[2]);
    y += 16;

    // ===== DISPOSITION DONUT + LEGEND =====
    var activeDispo = Object.keys(counts).filter(function (k) { return counts[k] > 0; });
    y = sectionHeader(doc, y, 'Call Disposition Breakdown', SWIGGY);

    var donutImg = makeChartImage({
      type: 'doughnut',
      data: {
        labels: activeDispo.map(function (k) { return DISPO_META[k].label; }),
        datasets: [{
          data: activeDispo.map(function (k) { return counts[k]; }),
          backgroundColor: activeDispo.map(function (k) { return rgb(DISPO_META[k].color); }),
          borderColor: '#ffffff', borderWidth: 3
        }]
      },
      options: { cutout: '60%', plugins: { legend: { display: false }, tooltip: { enabled: false } } }
    }, 520, 520);

    var dSize = 58, dX = 15, dY = y + 2;
    if (donutImg) {
      doc.addImage(donutImg, 'PNG', dX, dY, dSize, dSize);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(19);
      doc.setTextColor(INK[0], INK[1], INK[2]);
      doc.text(String(totalDialed), dX + dSize / 2, dY + dSize / 2 - 0.5, { align: 'center' });
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5);
      doc.setTextColor(MUTED[0], MUTED[1], MUTED[2]);
      doc.text('CALLS', dX + dSize / 2, dY + dSize / 2 + 5, { align: 'center' });
    }

    // legend with percentage bars
    var ly = dY + 2;
    var lx = 84, barX = 84, barW = 78;
    activeDispo.forEach(function (k) {
      var c = DISPO_META[k].color;
      var v = counts[k];
      var pct = Math.round(v / totalDialed * 100);
      doc.setFillColor(c[0], c[1], c[2]);
      doc.roundedRect(lx, ly, 3.2, 3.2, 0.6, 0.6, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(8.5);
      doc.setTextColor(INK[0], INK[1], INK[2]);
      doc.text(DISPO_META[k].label, lx + 5.5, ly + 3);
      doc.setTextColor(MUTED[0], MUTED[1], MUTED[2]);
      doc.text(v + '  (' + pct + '%)', 196, ly + 3, { align: 'right' });
      doc.setFillColor(233, 236, 242);
      doc.roundedRect(barX, ly + 4.4, barW, 1.8, 0.9, 0.9, 'F');
      doc.setFillColor(c[0], c[1], c[2]);
      doc.roundedRect(barX, ly + 4.4, Math.max(1.4, barW * (pct / 100)), 1.8, 0.9, 0.9, 'F');
      ly += 9.6;
    });
    y = Math.max(dY + dSize, ly) + 8;

    // ===== FUNNEL BAR CHART =====
    y = sectionHeader(doc, y, 'Performance Funnel', SWIGGY);
    var funnelImg = makeChartImage({
      type: 'bar',
      data: {
        labels: ['Dialed', 'Interested', 'Follow-ups', 'Onboarded'],
        datasets: [{
          data: [totalDialed, interestedN, followupN, convertedN],
          backgroundColor: [rgb(SWIGGY), rgb([16, 185, 129]), rgb([59, 130, 246]), rgb([236, 72, 153])],
          borderRadius: 6, barThickness: 26
        }]
      },
      options: {
        indexAxis: 'y',
        plugins: { legend: { display: false }, tooltip: { enabled: false } },
        scales: {
          x: { beginAtZero: true, grid: { color: '#eef0f4' }, ticks: { precision: 0, font: { size: 13 } } },
          y: { grid: { display: false }, ticks: { font: { size: 14, weight: 'bold' } } }
        }
      }
    }, 960, 300);
    if (funnelImg) { doc.addImage(funnelImg, 'PNG', 15, y, 182, 57); y += 62; }

    // ===== DETAIL PAGE =====
    doc.addPage();
    y = 18;

    y = sectionHeader(doc, y, 'Interested Riders (Today)', [16, 185, 129]);
    doc.autoTable({
      startY: y,
      head: [['Phone', 'Name', 'Vehicle Type', 'Area']],
      body: d.interested,
      theme: 'striped',
      headStyles: { fillColor: [16, 185, 129], fontStyle: 'bold', fontSize: 9 },
      margin: { left: 14, right: 14 },
      styles: { fontSize: 9, cellPadding: 4 },
      alternateRowStyles: { fillColor: [236, 253, 245] }
    });
    y = doc.lastAutoTable.finalY + 12;

    y = sectionHeader(doc, y, 'Follow-ups (Scheduled)', [59, 130, 246]);
    doc.autoTable({
      startY: y,
      head: [['Phone', 'Name', 'Vehicle Type', 'Date', 'Time']],
      body: d.followups,
      theme: 'striped',
      headStyles: { fillColor: [59, 130, 246], fontStyle: 'bold', fontSize: 9 },
      margin: { left: 14, right: 14 },
      styles: { fontSize: 9, cellPadding: 4 },
      alternateRowStyles: { fillColor: [239, 246, 255] }
    });
    y = doc.lastAutoTable.finalY + 12;

    y = sectionHeader(doc, y, 'Onboarded Riders (Today)', [236, 72, 153]);
    doc.setFillColor(253, 242, 248);
    doc.roundedRect(14, y - 3, 182, 9, 2, 2, 'F');
    doc.setFontSize(10); doc.setFont('helvetica', 'bold');
    doc.setTextColor(190, 24, 93);
    doc.text('Total Riders Onboarded:  ' + d.converted.length, 18, y + 3);
    doc.setTextColor(INK[0], INK[1], INK[2]);
    doc.setFont('helvetica', 'normal');
    y += 11;
    doc.autoTable({
      startY: y,
      head: [['Phone', 'Name', 'Vehicle Type', 'Area', 'Date']],
      body: d.converted,
      theme: 'striped',
      headStyles: { fillColor: [236, 72, 153], fontStyle: 'bold', fontSize: 9 },
      margin: { left: 14, right: 14 },
      styles: { fontSize: 9, cellPadding: 4 },
      alternateRowStyles: { fillColor: [253, 242, 248] }
    });

    addFooters(doc);
    doc.save(fileName);
  }

  // ── Public API ────────────────────────────────────────────────────────────────
  window.SwiggyReport = {
    downloadDemo: function () {
      try { build(); }
      catch (e) {
        console.error('Sample report generation failed:', e);
        alert('Could not generate the sample report. Please refresh and try again.');
      }
    }
  };
})();
