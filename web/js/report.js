/* FEED M/H Calculator - Word report generation
 * Builds the same structure as create_word_report() in report_utils.py - cover,
 * Executive Summary, cost estimate, charts, Activity detail, calculation
 * standards and review comments - as HTML, downloaded as a .doc file.
 * Word opens the HTML document directly, so no library is needed. */
(function (global) {
  'use strict';

  var HYUNDAI_BLUE = '#004EA2';
  var DARK_BLUE = '#17365D';

  function fmt(v, d) { return MH.money(v, d || 0); }
  function nnum(v) { return MH.num(v); }
  function pct(a, b) { a = nnum(a); b = nnum(b); return b ? a / b * 100 : 0; }

  function esc(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  var TD = 'border:1px solid #9AA7B4;padding:4px 6px;font-size:9pt;';
  var TH = TD + 'background:' + DARK_BLUE + ';color:#FFFFFF;font-weight:bold;text-align:center;';

  function table(headers, rows, aligns, widths) {
    var h = '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:10px;">';
    h += '<tr>' + headers.map(function (t, i) {
      return '<th style="' + TH + (widths && widths[i] ? 'width:' + widths[i] + ';' : '') + '">' + esc(t) + '</th>';
    }).join('') + '</tr>';
    rows.forEach(function (r) {
      h += '<tr>' + r.map(function (c, i) {
        var a = aligns && aligns[i] ? aligns[i] : 'left';
        return '<td style="' + TD + 'text-align:' + a + ';">' + esc(c) + '</td>';
      }).join('') + '</tr>';
    });
    return h + '</table>';
  }

  function noteBox(title, lines, fill) {
    var h = '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:12px;">';
    h += '<tr><td style="border:1px solid #9AA7B4;padding:8px 10px;background:' + (fill || '#EDF4FB') + ';font-size:9pt;">';
    h += '<b>' + esc(title) + '</b><br>';
    h += lines.map(function (l) { return '• ' + esc(l); }).join('<br>');
    return h + '</td></tr></table>';
  }

  function heading(text, level) {
    var size = level === 1 ? '15pt' : '12pt';
    return '<p style="font-size:' + size + ';font-weight:bold;color:' + DARK_BLUE +
      ';margin:14px 0 6px;">' + esc(text) + '</p>';
  }

  /* Bars are drawn as table-cell widths so Word renders them reliably. */
  function bar(label, value, maxValue, color) {
    var w = maxValue > 0 ? Math.max(1, Math.round(value / maxValue * 100)) : 1;
    return '<tr>' +
      '<td style="' + TD + 'width:22%;font-weight:bold;">' + esc(label) + '</td>' +
      '<td style="' + TD + 'width:58%;padding:2px 4px;">' +
      '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;"><tr>' +
      '<td style="width:' + w + '%;background:' + color + ';height:14px;font-size:1pt;">&nbsp;</td>' +
      '<td style="width:' + (100 - w) + '%;font-size:1pt;">&nbsp;</td></tr></table></td>' +
      '<td style="' + TD + 'width:20%;text-align:right;">' + fmt(value) + ' M/H</td></tr>';
  }

  function chartSection(data) {
    var t = data.total, ci = data.ci, tel = data.tel;
    var maxAll = Math.max(nnum(t.internal), nnum(t.external), 1);
    var maxPart = Math.max(nnum(ci.internal), nnum(ci.external), nnum(tel.internal), nnum(tel.external), 1);
    var h = heading('2. M/H composition charts', 1);
    h += '<p style="font-size:9pt;font-weight:bold;margin:6px 0 4px;">Internal vs outsourced M/H</p>';
    h += '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:10px;">';
    h += bar('Internal M/H (' + fmt(pct(t.internal, t.total), 1) + '%)', nnum(t.internal), maxAll, HYUNDAI_BLUE);
    h += bar('Outsourced M/H (' + fmt(pct(t.external, t.total), 1) + '%)', nnum(t.external), maxAll, '#7FB3D5');
    h += '</table>';
    h += '<p style="font-size:9pt;font-weight:bold;margin:6px 0 4px;">Internal vs outsourced M/H by discipline</p>';
    h += '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:10px;">';
    h += bar('C&I internal', nnum(ci.internal), maxPart, HYUNDAI_BLUE);
    h += bar('C&I outsourced', nnum(ci.external), maxPart, '#70AD47');
    h += bar('Telecom internal', nnum(tel.internal), maxPart, HYUNDAI_BLUE);
    h += bar('Telecom outsourced', nnum(tel.external), maxPart, '#70AD47');
    h += '</table>';
    return h;
  }

  function buildHtml(data, logoB64) {
    var total = data.total || {}, ci = data.ci || {}, tel = data.tel || {};
    var project = data.project || '-';
    var base = data.base_mh, months = data.months;

    var h = '';
    h += '<div style="font-family:\'Malgun Gothic\',sans-serif;font-size:9pt;color:#172033;">';

    // Cover / header
    h += '<table cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;margin-bottom:10px;"><tr>';
    h += '<td style="width:40%;vertical-align:middle;">';
    if (logoB64) h += '<img src="data:image/png;base64,' + logoB64 + '" width="150" alt="HYUNDAI ENGINEERING">';
    else h += '<span style="font-size:14pt;font-weight:bold;color:' + HYUNDAI_BLUE + ';">HYUNDAI ENGINEERING</span>';
    h += '</td><td style="width:60%;text-align:right;vertical-align:middle;">';
    h += '<span style="font-size:18pt;font-weight:bold;color:' + DARK_BLUE + ';">FEED M/H Calculation Report</span><br>';
    h += '<span style="font-size:9pt;">Control &amp; Instrumentation / Telecommunication</span>';
    h += '</td></tr></table>';

    h += table(['Item', 'Detail', 'Item', 'Detail'], [
      ['Project', String(project), 'Prepared with', 'FEED M/H Calculator V0.4'],
      ['Outsourcing Minimization', String(data.external_min || 'No'), 'Base / Duration',
       fmt(base) + ' M/H per M/M / ' + fmt(months) + ' months'],
      ['Outsourcing unit rate',
       'Single ' + fmt(data.rate_short) + ' KRW / Comprehensive ' + fmt(data.rate_comp) + ' KRW',
       'Reporting scope', 'C&I + Telecom FEED Engineering M/H']
    ], ['left', 'left', 'left', 'left'], ['14%', '36%', '14%', '36%']);

    // 1. Executive Summary
    h += heading('1. Executive Summary', 1);
    h += table(['Item', 'Internal M/H', 'Outsourced M/H', 'Total M/H', 'Total M/M', 'Average manpower'],
      [['C&I', ci, ci], ['Telecom', tel, tel], ['Total', total, total]].map(function (x) {
        var o = x[1];
        return [x[0], fmt(o.internal), fmt(o.external), fmt(o.total), fmt(o.mm, 1), fmt(o.avg, 1)];
      }), ['left', 'right', 'right', 'right', 'right', 'right']);

    h += table(['Cost item', 'Basis', 'Estimated cost', 'Remarks'], [
      ['Outsourcing-Single estimate',
       'Outsourced M/H ' + fmt(total.external) + ' × unit rate ' + fmt(data.rate_short),
       fmt(data.short_cost) + ' KRW', 'VAT and contract terms reviewed separately'],
      ['Outsourcing-Comprehensive estimate',
       'Outsourced M/H ' + fmt(total.external) + ' × unit rate ' + fmt(data.rate_comp),
       fmt(data.comp_cost) + ' KRW', 'VAT and contract terms reviewed separately']
    ], ['left', 'left', 'right', 'left']);

    var totalMh = nnum(total.total);
    var bigger = nnum(ci.total) >= nnum(tel.total) ? 'C&I' : 'Telecom';
    h += noteBox('Key analysis', [
      'Total calculated effort is ' + fmt(totalMh) + ' M/H, split ' + fmt(pct(total.internal, totalMh), 1) +
        '% internal and ' + fmt(pct(total.external, totalMh), 1) + '% outsourced.',
      bigger + ' carries the larger share; the supporting detail is set out in the per-activity tables.',
      'Applying Outsourcing Minimization changes the internal/outsourced split, so the outsourcing estimate and the staffing plan should be reviewed together.',
      'This report covers the main FEED-stage scope: technical document review, data sheets, drawings, MTO/inform and 3D/SPI/CER work.'
    ]);

    // 2. charts
    h += chartSection(data);

    // 3. Activity detail
    h += heading('3. Activity calculation detail by discipline', 1);
    [['3.1 C&I detail', ci], ['3.2 Telecom detail', tel]].forEach(function (x) {
      h += heading(x[0], 2);
      var rows = (x[1].rows || []).slice(0, 45).map(function (r) {
        return [r.code, r.activity, r.unit, fmt(r.qty, 1), MH.diffLabel(r.diff),
                fmt(r.hec), fmt(r.ext), fmt(r.total)];
      });
      h += table(['Code', 'Activity', 'Unit', 'Qty', 'Difficulty', 'Internal M/H', 'Outsourced M/H', 'Total'], rows,
        ['center', 'left', 'center', 'right', 'center', 'right', 'right', 'right'],
        ['7%', '34%', '7%', '8%', '12%', '11%', '11%', '10%']);
    });

    // 4. calculation standards
    h += heading('4. Calculation basis and principles', 1);
    h += noteBox('Calculation basis summary', [
      'Internal and outsourced Unit M/H are applied per activity, driven by the entered quantities and the selected Project Conditions.',
      'Difficulty is derived from the Project standard difficulty, the specification / data sheet / drawing conditions and whether SPI is applied.',
      'M/M is Total M/H divided by Base M/H; average manpower is M/M divided by the design duration.',
      'The outsourcing estimate is the outsourced M/H multiplied by the outsourcing unit rate, for budget review purposes.'
    ], '#F8FBFE');

    var std = data.std || {};
    [['ci', 'Representative C&I calculation standards'],
     ['tel', 'Representative Telecom calculation standards']].forEach(function (x) {
      var entries = Object.keys(std[x[0]] || {}).map(function (k) { return std[x[0]][k]; }).slice(0, 18);
      if (!entries.length) return;
      h += heading(x[1], 2);
      var rows = entries.map(function (e) {
        function pack(o) {
          var parts = Object.keys(o || {}).filter(function (k) { return nnum(o[k]) !== 0; })
            .map(function (k) { return MH.diffLabel(k) + ':' + fmt(o[k], 2); });
          return parts.length ? parts.join('/') : '-';
        }
        return [e.activity || '', e.unit || '', pack(e.int), pack(e.ext),
                String(e.guide || '').slice(0, 420), 'Unit M/H by difficulty'];
      });
      h += table(['Activity', 'Unit', 'Internal basis', 'Outsourced basis', 'Guide', 'Remarks'], rows,
        ['left', 'center', 'left', 'left', 'left', 'left'], ['20%', '6%', '14%', '14%', '34%', '12%']);
    });

    // 5. review comments
    h += heading('5. Review comments', 1);
    h += noteBox('Review and follow-up actions', [
      'For high-quantity items the reliability of the entered quantity drives the total M/H directly, so final quantities must be confirmed before pricing.',
      'For heavily outsourced scope, the calculation conditions, outsourcing scope and deliverable level should be defined clearly to reduce variance in the estimate.',
      'The selected Project Conditions feed the difficulty grades and the calculation basis, so they must agree with the ITB and the client requirement review.',
      'This report is for FEED-stage M/H review; the calculation standards and quantities must be revisited on conversion to EPC or at the start of detailed design.'
    ], '#FFF8E6');

    h += '<p style="text-align:center;font-size:8pt;color:#666666;margin-top:18px;">Hyundai Engineering FEED M/H Calculator | Generated Report</p>';
    h += '</div>';
    return h;
  }

  function fullDocument(data, logoB64) {
    return '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head><meta charset="utf-8"><title>FEED M/H Calculation Report</title>' +
      '<style>@page{size:A4;margin:1.4cm;} body{font-family:"Malgun Gothic",sans-serif;}</style>' +
      '</head><body>' + buildHtml(data, logoB64) + '</body></html>';
  }

  function download(data, logoB64, filename) {
    var html = fullDocument(data, logoB64);
    var blob = new Blob(['﻿', html], { type: 'application/msword;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename || 'FEED_MH_Report_Hyundai.doc';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1500);
  }

  function preview(data, logoB64) {
    var w = window.open('', '_blank');
    if (!w) return false;
    w.document.write(fullDocument(data, logoB64));
    w.document.close();
    return true;
  }

  global.MHReport = { buildHtml: buildHtml, fullDocument: fullDocument, download: download, preview: preview };
})(window);
