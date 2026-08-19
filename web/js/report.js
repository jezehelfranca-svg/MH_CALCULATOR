/* FEED M/H Calculator - Word Report 생성
 * report_utils.py 의 create_word_report() 와 같은 구성(표지/Executive Summary/예가/차트/
 * Activity 상세/산출기준/검토 의견)을 HTML 로 만들어 .doc 파일로 내려받습니다.
 * Word 가 그대로 열 수 있는 HTML 문서이므로 별도 라이브러리가 필요 없습니다. */
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

  /* 막대그래프는 Word 가 확실히 표시할 수 있도록 표 셀 폭으로 그립니다. */
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
    var h = heading('2. M/H 구성 차트', 1);
    h += '<p style="font-size:9pt;font-weight:bold;margin:6px 0 4px;">내부/외주 M-H 구성</p>';
    h += '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:10px;">';
    h += bar('내부 M/H (' + fmt(pct(t.internal, t.total), 1) + '%)', nnum(t.internal), maxAll, HYUNDAI_BLUE);
    h += bar('외주 M/H (' + fmt(pct(t.external, t.total), 1) + '%)', nnum(t.external), maxAll, '#7FB3D5');
    h += '</table>';
    h += '<p style="font-size:9pt;font-weight:bold;margin:6px 0 4px;">분야별 내부/외주 M-H</p>';
    h += '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin-bottom:10px;">';
    h += bar('C&I 내부', nnum(ci.internal), maxPart, HYUNDAI_BLUE);
    h += bar('C&I 외주', nnum(ci.external), maxPart, '#70AD47');
    h += bar('Telecom 내부', nnum(tel.internal), maxPart, HYUNDAI_BLUE);
    h += bar('Telecom 외주', nnum(tel.external), maxPart, '#70AD47');
    h += '</table>';
    return h;
  }

  function buildHtml(data, logoB64) {
    var total = data.total || {}, ci = data.ci || {}, tel = data.tel || {};
    var project = data.project || '-';
    var base = data.base_mh, months = data.months;

    var h = '';
    h += '<div style="font-family:\'Malgun Gothic\',sans-serif;font-size:9pt;color:#172033;">';

    // 표지 / 머리말
    h += '<table cellspacing="0" cellpadding="0" style="width:100%;border-collapse:collapse;margin-bottom:10px;"><tr>';
    h += '<td style="width:40%;vertical-align:middle;">';
    if (logoB64) h += '<img src="data:image/png;base64,' + logoB64 + '" width="150" alt="HYUNDAI ENGINEERING">';
    else h += '<span style="font-size:14pt;font-weight:bold;color:' + HYUNDAI_BLUE + ';">HYUNDAI ENGINEERING</span>';
    h += '</td><td style="width:60%;text-align:right;vertical-align:middle;">';
    h += '<span style="font-size:18pt;font-weight:bold;color:' + DARK_BLUE + ';">FEED M/H 산출 보고서</span><br>';
    h += '<span style="font-size:9pt;">Control &amp; Instrumentation / Telecommunication</span>';
    h += '</td></tr></table>';

    h += table(['구분', '내용', '구분', '내용'], [
      ['Project', String(project), '작성 기준', 'FEED M/H Calculator V0.4'],
      ['외주최소화', String(data.external_min || 'No'), 'Base / Duration', fmt(base) + ' M/H per M/M / ' + fmt(months) + '개월'],
      ['외주 단가', '단종 ' + fmt(data.rate_short) + ' 원 / 종합 ' + fmt(data.rate_comp) + ' 원', '보고 범위', 'C&I + Telecom FEED Engineering M/H']
    ], ['left', 'left', 'left', 'left'], ['14%', '36%', '14%', '36%']);

    // 1. Executive Summary
    h += heading('1. Executive Summary', 1);
    h += table(['구분', '내부 M/H', '외주 M/H', 'Total M/H', 'Total M/M', '평균 투입인원'],
      [['C&I', ci, ci], ['Telecom', tel, tel], ['Total', total, total]].map(function (x) {
        var o = x[1];
        return [x[0], fmt(o.internal), fmt(o.external), fmt(o.total), fmt(o.mm, 1), fmt(o.avg, 1)];
      }), ['left', 'right', 'right', 'right', 'right', 'right']);

    h += table(['비용 항목', '산출 기준', '예상 비용', '비고'], [
      ['외주-단종 예가', '외주 M/H ' + fmt(total.external) + ' × 단가 ' + fmt(data.rate_short), fmt(data.short_cost) + ' 원', 'VAT 및 계약 조건 별도 검토'],
      ['외주-종합 예가', '외주 M/H ' + fmt(total.external) + ' × 단가 ' + fmt(data.rate_comp), fmt(data.comp_cost) + ' 원', 'VAT 및 계약 조건 별도 검토']
    ], ['left', 'left', 'right', 'left']);

    var totalMh = nnum(total.total);
    var bigger = nnum(ci.total) >= nnum(tel.total) ? 'C&I' : 'Telecom';
    h += noteBox('주요 분석', [
      '총 산출 M/H는 ' + fmt(totalMh) + ' M/H이며, 내부 ' + fmt(pct(total.internal, totalMh), 1) + '%, 외주 ' + fmt(pct(total.external, totalMh), 1) + '% 구성입니다.',
      '분야별로는 ' + bigger + ' 비중이 더 크며, 상세 산출 근거는 Activity별 산출 상세표에 정리했습니다.',
      '외주최소화 적용 시 내부/외주 배분비가 변경되므로, 외주 예가와 투입계획을 함께 검토해야 합니다.',
      '본 보고서는 FEED 단계의 기술자료 검토, Data Sheet, Drawing, MTO/Inform, 3D/SPI/CER 등 주요 역무 기준으로 작성되었습니다.'
    ]);

    // 2. 차트
    h += chartSection(data);

    // 3. Activity 상세
    h += heading('3. 분야별 Activity 산출 상세', 1);
    [['3.1 C&I 상세', ci], ['3.2 Telecom 상세', tel]].forEach(function (x) {
      h += heading(x[0], 2);
      var rows = (x[1].rows || []).slice(0, 45).map(function (r) {
        return [r.code, r.activity, r.unit, fmt(r.qty, 1), r.diff, fmt(r.hec), fmt(r.ext), fmt(r.total)];
      });
      h += table(['Code', 'Activity', 'Unit', 'Qty', '난이도', '내부 M/H', '외주 M/H', 'Total'], rows,
        ['center', 'left', 'center', 'right', 'center', 'right', 'right', 'right'],
        ['7%', '34%', '7%', '8%', '12%', '11%', '11%', '10%']);
    });

    // 4. 산출 기준
    h += heading('4. 산출 기준 및 적용 원칙', 1);
    h += noteBox('산출 기준 요약', [
      '수량 입력값과 Project Condition 선택값을 기준으로 Activity별 내부/외주 Unit M/H를 적용합니다.',
      '난이도는 Project 표준 난이도, Specification/Data Sheet/Drawing 조건, SPI 적용 여부 등을 기준으로 산정합니다.',
      'M/M은 Total M/H를 Base M/H로 나누어 산출하며, 평균 투입인원은 M/M을 설계기간으로 나누어 산정합니다.',
      '외주 예가는 외주 M/H와 외주 적용단가를 곱한 예산 검토용 값입니다.'
    ], '#F8FBFE');

    var std = data.std || {};
    [['ci', 'C&I 대표 산출 기준'], ['tel', 'Telecom 대표 산출 기준']].forEach(function (x) {
      var entries = Object.keys(std[x[0]] || {}).map(function (k) { return std[x[0]][k]; }).slice(0, 18);
      if (!entries.length) return;
      h += heading(x[1], 2);
      var rows = entries.map(function (e) {
        function pack(o) {
          var parts = Object.keys(o || {}).filter(function (k) { return nnum(o[k]) !== 0; })
            .map(function (k) { return k + ':' + fmt(o[k], 2); });
          return parts.length ? parts.join('/') : '-';
        }
        return [e.activity || '', e.unit || '', pack(e.int), pack(e.ext), String(e.guide || '').slice(0, 420), '난이도별 Unit M/H'];
      });
      h += table(['Activity', 'Unit', '내부 기준', '외주 기준', 'Guide', '비고'], rows,
        ['left', 'center', 'left', 'left', 'left', 'left'], ['20%', '6%', '14%', '14%', '34%', '12%']);
    });

    // 5. 검토 의견
    h += heading('5. 검토 의견', 1);
    h += noteBox('검토 및 후속 조치', [
      '대형 Quantity 항목은 입력 수량의 신뢰성이 전체 M/H에 직접적인 영향을 주므로, 견적 전 최종 수량 확인이 필요합니다.',
      '외주 비중이 큰 역무는 산출조건, 외주 Scope, 제출물 수준을 명확히 정의하여 예가 산정 편차를 줄이는 것이 필요합니다.',
      'Project Condition 선택값이 난이도와 산출 기준에 반영되므로, ITB 및 발주처 요구사항 검토 결과와 일치하는지 확인해야 합니다.',
      '본 보고서는 FEED 단계의 M/H 산출 검토용이며, EPC 전환 또는 상세설계 착수 시 산출기준과 수량을 재검토해야 합니다.'
    ], '#FFF8E6');

    h += '<p style="text-align:center;font-size:8pt;color:#666666;margin-top:18px;">Hyundai Engineering FEED M/H Calculator | Generated Report</p>';
    h += '</div>';
    return h;
  }

  function fullDocument(data, logoB64) {
    return '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head><meta charset="utf-8"><title>FEED M/H 산출 보고서</title>' +
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
