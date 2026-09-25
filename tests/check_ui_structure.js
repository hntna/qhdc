// tests/check_ui_structure.js
// Kiem tra bat bien cau truc index.html: cac id/attr ma JS phu thuoc phai luon ton tai.
const fs = require('fs');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

const REQUIRED_IDS = [
  'btnOpenVTBConverter','btnSaveToServer','btnOpenResultFile','btnReset',
  'btnLoadMultiFile','inputMultiFile','btnOpenServerDrafts','badgeServerDraftCount',
  'input_VT','input_ML','input_CDBR','input_CNTT','input_TD','input_CD','input_HT',
  'tabValidation','tabPreview','tabHierarchy','tabStrategy',
  'badgeValidationCount','badgeRowCount','badgeGroupCount',
  'btnReindexTT','btnFixOutlineGroups','btnSplitWalletDV','btnSplitWalletMang',
  'btnRestoreOriginalTT','selectReportUnit','btnReloadReport',
];
const REQUIRED_SUBTABS = ['subtabTongHop','subtabMang','subtabDichVu','subtabStratDetail'];

const missing = [];
for (const id of REQUIRED_IDS) {
  const re = new RegExp('id\\s*=\\s*["\']' + id + '["\']');
  if (!re.test(html)) missing.push('id=' + id);
}
for (const st of REQUIRED_SUBTABS) {
  if (!html.includes('data-subtab="' + st + '"')) missing.push('data-subtab=' + st);
}

if (missing.length) {
  console.error('FAIL - thieu ' + missing.length + ' phan tu bat buoc:');
  missing.forEach(m => console.error('  - ' + m));
  process.exit(1);
}
console.log('PASS - tat ca ' + (REQUIRED_IDS.length + REQUIRED_SUBTABS.length) + ' phan tu bat buoc con nguyen.');
process.exit(0);
