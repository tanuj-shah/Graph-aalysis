const XLSX = require('xlsx');
const fs = require('fs');

const wb = XLSX.readFile('data.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

const DATE = '2026-04-03';
const FUEL_TYPES = ['Petrol', 'Diesel', 'Kerosene', 'ATF_DP', 'ATF_DF', 'LP_Gas'];
const FUEL_COLS = [2, 3, 4, 5, 6, 7]; // columns C-H

function cleanNum(val) {
  if (val === null || val === undefined) return '';
  let s = String(val).trim();
  if (s === '' || s === '-' || s.replace(/\s+/g, '') === '-') return '';
  // Strip currency symbols and commas
  s = s.replace(/[₹$Rs,\s]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? '' : n;
}

function unitFor(fuel) {
  return fuel === 'LP_Gas' ? 'NRs_per_cylinder' : 'NRs_per_KL';
}

const output = [];

// Row mapping: [excelRow, variable_name, variable_group]
const mapping = [
  // tax
  [11, 'custom_duty', 'tax'],
  [13, 'custom_service_charge', 'tax'],
  [14, 'road_maintenance_charge', 'tax'],
  [15, 'pollution_charge', 'tax'],
  [16, 'green_tax', 'tax'],
  [17, 'psf_transfer', 'tax'],
  [18, 'infrastructure_development_tax', 'tax'],
  [19, 'vat_13pct', 'tax'],
  [20, 'fuel_throughput_charge', 'tax'],
  [21, 'total_government_tax', 'tax'],

  // cost
  [9,  'ex_depot_price_irs_per_kl', 'cost'],
  [10, 'ex_depot_price_nrs_per_kl', 'cost'],
  [27, 'landed_cost', 'cost'],
  [22, 'administration_cost', 'cost'],
  [23, 'insurance_premium', 'cost'],
  [24, 'technical_loss', 'cost'],
  [25, 'total_noc_expenses', 'cost'],
  [26, 'transportation', 'cost'],
  [29, 'petroleum_dealer_expenses', 'cost'],
  [30, 'total_cost_noc_without_vat', 'cost'],
  [32, 'total_cost_noc_with_vat_per_kl', 'cost'],
  [33, 'total_cost_noc_per_litre', 'cost'],

  // selling_price
  [34, 'retail_price_per_litre_cylinder', 'selling_price'],
  [45, 'rsp_india_border_irs', 'selling_price'],
  [46, 'rsp_india_border_nrs', 'selling_price'],
  [47, 'rsp_nepal_border_birgunj', 'selling_price'],
  [28, 'noc_ex_depot_price_without_vat', 'selling_price'],

  // demand_profit
  [35, 'fortnightly_projected_demand', 'demand_profit'],
  [36, 'total_profit_loss_per_litre', 'demand_profit'],
  [37, 'total_profit_loss_in_crore', 'demand_profit'],
  [40, 'government_revenue_per_litre', 'demand_profit'],
  [41, 'total_govt_revenue_crore_fortnightly', 'demand_profit'],
];

for (const [rowIdx, varName, varGroup] of mapping) {
  const row = rows[rowIdx];
  if (!row) continue;
  for (let i = 0; i < FUEL_TYPES.length; i++) {
    const fuel = FUEL_TYPES[i];
    const rawVal = row[FUEL_COLS[i]];
    const val = cleanNum(rawVal);
    output.push({
      date: DATE,
      fuel_type: fuel,
      variable_group: varGroup,
      variable_name: varName,
      unit: unitFor(fuel),
      value: val
    });
  }
}

// Build CSV
const headers = ['date', 'fuel_type', 'variable_group', 'variable_name', 'unit', 'value'];
const csvLines = [headers.join(',')];
const unparsed = [];

for (const r of output) {
  const vals = headers.map(h => {
    const v = r[h];
    if (typeof v === 'string' && v.includes(',')) return `"${v}"`;
    return v;
  });
  csvLines.push(vals.join(','));
  if (r.value !== '' && isNaN(parseFloat(r.value))) {
    unparsed.push(r);
  }
}

const outFile = `noc_petroleum_pricing_${DATE}.csv`;
fs.writeFileSync(outFile, csvLines.join('\n'), 'utf-8');
console.log(`Written: ${outFile} (${output.length} rows)\n`);

// Validation
console.log('--- Rows per fuel_type ---');
const fuelCounts = {};
for (const r of output) { fuelCounts[r.fuel_type] = (fuelCounts[r.fuel_type] || 0) + 1; }
for (const [k, v] of Object.entries(fuelCounts)) console.log(`  ${k}: ${v}`);

console.log('\n--- Rows per variable_group ---');
const groupCounts = {};
for (const r of output) { groupCounts[r.variable_group] = (groupCounts[r.variable_group] || 0) + 1; }
for (const [k, v] of Object.entries(groupCounts)) console.log(`  ${k}: ${v}`);

console.log('\n--- Rows where value could not be parsed as float ---');
if (unparsed.length === 0) console.log('  None');
else unparsed.forEach(r => console.log(`  ${r.fuel_type} / ${r.variable_name}: "${r.value}"`));

console.log('\n--- First 10 rows ---');
console.log(headers.join(' | '));
for (let i = 0; i < 10; i++) {
  console.log(headers.map(h => output[i][h]).join(' | '));
}
