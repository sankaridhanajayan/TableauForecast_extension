let forecastData = [];

tableau.extensions.initializeAsync({ 'configure': configure }).then(() => {
  const dashboard = tableau.extensions.dashboardContent.dashboard;

  // ── Register ALL event listeners on every worksheet ───────────────────
  dashboard.worksheets.forEach(sheet => {

    // Fires when any filter changes (Region, Date, etc.)
    sheet.addEventListener(
      tableau.TableauEventType.FilterChanged,
      handleChange
    );

    // Fires when marks are selected/deselected on the worksheet
    sheet.addEventListener(
      tableau.TableauEventType.MarkSelectionChanged,
      handleChange
    );

    // Fires when the underlying data in the worksheet changes
    sheet.addEventListener(
      tableau.TableauEventType.SummaryDataChanged,
      handleChange
    );

  });

  // ── Also listen to parameter changes at dashboard level ───────────────
  dashboard.getParametersAsync().then(parameters => {
    parameters.forEach(param => {
      param.addEventListener(
        tableau.TableauEventType.ParameterChanged,
        handleChange
      );
    });
  });

  // Load data on first run
  fetchForecastData();
});

// ── Single unified handler for all events ─────────────────────────────
function handleChange(event) {
  // Small debounce — avoids multiple rapid re-fetches if several
  // events fire together (e.g. filter + data change simultaneously)
  clearTimeout(window._reloadTimer);
  window._reloadTimer = setTimeout(() => {
    fetchForecastData();
  }, 300);
}

async function fetchForecastData() {
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  const forecastSheet = dashboard.worksheets.find(
    ws => ws.name === 'Forecast_sales'
  );

  if (!forecastSheet) {
    console.warn('Worksheet "Forecast" not found.');
    return;
  }

  try {
    const options = {
      ignoreAliases:    false,
      ignoreSelection:  true,   // include all rows even if marks selected
      includeAllColumns: true
    };

    const dataTable = await forecastSheet.getSummaryDataAsync(options);
    const colNames  = dataTable.columns.map(c => c.fieldName);
    forecastData    = parseTableauData(dataTable, colNames);

    computeSummary(forecastData);

  } catch (err) {
    console.error('fetchForecastData error:', err);
  }
}

function parseTableauData(dataTable, colNames) {
  const find = (keywords) =>
    colNames.find(c =>
      keywords.every(k => c.toLowerCase().includes(k.toLowerCase()))
    ) || '';

  const monthCol = find(['month', 'order']);
  const typeCol  = find(['forecast', 'indicator']);
  const salesCol = find(['sales']);

  return dataTable.data.map(row => {
    const get = (colName) => {
      if (!colName) return '';
      const idx = colNames.indexOf(colName);
      return idx >= 0 ? row[idx].formattedValue : '';
    };
    return {
      month: get(monthCol),
      type:  get(typeCol),
      sales: get(salesCol)
    };
  });
}

function parseNumber(str) {
  if (!str || str.trim() === '') return 0;
  return parseFloat(
    str.replace(/,/g, '').replace(/\$/g, '').trim()
  ) || 0;
}

function computeSummary(data) {
  let actualSum   = 0;
  let estimateSum = 0;

  data.forEach(row => {
    const val = parseNumber(row.sales);
    if (row.type === 'Actual')   actualSum   += val;
    if (row.type === 'Estimate') estimateSum += val;
  });

  const grandTotal = actualSum + estimateSum;

  const fmt = (n) => '$' + n.toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });

  document.getElementById('sum-actual').textContent   = fmt(actualSum);
  document.getElementById('sum-estimate').textContent = fmt(estimateSum);
  document.getElementById('sum-total').textContent    = fmt(grandTotal);
}

function configure() {}