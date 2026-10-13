(function (root) {
  'use strict';
  function forecast({ leads, conversion, value, rate, mode = 'project', churn = 0 }) {
    if (![leads, conversion, value, rate, churn].every(Number.isFinite) || leads < 0 || conversion < 0 || conversion > 100 || value < 0 || rate <= 0 || churn < 0 || churn > 100 || !['project', 'retainer'].includes(mode)) throw new RangeError('Check your figures.');
    const salesRate = leads * conversion / 100;
    const retention = Math.pow(1 - churn / 100, 1 / 12);
    let active = 0, adjustedActive = 0, previousSales = 0;
    const months = Array.from({ length: 60 }, (_, index) => {
      const cumulativeSales = salesRate * (index + 1);
      if (!Number.isFinite(cumulativeSales) || cumulativeSales > Number.MAX_SAFE_INTEGER) throw new RangeError('These figures are too large. Enter smaller values.');
      // Correct machine rounding near whole deals without advancing genuinely fractional sales.
      const nearest = Math.round(cumulativeSales);
      const totalSales = Math.floor(Math.abs(cumulativeSales - nearest) <= Number.EPSILON * Math.max(1, cumulativeSales) * 2 ? nearest : cumulativeSales);
      const sales = totalSales - previousSales;
      previousSales = totalSales;
      active += sales;
      adjustedActive = adjustedActive * retention + sales;
      const revenue = (mode === 'retainer' ? active : sales) * value;
      const adjustedRevenue = mode === 'retainer' ? adjustedActive * value : revenue * Math.pow(retention, index);
      return { sales, revenue, adjustedRevenue };
    });
    let accumulated = 0, adjustedAccumulated = 0;
    const years = Array.from({ length: 5 }, (_, index) => {
      const revenue = months.slice(index * 12, (index + 1) * 12).reduce((total, month) => total + month.revenue, 0);
      const adjustedRevenue = months.slice(index * 12, (index + 1) * 12).reduce((total, month) => total + month.adjustedRevenue, 0);
      accumulated += revenue;
      adjustedAccumulated += adjustedRevenue;
      const investment = rate * 12 * (index + 1);
      return { year: index + 1, revenue, accumulated, adjustedRevenue, adjustedAccumulated, investment, ratio: accumulated / investment };
    });
    const firstYearMonths = months.slice(0, 12);
    const revenue = firstYearMonths.reduce((total, month) => total + month.revenue, 0);
    const investment = rate * 12;
    const ratio = revenue / investment;
    if (![salesRate * 12, revenue, investment, ratio, ...months.flatMap(month => [month.revenue, month.adjustedRevenue]), ...years.flatMap(year => [year.revenue, year.accumulated, year.adjustedRevenue, year.adjustedAccumulated, year.investment, year.ratio])].every(Number.isFinite)) throw new RangeError('These figures are too large. Enter smaller values.');
    return { months: firstYearMonths, years, deals: firstYearMonths.reduce((total, month) => total + month.sales, 0), expectedDeals: salesRate * 12, revenue, investment, net: revenue - investment, ratio };
  }
  if (typeof module !== 'undefined' && module.exports) module.exports = forecast;
  else root.roiForecast = forecast;
})(typeof window !== 'undefined' ? window : globalThis);
