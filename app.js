document.getElementById('fileInput').addEventListener('change', function(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(firstSheet);

    console.log("JSON Data:", jsonData);

    // Process data for all visualizations
    const rosterData = countStatus(jsonData);
    const lineChartData = countSignUps(jsonData);
    const scatterPlotData = calculateDaysWaited(jsonData);

    // Display all visualizations
    const totalWaiting = displayWaitingRoster(jsonData);
    const totalActive = displayActiveRoster(jsonData);
    const totalNonEmpty = totalWaiting + totalActive;

    displayLineChart(lineChartData);
    displayScatterPlot(scatterPlotData);
    displayLongWaitRoster(scatterPlotData, jsonData);
    displayOnderbrekingRoster(jsonData);
    displayAttendanceAnalysis(jsonData);
    displayGroepscodeRoster(jsonData);
  };

  reader.readAsArrayBuffer(file);
});

// Helper functions to handle different column name variations
function getValue(row, fieldName) {
  const variations = [
    fieldName,
    fieldName + ' ',
    fieldName.charAt(0).toUpperCase() + fieldName.slice(1),
    fieldName.charAt(0).toUpperCase() + fieldName.slice(1) + ' '
  ];

  for (const variation of variations) {
    if (row[variation] !== undefined) {
      return row[variation];
    }
  }
  return null;
}

function getNaam(row) {
  return getValue(row, 'Naam') || 'Unknown';
}

function getOpmerkingen(row) {
  return getValue(row, 'Opmerkingen') || 'Geen opmerkingen';
}

function getGroepscode(row) {
  return getValue(row, 'Groepscode') || '';
}

function getStatus(row) {
  return getValue(row, 'Status') || '';
}

function getDatumOpvolgingActie(row) {
  return getValue(row, 'Datum opvolging actie (als dit van toepassing  is)') || '';
}

function getStartInburgeringstermijn(row) {
  return getValue(row, 'Start inburgeringstermijn') || '';
}

function getEindeInburgeringstermijn(row) {
  return getValue(row, 'Einde inburgeringstermijn') || '';
}

function getLeerroute(row) {
  return getValue(row, 'Leerroute') || '';
}

function getAanwezigheidspercentageOpTotaal(row) {
  return getValue(row, 'Aanwezigheidspercentage op totaal aantal uren') || '';
}

function getAanwezigheidspercentageAfgelopenMaand(row) {
  return getValue(row, 'Aanwezigheidspercentage afgelopen maand VGR') || '';
}

function getKNM(row) {
  return getValue(row, 'KNM') || '';
}

// Helper function to convert decimal to percentage
function toPercentage(value) {
  if (value === undefined || value === null) return null;
  return parseFloat(value) * 100;
}

// Convert Excel date to JavaScript Date
function excelDateToJSDate(excelDate) {
  if (typeof excelDate === 'number') {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(excelEpoch.getTime() + (excelDate * 24 * 60 * 60 * 1000));
  } else if (typeof excelDate === 'string' || excelDate instanceof Date) {
    return new Date(excelDate);
  }
  return null;
}

// Function to count statuses, disregarding empty statuses
function countStatus(data) {
  const statusCounts = {};

  data.forEach(row => {
    const status = getStatus(row).trim();
    if (status) {
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    }
  });

  return statusCounts;
}

// Function to display waiting roster
function displayWaitingRoster(data) {
  const waitingStatuses = ['uitnodigen intake', 'uitnodigen', 'plaatsen', 'geplaatst', 'verplaatsen'];
  const statusCounts = countStatus(data);
  let totalWaiting = 0;
  let totalNonEmpty = 0;

  // Calculate total non-empty statuses
  Object.values(statusCounts).forEach(count => {
    totalNonEmpty += count;
  });

  let html = '<h3>Wachtlijst statussen</h3><table><tr><th>Status</th><th>Aantal</th></tr>';

  Object.entries(statusCounts).forEach(([status, count]) => {
    if (waitingStatuses.includes(status.toLowerCase())) {
      html += `<tr><td>${status}</td><td>${count}</td></tr>`;
      totalWaiting += count;
    }
  });

  const waitingPercentage = totalNonEmpty > 0 ? (totalWaiting / totalNonEmpty * 100).toFixed(2) : 0;
  html += `<tr><td><strong>Subtotal</strong></td><td><strong>${totalWaiting} (${waitingPercentage}%)</strong></td></tr>`;
  html += '</table>';

  document.getElementById('waitingRoster').innerHTML = html;
  return totalWaiting;
}

// Function to display active roster
function displayActiveRoster(data) {
  const waitingStatuses = ['uitnodigen intake', 'uitnodigen', 'plaatsen', 'geplaatst', 'verplaatsen'];
  const statusCounts = countStatus(data);
  let totalActive = 0;
  let totalNonEmpty = 0;

  // Calculate total non-empty statuses
  Object.values(statusCounts).forEach(count => {
    totalNonEmpty += count;
  });

  let html = '<h3>Actieve Statussen</h3><table><tr><th>Status</th><th>Aantal</th></tr>';

  Object.entries(statusCounts).forEach(([status, count]) => {
    if (!waitingStatuses.includes(status.toLowerCase())) {
      html += `<tr><td>${status}</td><td>${count}</td></tr>`;
      totalActive += count;
    }
  });

  const activePercentage = totalNonEmpty > 0 ? (totalActive / totalNonEmpty * 100).toFixed(2) : 0;
  html += `<tr><td><strong>Subtotal</strong></td><td><strong>${totalActive} (${activePercentage}%)</strong></td></tr>`;
  html += '</table>';

  document.getElementById('activeRoster').innerHTML = html;
  return totalActive;
}

// Function to display group code roster
function displayGroepscodeRoster(data) {
  const targetStatuses = ['plaatsen', 'verplaatsen'];
  const groupCounts = {};

  data.forEach(row => {
    const status = getStatus(row).toLowerCase().trim();
    const groepscode = getGroepscode(row).trim();

    if (targetStatuses.includes(status) && groepscode) {
      if (!groupCounts[groepscode]) {
        groupCounts[groepscode] = 0;
      }
      groupCounts[groepscode]++;
    }
  });

  let html = '<h3>Te plaatsen cursisten per groepscode:</h3><table><tr><th>Groepscode</th><th>Aantal</th></tr>';

  for (const [code, count] of Object.entries(groupCounts)) {
    html += `<tr><td>${code}</td><td>${count}</td></tr>`;
  }

  html += '</table>';
  document.getElementById('groepscodeRoster').innerHTML = html;
}

// Function to count sign-ups per month
function countSignUps(data) {
  const signUpsPerMonth = {};

  data.forEach(row => {
    const date = getValue(row, 'Aanmelddatum');
    if (date) {
      const dateObj = excelDateToJSDate(date);
      if (dateObj && !isNaN(dateObj.getTime())) {
        const year = dateObj.getFullYear();
        const month = dateObj.getMonth() + 1;
        const monthKey = `${year}-${month.toString().padStart(2, '0')}`;
        signUpsPerMonth[monthKey] = (signUpsPerMonth[monthKey] || 0) + 1;
      }
    }
  });

  return Object.entries(signUpsPerMonth)
    .map(([month, signUps]) => {
      const [year, monthNum] = month.split('-').map(Number);
      return { date: new Date(year, monthNum - 1), signUps };
    })
    .sort((a, b) => a.date - b.date);
}

// Function to calculate days waited
function calculateDaysWaited(data) {
  const today = new Date();
  return data.map(row => {
    const naam = getNaam(row);
    const aanmelddatum = getValue(row, 'Aanmelddatum');
    let startdatum = getValue(row, 'Startdatum (1e les)');
    if (!startdatum) {
      startdatum = today;
    }
    if (aanmelddatum) {
      const aanmelDate = excelDateToJSDate(aanmelddatum);
      const startDate = excelDateToJSDate(startdatum);
      if (aanmelDate && !isNaN(aanmelDate.getTime()) && startDate && !isNaN(startDate.getTime())) {
        const daysWaited = Math.round((startDate - aanmelDate) / (1000 * 60 * 60 * 24));
        return { naam, aanmelddatum: aanmelDate, daysWaited };
      }
    }
    return null;
  }).filter(Boolean);
}

// Function to display line chart
function displayLineChart(data) {
  if (data.length === 0) {
    console.log("No data to plot.");
    return;
  }

  // Clear previous chart
  d3.select("#lineChart").selectAll("*").remove();

  const margin = { top: 20, right: 20, bottom: 50, left: 50 };
  const width = 1100 - margin.left - margin.right;
  const height = 600 - margin.top - margin.bottom;

  const svg = d3.select("#lineChart")
    .append("svg")
    .attr("width", width + margin.left + margin.right)
    .attr("height", height + margin.top + margin.bottom)
    .append("g")
    .attr("transform", `translate(${margin.left},${margin.top})`);

  // Set date range to last two years
  const today = new Date();
  const twoYearsAgo = new Date(today.getFullYear() - 2, today.getMonth(), today.getDate());

  // X scale
  const x = d3.scaleTime()
    .domain([twoYearsAgo, today])
    .range([0, width]);

  svg.append("g")
    .attr("transform", `translate(0,${height})`)
    .call(d3.axisBottom(x).ticks(d3.timeMonth.every(1)).tickFormat(d3.timeFormat("%b %Y")));

  // Y scale
  const y = d3.scaleLinear()
    .domain([0, d3.max(data, d => d.signUps)])
    .range([height, 0]);

  svg.append("g")
    .call(d3.axisLeft(y));

  // Line
  const line = d3.line()
    .x(d => x(d.date))
    .y(d => y(d.signUps));

  svg.append("path")
    .datum(data)
    .attr("fill", "none")
    .attr("stroke", "steelblue")
    .attr("stroke-width", 2)
    .attr("d", line);

  // Dots
  svg.selectAll("dot")
    .data(data)
    .enter()
    .append("circle")
    .attr("cx", d => x(d.date))
    .attr("cy", d => y(d.signUps))
    .attr("r", 4)
    .attr("fill", "steelblue");
}

// Function to display scatter plot of days waited
function displayScatterPlot(data) {
  if (data.length === 0) {
    console.log("No data to plot.");
    return;
  }

  // Clear previous chart
  d3.select("#scatterPlot").selectAll("*").remove();

  const margin = { top: 20, right: 20, bottom: 50, left: 50 };
  const width = 1100 - margin.left - margin.right;
  const height = 600 - margin.top - margin.bottom;

  const svg = d3.select("#scatterPlot")
    .append("svg")
    .attr("width", width + margin.left + margin.right)
    .attr("height", height + margin.top + margin.bottom)
    .append("g")
    .attr("transform", `translate(${margin.left},${margin.top})`);

  // Set date range to last two years
  const today = new Date();
  const twoYearsAgo = new Date(today.getFullYear() - 2, today.getMonth(), today.getDate());

  // X scale
  const x = d3.scaleTime()
    .domain([twoYearsAgo, today])
    .range([0, width]);

  svg.append("g")
    .attr("transform", `translate(0,${height})`)
    .call(d3.axisBottom(x).ticks(d3.timeMonth.every(1)).tickFormat(d3.timeFormat("%b %Y")))
    .append("text")
    .attr("fill", "#000")
    .attr("x", width / 2)
    .attr("y", 40)
    .attr("text-anchor", "middle")
    .text("Aanmelddatum");

  // Y scale
  const y = d3.scaleLinear()
    .domain([0, 200])
    .range([height, 0]);

  svg.append("g")
    .call(d3.axisLeft(y))
    .append("text")
    .attr("fill", "#000")
    .attr("transform", "rotate(-90)")
    .attr("y", -40)
    .attr("x", -height / 2)
    .attr("text-anchor", "middle")
    .text("Days Waited");

  // Add tooltip
  const tooltip = d3.select("#scatterPlot")
    .append("div")
    .attr("class", "tooltip")
    .style("opacity", 0)
    .style("position", "absolute")
    .style("background-color", "white")
    .style("border", "1px solid #ccc")
    .style("border-radius", "5px")
    .style("padding", "10px")
    .style("pointer-events", "none");

  // Add dots with color coding
  svg.selectAll("dot")
    .data(data)
    .enter()
    .append("circle")
    .attr("cx", d => x(d.aanmelddatum))
    .attr("cy", d => y(d.daysWaited))
    .attr("r", 5)
    .attr("fill", d => d.daysWaited <= 56 ? "green" : "red")
    .on("mouseover", function(event, d) {
      tooltip.transition()
        .duration(200)
        .style("opacity", .9);
      tooltip.html(`Naam: ${d.naam || 'Onbekend'}<br/>Dagen in de wacht: ${d.daysWaited}`)
        .style("left", (event.pageX + 10) + "px")
        .style("top", (event.pageY - 28) + "px");
    })
    .on("mouseout", function(d) {
      tooltip.transition()
        .duration(500)
        .style("opacity", 0);
    });

  // Calculate and add trend line
  const trendLineData = calculateTrendLine(data);
  const line = d3.line()
    .x(d => x(d.aanmelddatum))
    .y(d => y(d.trendDaysWaited));

  svg.append("path")
    .datum(trendLineData)
    .attr("fill", "none")
    .attr("stroke", "blue")
    .attr("stroke-width", 2)
    .attr("d", line);
}

// Function to calculate trend line using linear regression
function calculateTrendLine(data) {
  const n = data.length;
  let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;

  data.forEach((d, i) => {
    const x = d.aanmelddatum.getTime();
    const y = d.daysWaited;
    sumX += x;
    sumY += y;
    sumXY += x * y;
    sumX2 += x * x;
  });

  const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
  const intercept = (sumY - slope * sumX) / n;

  return data.map(d => {
    const x = d.aanmelddatum.getTime();
    const trendDaysWaited = slope * x + intercept;
    return { aanmelddatum: d.aanmelddatum, trendDaysWaited };
  });
}

// Filter and display roster of people waiting too long
function displayLongWaitRoster(data, jsonData) {
  const today = new Date();
  const longWaitData = jsonData.filter(row => {
    const naam = getNaam(row);
    const aanmelddatum = getValue(row, 'Aanmelddatum');
    const startdatum = getValue(row, 'Startdatum (1e les)');

    // Only include if there is no Startdatum (still waiting)
    if (startdatum) {
      return false;
    }

    // Only include if there is an Aanmelddatum
    if (!aanmelddatum) {
      return false;
    }

    const aanmelDate = excelDateToJSDate(aanmelddatum);
    if (aanmelDate && !isNaN(aanmelDate.getTime())) {
      const daysWaited = Math.round((today - aanmelDate) / (1000 * 60 * 60 * 24));
      return daysWaited > 56; // More than 8 weeks
    }
    return false;
  });

  const rosterDiv = document.getElementById('longWaitRoster');
  let table = '<h3>Wachtenden die al langer dan 8 weken wachten:</h3><table><tr><th>Naam</th><th>Aanmelddatum</th><th>Dagen gewacht</th></tr>';

  longWaitData.forEach(row => {
    const naam = getNaam(row);
    const aanmelddatum = getValue(row, 'Aanmelddatum');
    let daysWaited = 'N/A';
    let formattedAanmelddatum = aanmelddatum ? new Date(aanmelddatum).toLocaleDateString() : 'N/A';

    if (aanmelddatum) {
      const aanmelDate = excelDateToJSDate(aanmelddatum);
      if (aanmelDate && !isNaN(aanmelDate.getTime())) {
        daysWaited = Math.round((today - aanmelDate) / (1000 * 60 * 60 * 24));
      }
    }

    table += `<tr><td>${naam}</td><td>${formattedAanmelddatum}</td><td>${daysWaited}</td></tr>`;
  });

  table += '</table>';
  rosterDiv.innerHTML = table;
}

// Display onderbreking roster
function displayOnderbrekingRoster(jsonData) {
  const onderbrekingData = jsonData.filter(row => {
    const status = getStatus(row);
    return status === 'onderbreking' || status === 'Onderbreking';
  });

  const rosterDiv = document.getElementById('onderbrekingRoster');
  let table = '<h3>Cursisten in Onderbreking</h3><table><tr><th>Naam</th><th>Opmerkingen</th></tr>';

  onderbrekingData.forEach(row => {
    const naam = getNaam(row);
    const opmerkingen = getOpmerkingen(row);
    table += `<tr><td>${naam}</td><td>${opmerkingen}</td></tr>`;
  });

  table += '</table>';
  rosterDiv.innerHTML = table;
}

// Display attendance analysis
function displayAttendanceAnalysis(jsonData) {
  // Filter out rows with 0% attendance
  const filteredData = jsonData.filter(row => {
    const totalAttendance = toPercentage(getValue(row, 'Aanwezigheidspercentage op totaal aantal uren'));
    const monthlyAttendance = toPercentage(getValue(row, 'Aanwezigheidspercentage afgelopen maand'));
    return (totalAttendance && totalAttendance > 0) || (monthlyAttendance && monthlyAttendance > 0);
  });

  // Create pie chart for total attendance
  createAttendancePieChart(filteredData, 'Aanwezigheidspercentage op totaal aantal uren', 'totalAttendancePieChart', 'Total Attendance');

  // Create pie chart for monthly attendance
  createAttendancePieChart(filteredData, 'Aanwezigheidspercentage afgelopen maand', 'monthlyAttendancePieChart', 'Monthly Attendance');

  // Create roster for low monthly attendance
  createLowAttendanceRoster(filteredData);
}

// Create pie chart for attendance
function createAttendancePieChart(data, attendanceField, containerId, title) {
  // Count students in different attendance ranges
  let above80 = 0;
  let between70and80 = 0;
  let below70 = 0;

  data.forEach(row => {
    const attendance = toPercentage(getValue(row, attendanceField));
    if (attendance && !isNaN(attendance)) {
      if (attendance >= 80) {
        above80++;
      } else if (attendance >= 70) {
        between70and80++;
      } else {
        below70++;
      }
    }
  });

  // Clear previous chart
  d3.select(`#${containerId}`).selectAll("*").remove();

  // Set dimensions
  const width = 400;
  const height = 400;
  const radius = Math.min(width, height) / 2;

  // Create SVG
  const svg = d3.select(`#${containerId}`)
    .append("svg")
    .attr("width", width)
    .attr("height", height)
    .append("g")
    .attr("transform", `translate(${width / 2}, ${height / 2})`);

  // Create pie chart data
  const pieData = [
    { label: "Meer dan 80%", value: above80, color: "#4CAF50" },
    { label: "70-80%", value: between70and80, color: "#FF9800" },
    { label: "Minder dan 70%", value: below70, color: "#F44336" }
  ].filter(d => d.value > 0); // Filter out empty categories

  // Set color scale
  const color = d3.scaleOrdinal()
    .domain(pieData.map(d => d.label))
    .range(pieData.map(d => d.color));

  // Create pie layout
  const pie = d3.pie()
    .value(d => d.value)
    .sort(null);

  // Create arc
  const arc = d3.arc()
    .innerRadius(0)
    .outerRadius(radius);

  // Add title
  svg.append("text")
    .attr("text-anchor", "middle")
    .attr("dy", "-1.5em")
    .text(title);

  // Add pie slices
  const arcs = svg.selectAll("arc")
    .data(pie(pieData))
    .enter()
    .append("g")
    .attr("class", "arc");

  arcs.append("path")
    .attr("d", arc)
    .attr("fill", d => color(d.data.label));

  // Add labels
  arcs.append("text")
    .attr("transform", d => `translate(${arc.centroid(d)})`)
    .attr("text-anchor", "middle")
    .text(d => {
      const total = pieData.reduce((sum, item) => sum + item.value, 0);
      const percentage = total > 0 ? Math.round((d.data.value / total) * 100) : 0;
      return percentage > 5 ? `${percentage}%` : '';
    })
    .style("font-size", "12px")
    .style("font-weight", "bold");

  // Add legend
  const legend = svg.selectAll(".legend")
    .data(pieData)
    .enter()
    .append("g")
    .attr("class", "legend")
    .attr("transform", (d, i) => `translate(-${radius}, ${i * 20 - radius})`);

  legend.append("rect")
    .attr("width", 18)
    .attr("height", 18)
    .attr("fill", d => color(d.label));

  legend.append("text")
    .attr("x", 24)
    .attr("y", 9)
    .attr("dy", ".35em")
    .text(d => `${d.label}: ${d.value}`)
    .style("font-size", "12px");
}

// Create roster for low monthly attendance
function createLowAttendanceRoster(data) {
  const lowAttendanceData = data.filter(row => {
    const monthlyAttendance = toPercentage(getValue(row, 'Aanwezigheidspercentage afgelopen maand'));
    return monthlyAttendance && monthlyAttendance > 0 && monthlyAttendance < 80;
  });

  const rosterDiv = document.getElementById('lowAttendanceRoster');
  let table = '<h3>Cursisten met een aanwezigheid onder (&lt;80%)</h3><table><tr><th>Naam</th><th>Maandelijkse Aanwezigheid %</th></tr>';

  lowAttendanceData.forEach(row => {
    const naam = getNaam(row);
    const attendance = toPercentage(getValue(row, 'Aanwezigheidspercentage afgelopen maand'));
    table += `<tr><td>${naam}</td><td>${attendance.toFixed(1)}%</td></tr>`;
  });

  table += '</table>';
  rosterDiv.innerHTML = table;
}
