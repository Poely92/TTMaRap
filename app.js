document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    const reader = new FileReader();
  
    reader.onload = function(e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(firstSheet);
  
      console.log("JSON Data:", jsonData);
  
      // Process data for roster and line chart
      const rosterData = countStatus(jsonData);
      const lineChartData = countSignUps(jsonData);
  
      console.log("Roster Data:", rosterData);
      console.log("Line Chart Data:", lineChartData);
  
      // Display roster
      displayRoster(rosterData);
  
      // Display line chart
      displayLineChart(lineChartData);
    };
  
    reader.readAsArrayBuffer(file);
  });
  
  // Count occurrences of each status
  function countStatus(data) {
    const statusCounts = {};
    data.forEach(row => {
      const status = row['Status'];
      if (status) {
        statusCounts[status] = (statusCounts[status] || 0) + 1;
      }
    });
    return Object.entries(statusCounts).map(([status, count]) => ({ status, count }));
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
  
  // Count sign-ups per month
  function countSignUps(data) {
    const signUpsPerMonth = {};
    data.forEach(row => {
      const date = row['Aanmelddatum'];
      if (date) {
        const dateObj = excelDateToJSDate(date);
        if (dateObj && !isNaN(dateObj.getTime())) {
          // Extract the year and month
          const year = dateObj.getFullYear();
          const month = dateObj.getMonth() + 1; // Months are 0-indexed
          const monthKey = `${year}-${month.toString().padStart(2, '0')}`;
  
          signUpsPerMonth[monthKey] = (signUpsPerMonth[monthKey] || 0) + 1;
        }
      }
    });
  
    // Convert the object to an array of objects
    return Object.entries(signUpsPerMonth)
      .map(([month, signUps]) => {
        const [year, monthNum] = month.split('-').map(Number);
        return { date: new Date(year, monthNum - 1), signUps };
      })
      .sort((a, b) => a.date - b.date); // Sort by date
  }
  
  // Display roster as a table
  function displayRoster(data) {
    const rosterDiv = document.getElementById('roster');
    let table = '<table id="rosterTable"><thead><tr><th>Status</th><th>Count</th></tr></thead><tbody>';
    data.forEach(item => {
      table += `<tr><td>${item.status}</td><td>${item.count}</td></tr>`;
    });
    table += '</tbody></table>';
    rosterDiv.innerHTML = table;
  }
  
  // Display line chart
  function displayLineChart(data) {
    if (data.length === 0) {
      console.log("No data to plot.");
      return;
    }
  
    // Clear previous chart
    d3.select("#lineChart").selectAll("*").remove();
  
    const margin = { top: 20, right: 20, bottom: 30, left: 40 };
    const width = 600 - margin.left - margin.right;
    const height = 400 - margin.top - margin.bottom;
  
    const svg = d3.select("#lineChart")
      .append("svg")
      .attr("width", width + margin.left + margin.right)
      .attr("height", height + margin.top + margin.bottom)
      .attr("id", "lineChartSvg")
      .append("g")
      .attr("transform", `translate(${margin.left},${margin.top})`);
  
    // X scale
    const x = d3.scaleTime()
      .domain(d3.extent(data, d => d.date))
      .range([0, width]);
  
    svg.append("g")
      .attr("transform", `translate(0,${height})`)
      .call(d3.axisBottom(x).ticks(d3.timeMonth.every(1)));
  
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
  
  // Download as PDF
  document.getElementById('downloadPdf').addEventListener('click', function() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
  
    console.log("Creating PDF...");
  
    // Add title
    doc.setFontSize(16);
    doc.text('Roster and Sign-Ups Report', 10, 10);
  
    // Capture the roster table
    console.log("Adding roster table...");
    doc.autoTable({ html: '#rosterTable' });
  
    // Capture the line chart as an image
    console.log("Capturing line chart...");
    setTimeout(() => {
      const svgElement = document.getElementById('lineChartSvg').parentNode;
      html2canvas(svgElement).then(canvas => {
        console.log("Canvas captured:", canvas);
        const imgData = canvas.toDataURL('image/png');
        doc.addImage(imgData, 'PNG', 10, doc.autoTable.previous.finalY + 10, 180, 100);
        console.log("Saving PDF...");
        doc.save('roster_and_signups.pdf');
      }).catch(error => {
        console.error("Error capturing chart:", error);
      });
    }, 500); // Delay to ensure the chart is rendered
  });
  