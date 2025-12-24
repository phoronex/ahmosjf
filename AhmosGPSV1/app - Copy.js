// Main GPS Dashboard Application
const GPSDashboard = (function () {
    // App state
    const appState = {
        devices: {},
        expandedDevices: new Set(),
        selectedSensors: new Map(), // deviceId -> Set of sensor names
        charts: new Map(), // deviceId -> Chart instance
        viewModes: new Map(), // deviceId -> view mode ('graph' or 'table')
        sensorColors: {
            'acc': '#FF6384',
            'spd': '#36A2EB',
            'vBattery': '#FFCE56',
            'dBattery': '#4BC0C0',
            'weightRaw': '#9966FF',
            'weightCalibrated': '#FF9F40',
            'temp1': '#FF6384',
            'temp2': '#36A2EB',
            'temp_avg': '#FF3333',
            'hum1': '#4BC0C0',
            'hum2': '#9966FF',
            'hum_avg': '#33CCFF',
            'fuelLevel': '#9C27B0'
        },
        sensorUnits: {
            'acc': 'g',
            'spd': 'km/h',
            'vBattery': 'V',
            'dBattery': '%',
            'weightRaw': 'V',
            'weightCalibrated': 'kg',
            'temp1': '°C',
            'temp2': '°C',
            'temp_avg': '°C',
            'hum1': '%',
            'hum2': '%',
            'hum_avg': '%',
            'fuelLevel': '%'
        },
        sensorNames: {
            'acc': 'Acceleration',
            'spd': 'Speed',
            'vBattery': 'Battery Voltage',
            'dBattery': 'Battery Percentage',
            'weightRaw': 'Weight Raw',
            'weightCalibrated': 'Weight Calibrated',
            'temp1': 'Temperature 1',
            'temp2': 'Temperature 2',
            'temp_avg': 'Temperature Average',
            'hum1': 'Humidity 1',
            'hum2': 'Humidity 2',
            'hum_avg': 'Humidity Average',
            'fuelLevel': 'Fuel Level'
        }
    };

    // Initialize dashboard
    function initDashboard() {
        updateDateRange();
        console.log('GPS Dashboard initialized');
    }

    // Process device data
    function processDeviceData(jsonData) {
        console.log('Processing device data:', jsonData);

        appState.devices = {};
        appState.selectedSensors.clear();
        appState.expandedDevices.clear();
        appState.viewModes.clear();

        if (!jsonData || !jsonData.devices || Object.keys(jsonData.devices).length === 0) {
            console.error("No devices found in data");
            renderDevices();
            return;
        }

        Object.entries(jsonData.devices).forEach(([deviceId, deviceData]) => {
            const points = deviceData.points || [];
            if (!Array.isArray(points) || points.length === 0) {
                console.warn(`No data points for device ${deviceId}`);
                return;
            }

            console.log(`Processing ${deviceId} with ${points.length} points`);

            // Initialize sensor selection and view mode
            appState.selectedSensors.set(deviceId, new Set());
            appState.viewModes.set(deviceId, 'graph'); // Default to graph view

            // Extract available sensors from first point
            const firstPoint = points[0];
            let availableSensors = Object.keys(firstPoint).filter(key =>
                key !== 'dtt' && key !== 'dts'
            );

            console.log(`Available sensors for ${deviceId}:`, availableSensors);

            // Process weight calibration
            if (availableSensors.includes('weightRaw') && typeof WeightCalibration !== 'undefined') {
                const sourceUnit = deviceData.metadata?.weightUnit || 'mV';
                points.forEach(point => {
                    point.weightCalibrated = WeightCalibration.convertWeight(
                        point.weightRaw,
                        sourceUnit
                    );
                });
                if (!availableSensors.includes('weightCalibrated')) {
                    availableSensors.push('weightCalibrated');
                }
            }

            // Calculate averages for multiple sensors
            const tempSensors = availableSensors.filter(s => s.startsWith('temp'));
            const humSensors = availableSensors.filter(s => s.startsWith('hum'));

            if (tempSensors.length > 1) {
                points.forEach(point => {
                    const tempValues = tempSensors.map(s => point[s]).filter(v => v != null);
                    if (tempValues.length > 0) {
                        point.temp_avg = tempValues.reduce((a, b) => a + b, 0) / tempValues.length;
                    }
                });
                if (!availableSensors.includes('temp_avg')) {
                    availableSensors.push('temp_avg');
                }
            }

            if (humSensors.length > 1) {
                points.forEach(point => {
                    const humValues = humSensors.map(s => point[s]).filter(v => v != null);
                    if (humValues.length > 0) {
                        point.hum_avg = humValues.reduce((a, b) => a + b, 0) / humValues.length;
                    }
                });
                if (!availableSensors.includes('hum_avg')) {
                    availableSensors.push('hum_avg');
                }
            }

            // Auto-select first 4 sensors
            const selectedSet = appState.selectedSensors.get(deviceId);
            const sensorsToSelect = availableSensors.slice(0, 4);
            sensorsToSelect.forEach(sensor => selectedSet.add(sensor));

            // Generate unique colors for any additional sensors
            availableSensors.forEach(sensor => {
                if (!appState.sensorColors[sensor]) {
                    // Generate a unique color
                    const hue = (availableSensors.indexOf(sensor) * 137) % 360; // Golden angle
                    appState.sensorColors[sensor] = `hsl(${hue}, 70%, 60%)`;

                    // Generate default unit and name if not exists
                    if (!appState.sensorUnits[sensor]) {
                        appState.sensorUnits[sensor] = '';
                    }
                    if (!appState.sensorNames[sensor]) {
                        appState.sensorNames[sensor] = sensor.charAt(0).toUpperCase() + sensor.slice(1);
                    }
                }
            });

            // Calculate sensor statistics
            const sensorStats = {};
            availableSensors.forEach(sensor => {
                const values = points.map(point => point[sensor]).filter(v => v != null);
                if (values.length > 0) {
                    sensorStats[sensor] = {
                        min: Math.min(...values),
                        max: Math.max(...values),
                        avg: values.reduce((a, b) => a + b, 0) / values.length,
                        latest: values[values.length - 1],
                        unit: appState.sensorUnits[sensor] || '',
                        values: values
                    };
                }
            });

            appState.devices[deviceId] = {
                data: points,
                metadata: {
                    lastUpdate: deviceData.lastUpdate || Date.now(),
                    dataPoints: points.length,
                    interval: jsonData.date?.interval || 'N/A',
                    weightUnit: deviceData.metadata?.weightUnit || 'mV',
                    deviceType: deviceData.metadata?.deviceType || 'GPS Device',
                    location: deviceData.metadata?.location || 'Unknown'
                },
                availableSensors: availableSensors,
                sensorStats: sensorStats
            };

            // Auto-expand first device
            if (Object.keys(appState.devices).length === 1) {
                appState.expandedDevices.add(deviceId);
            }
        });

        // Update weight unit in sensorUnits if WeightCalibration module exists
        if (typeof WeightCalibration !== 'undefined') {
            appState.sensorUnits.weightCalibrated = WeightCalibration.getWeightUnit();
        }

        renderDevices();

        // Update date range
        if (jsonData.date) {
            updateDateRangeFromData(jsonData.date);
        }

        console.log('Processed devices:', Object.keys(appState.devices));
    }

    // Sample data generator
    function loadSampleData() {
        showLoading(true, 'Generating sample data...');

        // Generate sample data structure
        const sampleData = {
            date: {
                interval: "5 minutes",
                from: "2025-12-17 00:00:00",
                to: "2025-12-23 10:37:00",
                interval_count: 1,
                interval_unit: "minutes"
            },
            devices: {}
        };

        // Create 3 sample devices
        const deviceIds = ['GPS-001', 'GPS-002', 'GPS-003'];
        const startTime = new Date("2025-12-17T00:00:00").getTime();
        const endTime = new Date("2025-12-23T10:37:00").getTime();
        const timeDiff = endTime - startTime;
        const numPoints = 168; // 7 days * 24 hours = 168 data points

        deviceIds.forEach((deviceId, deviceIndex) => {
            const points = [];

            for (let i = 0; i < numPoints; i++) {
                const timestamp = startTime + (i * (timeDiff / numPoints));
                const hour = new Date(timestamp).getHours();

                // Generate realistic sensor values with patterns
                const weightRaw = generateWeightPattern(i, deviceIndex);
                const temp1 = 20 + Math.sin(i * 0.2 + deviceIndex) * 5 + Math.random() * 2;
                const temp2 = 22 + Math.sin(i * 0.25 + deviceIndex) * 4 + Math.random() * 1.5;
                const hum1 = 50 + Math.sin(i * 0.15 + deviceIndex) * 15 + Math.random() * 5;
                const hum2 = 55 + Math.sin(i * 0.18 + deviceIndex) * 12 + Math.random() * 4;

                const point = {
                    dtt: timestamp,
                    dts: timestamp + Math.random() * 10000,
                    spd: generateSpeedPattern(i, hour),
                    acc: Math.random() * 2,
                    vBattery: 12 + Math.sin(i * 0.1 + deviceIndex) * 1 + Math.random() * 0.2,
                    dBattery: Math.max(0, Math.min(100, 80 + Math.sin(i * 0.05 + deviceIndex) * 20 + Math.random() * 5)),
                    weightRaw: weightRaw,
                    temp1: temp1,
                    temp2: temp2,
                    hum1: hum1,
                    hum2: hum2,
                    fuelLevel: Math.max(0, Math.min(100, 60 + Math.sin(i * 0.08 + deviceIndex) * 30 + Math.random() * 8))
                };

                points.push(point);
            }

            sampleData.devices[deviceId] = {
                points: points,
                lastUpdate: Date.now(),
                metadata: {
                    weightUnit: 'mV',
                    deviceType: 'Truck',
                    location: deviceIndex === 0 ? 'Warehouse A' :
                        deviceIndex === 1 ? 'Route 66' : 'Highway 101'
                }
            };
        });

        // Process the sample data after a short delay
        setTimeout(() => {
            processDeviceData(sampleData);
            showLoading(false);
            if (typeof FileHandler !== 'undefined') {
                FileHandler.showNotification('Sample data loaded successfully!', 'success');
            }
        }, 1000);
    }

    // Helper function for weight patterns
    function generateWeightPattern(index, deviceIndex) {
        // Create weight events: 600±50, 1200±200, 3500±300
        const hour = Math.floor(index / 7); // 24 data points per day
        const day = Math.floor(hour / 24);

        // Different patterns for different devices
        if (deviceIndex === 0) {
            // Device 1: Regular pattern
            if (day < 2) return 600 + Math.random() * 100 - 50;
            if (day < 4) return 1200 + Math.random() * 400 - 200;
            if (day < 6) return 3500 + Math.random() * 600 - 300;
            return 600 + Math.random() * 100 - 50;
        } else if (deviceIndex === 1) {
            // Device 2: Alternating pattern
            const pattern = index % 48; // 48 data points = 1 day at 30-min intervals
            if (pattern < 12) return 605 + Math.random() * 20 - 10;
            if (pattern < 24) return 1205 + Math.random() * 100 - 50;
            if (pattern < 36) return 3550 + Math.random() * 200 - 100;
            return 607 + Math.random() * 20 - 10;
        } else {
            // Device 3: Random transitions
            const rand = Math.random();
            if (rand < 0.4) return 607 + Math.random() * 40 - 20;
            if (rand < 0.7) return 1210 + Math.random() * 200 - 100;
            return 3520 + Math.random() * 400 - 200;
        }
    }

    // Helper function for speed patterns
    function generateSpeedPattern(index, hour) {
        // Speed based on time of day
        if (hour >= 22 || hour < 6) {
            // Night: low speed
            return 20 + Math.random() * 20;
        } else if (hour >= 6 && hour < 9) {
            // Morning rush hour
            return 40 + Math.random() * 30;
        } else if (hour >= 16 && hour < 19) {
            // Evening rush hour
            return 35 + Math.random() * 25;
        } else {
            // Daytime
            return 60 + Math.random() * 40;
        }
    }

    // Render all devices
    function renderDevices() {
        const container = document.getElementById('devicesContainer');
        const emptyState = document.getElementById('emptyState');

        if (!container) {
            console.error('devicesContainer not found');
            return;
        }

        if (Object.keys(appState.devices).length === 0) {
            if (emptyState) emptyState.style.display = 'block';
            container.innerHTML = '';
            return;
        }

        if (emptyState) emptyState.style.display = 'none';
        container.innerHTML = '';

        Object.entries(appState.devices).forEach(([deviceId, device], index) => {
            const deviceCard = createDeviceCard(deviceId, device, index);
            container.appendChild(deviceCard);
        });
    }

    // Create device card element
    function createDeviceCard(deviceId, device, index) {
        const isExpanded = appState.expandedDevices.has(deviceId);
        const selectedSensors = appState.selectedSensors.get(deviceId) || new Set();
        const selectedCount = selectedSensors.size;
        const totalSensors = device.availableSensors.length;
        const viewMode = appState.viewModes.get(deviceId) || 'graph';

        const card = document.createElement('div');
        card.className = `device-card ${isExpanded ? 'expanded' : ''}`;

        card.innerHTML = `
            <div class="device-header" onclick="GPSDashboard.toggleDevice('${deviceId}')">
                <div class="device-info">
                    <span class="device-icon">${index % 2 === 0 ? '🚛' : '🚚'}</span>
                    <span class="device-id">${deviceId}</span>
                    <span class="status-badge ${device.data.length > 0 ? 'status-active' : 'status-inactive'}">
                        ${device.data.length > 0 ? 'ACTIVE' : 'INACTIVE'}
                    </span>
                    <span class="device-stats">
                        <span>📊 ${device.metadata.dataPoints} points</span>
                        <span>⏰ ${device.metadata.interval}</span>
                        <span>🕒 ${formatTimeAgo(device.metadata.lastUpdate)}</span>
                        <span>📈 ${selectedCount}/${totalSensors} sensors</span>
                    </span>
                </div>
                <div class="device-controls">
                    <div class="device-select-all" onclick="GPSDashboard.toggleAllSensors('${deviceId}', event)">
                        <input type="checkbox" ${selectedCount === totalSensors ? 'checked' : ''}>
                        <span>All Sensors</span>
                    </div>
                    <div class="device-toggle">
                        ${isExpanded ? '▼' : '▶'}
                    </div>
                </div>
            </div>
            <div class="device-content ${isExpanded ? 'expanded' : ''}">
                ${isExpanded ? createDeviceContent(deviceId, device) : ''}
            </div>
        `;
        return card;
    }

    // Create device content
    function createDeviceContent(deviceId, device) {
        const selectedSensors = Array.from(appState.selectedSensors.get(deviceId) || []);
        const viewMode = appState.viewModes.get(deviceId) || 'graph';

        return `
            <div class="sensor-controls" id="sensorControls_${deviceId}">
                ${createSensorCheckboxes(deviceId, device.availableSensors)}
            </div>
            
            <div class="device-view-controls">
                <button class="view-btn ${viewMode === 'graph' ? 'active' : ''}" 
                        onclick="GPSDashboard.setDeviceViewMode('${deviceId}', 'graph')">
                    <span>📊</span> Graph
                </button>
                <button class="view-btn ${viewMode === 'table' ? 'active' : ''}" 
                        onclick="GPSDashboard.setDeviceViewMode('${deviceId}', 'table')">
                    <span>📋</span> Table
                </button>
                <button class="view-btn" 
                        onclick="GPSDashboard.showDeviceSummary('${deviceId}')">
                    <span>📈</span> Summary
                </button>
            </div>
            
            <div class="sensor-summary" id="sensorSummary_${deviceId}">
                ${createSensorSummary(deviceId, device.sensorStats, selectedSensors)}
            </div>
            
            <div class="chart-container" id="chartContainer_${deviceId}" 
                 style="display: ${viewMode === 'graph' ? 'block' : 'none'}">
                <canvas id="chart_${deviceId}"></canvas>
            </div>
            
            <div class="table-container" id="tableContainer_${deviceId}"
                 style="display: ${viewMode === 'table' ? 'block' : 'none'}">
                ${createDataTable(deviceId, device)}
            </div>
            
            <div class="legend" id="legend_${deviceId}"></div>
        `;
    }

    // Create sensor checkboxes
    function createSensorCheckboxes(deviceId, availableSensors) {
        const selectedSensors = appState.selectedSensors.get(deviceId) || new Set();

        return availableSensors.map(sensor => {
            const isSelected = selectedSensors.has(sensor);
            const color = appState.sensorColors[sensor] || '#666666';
            const name = appState.sensorNames[sensor] || sensor;
            const isAvgSensor = sensor.includes('_avg');

            return `
                <label class="sensor-checkbox ${isAvgSensor ? 'avg-sensor' : ''}" 
                       style="border-color: ${color}">
                    <input type="checkbox" 
                           ${isSelected ? 'checked' : ''}
                           onchange="GPSDashboard.toggleSensor('${deviceId}', '${sensor}', this.checked)">
                    <div class="sensor-color" style="background-color: ${color}"></div>
                    ${name} ${isAvgSensor ? '(Avg)' : ''}
                </label>
            `;
        }).join('');
    }

    // Create sensor summary
    function createSensorSummary(deviceId, sensorStats, selectedSensors) {
        return Object.entries(sensorStats)
            .filter(([sensor]) => selectedSensors.includes(sensor))
            .map(([sensor, stats]) => {
                const name = appState.sensorNames[sensor] || sensor;
                const color = appState.sensorColors[sensor] || '#666666';
                const isAvgSensor = sensor.includes('_avg');

                // Calculate trend
                let trend = 'stable';
                if (stats.values && stats.values.length > 1) {
                    const recentAvg = stats.values.slice(-5).reduce((a, b) => a + b, 0) / 5;
                    const olderAvg = stats.values.slice(-10, -5).reduce((a, b) => a + b, 0) / 5;
                    if (recentAvg > olderAvg * 1.05) trend = 'up';
                    else if (recentAvg < olderAvg * 0.95) trend = 'down';
                }

                return `
                    <div class="summary-card" style="border-left-color: ${color}">
                        <h4>${name} ${isAvgSensor ? '(Average)' : ''}</h4>
                        <div class="summary-value">
                            ${stats.latest.toFixed(2)}
                            <span class="summary-unit">${stats.unit}</span>
                        </div>
                        <div class="summary-trend ${'trend-' + trend}">
                            ${trend === 'up' ? '↗' : trend === 'down' ? '↘' : '→'}
                            Avg: ${stats.avg.toFixed(2)} | Min: ${stats.min.toFixed(2)} | Max: ${stats.max.toFixed(2)}
                        </div>
                        <div class="timestamp">Last updated</div>
                    </div>
                `;
            }).join('');
    }

    // Create data table
    function createDataTable(deviceId, device) {
        const selectedSensors = Array.from(appState.selectedSensors.get(deviceId) || []);
        if (selectedSensors.length === 0 || !device.data || device.data.length === 0) {
            return '<p class="no-data">No sensors selected or no data available.</p>';
        }

        let tableHTML = `
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Time</th>
        `;

        // Add headers for selected sensors
        selectedSensors.forEach(sensor => {
            const name = appState.sensorNames[sensor] || sensor;
            tableHTML += `<th>${name}</th>`;
        });

        tableHTML += `</tr></thead><tbody>`;

        // Add data rows (limit to 100 rows for performance)
        const displayData = device.data.slice(-100);
        displayData.forEach((point, index) => {
            const time = new Date(point.dtt).toLocaleString();
            tableHTML += `<tr><td>${time}</td>`;

            selectedSensors.forEach(sensor => {
                const value = point[sensor];
                const formattedValue = value != null ? value.toFixed(2) : 'N/A';
                const unit = appState.sensorUnits[sensor] || '';
                tableHTML += `<td>${formattedValue} ${unit}</td>`;
            });

            tableHTML += `</tr>`;
        });

        tableHTML += `</tbody></table>`;

        if (device.data.length > 100) {
            tableHTML += `<p class="table-note">Showing last 100 of ${device.data.length} records</p>`;
        }

        return tableHTML;
    }

    // Render device chart
    function renderDeviceChart(deviceId) {
        const device = appState.devices[deviceId];
        if (!device) return;

        const selectedSensors = Array.from(appState.selectedSensors.get(deviceId) || []);
        if (selectedSensors.length === 0) {
            // Clear chart if no sensors selected
            if (appState.charts.has(deviceId)) {
                appState.charts.get(deviceId).destroy();
                appState.charts.delete(deviceId);
            }
            updateLegend(deviceId, []);
            return;
        }

        const canvas = document.getElementById(`chart_${deviceId}`);
        if (!canvas) return;

        const ctx = canvas.getContext('2d');

        // Destroy existing chart
        if (appState.charts.has(deviceId)) {
            appState.charts.get(deviceId).destroy();
        }

        // Prepare datasets
        const datasets = selectedSensors.map(sensor => {
            const color = appState.sensorColors[sensor] || '#666666';
            const name = appState.sensorNames[sensor] || sensor;
            const isAvgSensor = sensor.includes('_avg');

            return {
                label: name,
                data: device.data.map(point => ({
                    x: point.dtt,
                    y: point[sensor] || 0
                })),
                borderColor: color,
                backgroundColor: color + '20',
                borderWidth: isAvgSensor ? 3 : 2,
                borderDash: isAvgSensor ? [5, 5] : [],
                fill: false,
                tension: 0.2,
                pointRadius: isAvgSensor ? 0 : 2,
                pointHoverRadius: 6
            };
        });

        // Create chart
        try {
            const chart = new Chart(ctx, {
                type: 'line',
                data: { datasets },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    interaction: {
                        mode: 'index',
                        intersect: false
                    },
                    plugins: {
                        tooltip: {
                            mode: 'index',
                            intersect: false,
                            callbacks: {
                                title: function (context) {
                                    return new Date(context[0].parsed.x).toLocaleString();
                                },
                                label: function (context) {
                                    const label = context.dataset.label;
                                    const value = context.parsed.y;
                                    const sensorKey = selectedSensors[context.datasetIndex];
                                    const unit = appState.sensorUnits[sensorKey] || '';
                                    return `${label}: ${value.toFixed(2)} ${unit}`;
                                }
                            }
                        },
                        legend: {
                            display: false
                        }
                    },
                    scales: {
                        x: {
                            type: 'time',
                            time: {
                                unit: 'day',
                                displayFormats: {
                                    hour: 'MMM d HH:mm',
                                    day: 'MMM d'
                                }
                            },
                            title: {
                                display: true,
                                text: 'Time'
                            },
                            ticks: {
                                maxRotation: 45,
                                minRotation: 45
                            }
                        },
                        y: {
                            beginAtZero: true,
                            title: {
                                display: true,
                                text: 'Sensor Values'
                            }
                        }
                    }
                }
            });

            appState.charts.set(deviceId, chart);
            updateLegend(deviceId, selectedSensors);
        } catch (error) {
            console.error(`Error rendering chart for ${deviceId}:`, error);
        }
    }

    // Update sensor summary
    function updateSensorSummary(deviceId) {
        const device = appState.devices[deviceId];
        if (!device) return;

        const summaryContainer = document.getElementById(`sensorSummary_${deviceId}`);
        if (!summaryContainer) return;

        const selectedSensors = Array.from(appState.selectedSensors.get(deviceId) || []);
        summaryContainer.innerHTML = createSensorSummary(deviceId, device.sensorStats, selectedSensors);
    }

    // Update table
    function updateTable(deviceId) {
        const device = appState.devices[deviceId];
        if (!device) return;

        const tableContainer = document.getElementById(`tableContainer_${deviceId}`);
        if (tableContainer) {
            tableContainer.innerHTML = createDataTable(deviceId, device);
        }
    }

    // Update legend
    function updateLegend(deviceId, selectedSensors) {
        const legendContainer = document.getElementById(`legend_${deviceId}`);
        if (!legendContainer) return;

        legendContainer.innerHTML = selectedSensors.map(sensor => {
            const color = appState.sensorColors[sensor] || '#666666';
            const name = appState.sensorNames[sensor] || sensor;
            const isAvgSensor = sensor.includes('_avg');

            return `
                <div class="legend-item">
                    <div class="sensor-color" style="background-color: ${color}"></div>
                    <span>${name} ${isAvgSensor ? '(Avg)' : ''}</span>
                </div>
            `;
        }).join('');
    }

    // Helper functions
    function formatTimeAgo(timestamp) {
        const seconds = Math.floor((Date.now() - timestamp) / 1000);
        if (seconds < 60) return 'just now';
        const minutes = Math.floor(seconds / 60);
        if (minutes < 60) return `${minutes}m ago`;
        const hours = Math.floor(minutes / 60);
        if (hours < 24) return `${hours}h ago`;
        const days = Math.floor(hours / 24);
        return `${days}d ago`;
    }

    function updateDateRange() {
        const dateRangeContainer = document.getElementById('dateRange');
        if (!dateRangeContainer) return;

        const now = new Date();
        const weekAgo = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));

        dateRangeContainer.innerHTML = `
            <span>📅 From: ${weekAgo.toLocaleDateString()} ${weekAgo.toLocaleTimeString()}</span>
            <span>➡️</span>
            <span>📅 To: ${now.toLocaleDateString()} ${now.toLocaleTimeString()}</span>
        `;
    }

    function updateDateRangeFromData(dateInfo) {
        const dateRangeContainer = document.getElementById('dateRange');
        if (!dateRangeContainer) return;

        if (dateInfo.from && dateInfo.to) {
            dateRangeContainer.innerHTML = `
                <span>📅 From: ${dateInfo.from}</span>
                <span>➡️</span>
                <span>📅 To: ${dateInfo.to}</span>
                <span>⏱️ Interval: ${dateInfo.interval || 'N/A'}</span>
            `;
        }
    }

    function showLoading(show, text = 'Loading Sensor Data...') {
        const loading = document.getElementById('loading');
        const loadingText = document.getElementById('loadingText');

        if (loading) {
            loading.classList.toggle('active', show);
        }
        if (loadingText) {
            loadingText.textContent = text;
        }
    }

    // Public methods
    return {
        appState,
        initDashboard,
        processDeviceData,
        loadSampleData,

        toggleDevice: function (deviceId) {
            if (appState.expandedDevices.has(deviceId)) {
                appState.expandedDevices.delete(deviceId);
            } else {
                appState.expandedDevices.add(deviceId);
            }

            renderDevices();

            // If expanding, render chart
            if (appState.expandedDevices.has(deviceId)) {
                setTimeout(() => {
                    const viewMode = appState.viewModes.get(deviceId) || 'graph';
                    if (viewMode === 'graph') {
                        renderDeviceChart(deviceId);
                    }
                }, 100);
            }
        },

        toggleSensor: function (deviceId, sensor, isSelected) {
            let selectedSet = appState.selectedSensors.get(deviceId);
            if (!selectedSet) {
                selectedSet = new Set();
                appState.selectedSensors.set(deviceId, selectedSet);
            }

            if (isSelected) {
                selectedSet.add(sensor);
            } else {
                selectedSet.delete(sensor);
            }

            // Update view based on current mode
            const viewMode = appState.viewModes.get(deviceId) || 'graph';
            if (viewMode === 'graph') {
                renderDeviceChart(deviceId);
            } else {
                updateTable(deviceId);
            }

            // Update sensor summary
            updateSensorSummary(deviceId);

            // Update device card to reflect selection count
            const deviceCard = document.querySelector(`[data-device-id="${deviceId}"]`);
            if (deviceCard) {
                const selectedCount = selectedSet.size;
                const totalSensors = appState.devices[deviceId]?.availableSensors.length || 0;
                const countElement = deviceCard.querySelector('.device-stats span:last-child');
                if (countElement) {
                    countElement.textContent = `📈 ${selectedCount}/${totalSensors} sensors`;
                }
            }
        },

        toggleAllSensors: function (deviceId, event) {
            if (event) event.stopPropagation();
            const device = appState.devices[deviceId];
            if (!device) return;

            let selectedSet = appState.selectedSensors.get(deviceId);
            if (!selectedSet) {
                selectedSet = new Set();
                appState.selectedSensors.set(deviceId, selectedSet);
            }

            const allSelected = selectedSet.size === device.availableSensors.length;

            if (allSelected) {
                // Deselect all
                selectedSet.clear();
            } else {
                // Select all
                device.availableSensors.forEach(sensor => selectedSet.add(sensor));
            }

            // Update UI
            const viewMode = appState.viewModes.get(deviceId) || 'graph';
            if (viewMode === 'graph') {
                renderDeviceChart(deviceId);
            } else {
                updateTable(deviceId);
            }

            updateSensorSummary(deviceId);
            renderDevices();
        },

        setDeviceViewMode: function (deviceId, mode) {
            if (mode !== 'graph' && mode !== 'table') return;

            appState.viewModes.set(deviceId, mode);

            // Update UI
            const chartContainer = document.getElementById(`chartContainer_${deviceId}`);
            const tableContainer = document.getElementById(`tableContainer_${deviceId}`);

            if (chartContainer) chartContainer.style.display = mode === 'graph' ? 'block' : 'none';
            if (tableContainer) tableContainer.style.display = mode === 'table' ? 'block' : 'none';

            // Update view buttons
            const deviceContent = document.querySelector(`.device-card[data-device-id="${deviceId}"] .device-content`);
            if (deviceContent) {
                const viewBtns = deviceContent.querySelectorAll('.view-btn');
                viewBtns.forEach(btn => btn.classList.remove('active'));
                const activeBtn = deviceContent.querySelector(`.view-btn[onclick*="'${mode}'"]`);
                if (activeBtn) activeBtn.classList.add('active');
            }

            // If switching to graph mode, render chart
            if (mode === 'graph') {
                setTimeout(() => {
                    renderDeviceChart(deviceId);
                }, 100);
            } else if (mode === 'table') {
                updateTable(deviceId);
            }
        },

        showDeviceSummary: function (deviceId) {
            if (typeof SummaryAnalysis !== 'undefined') {
                SummaryAnalysis.showSummary(deviceId);
            }
        },

        expandAllDevices: function () {
            Object.keys(appState.devices).forEach(deviceId => {
                appState.expandedDevices.add(deviceId);
            });
            renderDevices();

            // Render all charts
            setTimeout(() => {
                Object.keys(appState.devices).forEach(deviceId => {
                    const viewMode = appState.viewModes.get(deviceId) || 'graph';
                    if (viewMode === 'graph') {
                        renderDeviceChart(deviceId);
                    }
                });
            }, 300);
        },

        collapseAllDevices: function () {
            appState.expandedDevices.clear();
            renderDevices();
        },

        selectAllSensors: function () {
            Object.keys(appState.devices).forEach(deviceId => {
                const device = appState.devices[deviceId];
                if (device && device.availableSensors) {
                    const selectedSet = appState.selectedSensors.get(deviceId) || new Set();
                    device.availableSensors.forEach(sensor => selectedSet.add(sensor));
                    appState.selectedSensors.set(deviceId, selectedSet);
                }
            });
            renderDevices();

            // Update all charts and tables
            setTimeout(() => {
                Object.keys(appState.devices).forEach(deviceId => {
                    const viewMode = appState.viewModes.get(deviceId) || 'graph';
                    if (viewMode === 'graph') {
                        renderDeviceChart(deviceId);
                    } else {
                        updateTable(deviceId);
                    }
                    updateSensorSummary(deviceId);
                });
            }, 300);
        },

        deselectAllSensors: function () {
            Object.keys(appState.devices).forEach(deviceId => {
                appState.selectedSensors.set(deviceId, new Set());
            });
            renderDevices();

            // Clear all charts
            Object.keys(appState.devices).forEach(deviceId => {
                if (appState.charts.has(deviceId)) {
                    appState.charts.get(deviceId).destroy();
                    appState.charts.delete(deviceId);
                }
                const legendContainer = document.getElementById(`legend_${deviceId}`);
                if (legendContainer) {
                    legendContainer.innerHTML = '';
                }
            });
        },

        // Add this to the return object in app.js:
        toggleAllViewMode: function () {
            // Toggle all devices between graph and table
            const allDeviceIds = Object.keys(appState.devices);
            const currentMode = appState.viewModes.get(allDeviceIds[0]) || 'graph';
            const newMode = currentMode === 'graph' ? 'table' : 'graph';

            allDeviceIds.forEach(deviceId => {
                this.setDeviceViewMode(deviceId, newMode);
            });

            // Update button text
            const toggleBtn = document.getElementById('toggleViewBtn');
            if (toggleBtn) {
                toggleBtn.innerHTML = newMode === 'graph' ? '<span>📊</span> Graph View' : '<span>📋</span> Table View';
            }
        },

        showSummaryView: function () {
            // Show summary for first device
            const devices = Object.keys(appState.devices);
            if (devices.length > 0) {
                if (typeof SummaryAnalysis !== 'undefined') {
                    SummaryAnalysis.showSummary(devices[0]);
                }
            } else {
                alert('No devices loaded. Please load data first.');
            }
        },

        refreshData: function () {
            showLoading(true, 'Refreshing data...');
            setTimeout(() => {
                showLoading(false);
                if (typeof FileHandler !== 'undefined') {
                    FileHandler.showNotification('Data refreshed!', 'success');
                }
            }, 1000);
        },

        exportDeviceData: function (deviceId) {
            const device = appState.devices[deviceId];
            if (!device) return;

            // Create CSV content
            const selectedSensors = Array.from(appState.selectedSensors.get(deviceId) || []);

            // CSV header
            let csv = 'Time,' + selectedSensors.map(s =>
                `${appState.sensorNames[s] || s} (${appState.sensorUnits[s] || ''})`
            ).join(',') + '\n';

            // CSV data
            device.data.forEach(point => {
                const time = new Date(point.dtt).toISOString();
                const row = [time, ...selectedSensors.map(s => point[s] || '')];
                csv += row.join(',') + '\n';
            });

            // Create download link
            const blob = new Blob([csv], { type: 'text/csv' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${deviceId}-data-${new Date().toISOString().slice(0, 10)}.csv`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            if (typeof FileHandler !== 'undefined') {
                FileHandler.showNotification(`Data exported for ${deviceId}`, 'success');
            }
        }
    };
})();

// Initialize when ready
document.addEventListener('DOMContentLoaded', function () {
    GPSDashboard.initDashboard();

    // Auto-load sample data for demo
    setTimeout(() => {
        if (Object.keys(GPSDashboard.appState.devices).length === 0) {
            console.log('Auto-loading sample data...');
            GPSDashboard.loadSampleData();
        }
    }, 1000);
});

// Make available globally
window.GPSDashboard = GPSDashboard;