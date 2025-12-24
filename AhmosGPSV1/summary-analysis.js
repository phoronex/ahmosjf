// Summary Analysis Module
const SummaryAnalysis = (function () {
    let groupingThreshold = 0.1; // 10%
    let minDuration = 5 * 60 * 1000; // 5 minutes in milliseconds

    // Set grouping threshold
    function updateGroupingThreshold(value) {
        groupingThreshold = parseFloat(value) / 100;
        if (GPSDashboard.appState.currentSummaryDevice) {
            analyzeDeviceEvents(GPSDashboard.appState.currentSummaryDevice);
        }
    }

    // Set minimum duration
    function updateMinDuration(value) {
        minDuration = value * 60 * 1000; // Convert minutes to milliseconds
        if (GPSDashboard.appState.currentSummaryDevice) {
            analyzeDeviceEvents(GPSDashboard.appState.currentSummaryDevice);
        }
    }

    // Analyze events for a specific sensor
    function analyzeSensorEvents(deviceId, sensorName) {
        const device = GPSDashboard.appState.devices[deviceId];
        if (!device || !device.data) return [];

        const values = device.data.map(point => ({
            timestamp: point.dtt,
            value: point[sensorName],
            raw: point
        })).filter(item => item.value != null);

        if (values.length === 0) return [];

        // Group similar values
        const groups = [];
        let currentGroup = {
            start: values[0].timestamp,
            end: values[0].timestamp,
            values: [values[0].value],
            rawPoints: [values[0].raw],
            min: values[0].value,
            max: values[0].value,
            avg: values[0].value
        };

        for (let i = 1; i < values.length; i++) {
            const currentValue = values[i].value;
            const lastValue = currentGroup.values[currentGroup.values.length - 1];
            const avgValue = currentGroup.avg;

            // Check if value belongs to current group
            const threshold = avgValue * groupingThreshold;
            if (Math.abs(currentValue - avgValue) <= threshold) {
                // Add to current group
                currentGroup.values.push(currentValue);
                currentGroup.rawPoints.push(values[i].raw);
                currentGroup.end = values[i].timestamp;
                currentGroup.min = Math.min(currentGroup.min, currentValue);
                currentGroup.max = Math.max(currentGroup.max, currentValue);
                currentGroup.avg = currentGroup.values.reduce((a, b) => a + b, 0) / currentGroup.values.length;
            } else {
                // Finalize current group if it meets duration requirement
                const duration = currentGroup.end - currentGroup.start;
                if (duration >= minDuration) {
                    groups.push({ ...currentGroup });
                }

                // Start new group
                currentGroup = {
                    start: values[i].timestamp,
                    end: values[i].timestamp,
                    values: [currentValue],
                    rawPoints: [values[i].raw],
                    min: currentValue,
                    max: currentValue,
                    avg: currentValue
                };
            }
        }

        // Add last group if it meets duration requirement
        const duration = currentGroup.end - currentGroup.start;
        if (duration >= minDuration) {
            groups.push({ ...currentGroup });
        }

        return groups;
    }

    // Analyze all sensors for a device
    function analyzeDeviceEvents(deviceId) {
        const device = GPSDashboard.appState.devices[deviceId];
        if (!device) return null;

        const results = {};
        const sensors = device.availableSensors || [];

        sensors.forEach(sensor => {
            const groups = analyzeSensorEvents(deviceId, sensor);
            if (groups.length > 0) {
                results[sensor] = {
                    groups: groups,
                    totalEvents: groups.length,
                    sensorName: GPSDashboard.appState.sensorNames[sensor] || sensor,
                    unit: GPSDashboard.appState.sensorUnits[sensor] || ''
                };
            }
        });

        return results;
    }

    // Generate summary HTML
    function generateSummaryHTML(deviceId, analysisResults) {
        if (!analysisResults || Object.keys(analysisResults).length === 0) {
            return '<p>No significant events detected with current thresholds.</p>';
        }

        let html = '';

        Object.entries(analysisResults).forEach(([sensor, data]) => {
            html += `
                <div class="event-group">
                    <div class="event-group-header">
                        <div class="event-group-title">
                            ${data.sensorName} (${data.unit})
                        </div>
                        <div class="event-group-stats">
                            ${data.totalEvents} events detected
                        </div>
                    </div>
                    
                    ${data.groups.map((group, index) => {
                const duration = (group.end - group.start) / (1000 * 60); // minutes
                const startTime = new Date(group.start).toLocaleString();
                const endTime = new Date(group.end).toLocaleString();
                const range = `${group.min.toFixed(2)} - ${group.max.toFixed(2)} ${data.unit}`;

                return `
                            <div class="event-details">
                                <div><strong>Event ${index + 1}:</strong> ${range}</div>
                                <div class="event-range">Duration: ${duration.toFixed(1)} minutes</div>
                                <div>Time: ${startTime} to ${endTime}</div>
                                <div>Average: ${group.avg.toFixed(2)} ${data.unit}</div>
                                <div>Data Points: ${group.values.length}</div>
                            </div>
                        `;
            }).join('')}
                </div>
            `;
        });

        return html;
    }

    // Add this function to the SummaryAnalysis module:
    function hideSummaryModal() {
        document.getElementById('summaryModal').classList.remove('active');
    }

    // Make sure it's included in the return statement:
    return {
        updateGroupingThreshold,
        updateMinDuration,
        analyzeSensorEvents,
        analyzeDeviceEvents,
        generateSummaryHTML,
        showSummary,
        hideSummaryModal,  // Add this line
        getThreshold: () => groupingThreshold,
        getMinDuration: () => minDuration
    };

    // Show summary for device
    function showSummary(deviceId) {
        const device = GPSDashboard.appState.devices[deviceId];
        if (!device) return;

        GPSDashboard.appState.currentSummaryDevice = deviceId;

        const results = analyzeDeviceEvents(deviceId);
        const html = generateSummaryHTML(deviceId, results);

        document.getElementById('summaryResults').innerHTML = html;
        document.getElementById('summaryModal').classList.add('active');
    }

    // Export functions
    return {
        updateGroupingThreshold,
        updateMinDuration,
        analyzeSensorEvents,
        analyzeDeviceEvents,
        generateSummaryHTML,
        showSummary,
        getThreshold: () => groupingThreshold,
        getMinDuration: () => minDuration
    };
})();