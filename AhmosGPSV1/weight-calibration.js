// Weight Calibration Module
const WeightCalibration = (function() {
    let calibrationPoints = [];
    let weightUnit = 'mV';
    let customFormula = null;

    // Initialize with default points
    function init(points) {
        calibrationPoints = points || [];
        weightUnit = localStorage.getItem('weightUnit') || 'mV';
        loadSavedCalibration();
        updateWeightUnitDisplay();
    }

    // Load saved calibration from localStorage
    function loadSavedCalibration() {
        const saved = localStorage.getItem('weightCalibration');
        if (saved) {
            try {
                calibrationPoints = JSON.parse(saved);
            } catch (e) {
                console.error('Error loading calibration:', e);
            }
        }
    }

    // Save calibration to localStorage
    function saveCalibration() {
        localStorage.setItem('weightCalibration', JSON.stringify(calibrationPoints));
        localStorage.setItem('weightUnit', weightUnit);
    }

    // Set weight unit (mV, tons, kg)
    function setWeightUnit(unit) {
        weightUnit = unit;
        document.getElementById('weightUnit').value = unit;
        saveCalibration();
        
        // Recalculate all calibrated weights
        if (GPSDashboard && GPSDashboard.appState) {
            Object.keys(GPSDashboard.appState.devices).forEach(deviceId => {
                calibrateDeviceWeight(deviceId);
            });
        }
    }

    // Add calibration point
    function addCalibrationPoint() {
        const pointsContainer = document.getElementById('calibrationPoints');
        const index = calibrationPoints.length;
        
        calibrationPoints.push({ raw: 0, calibrated: 0 });
        
        const pointDiv = document.createElement('div');
        pointDiv.className = 'calibration-point';
        pointDiv.innerHTML = `
            <span>Point ${index + 1}:</span>
            <input type="number" placeholder="Raw value" 
                   value="0" step="0.01"
                   onchange="WeightCalibration.updatePoint(${index}, 'raw', this.value)">
            <input type="number" placeholder="Calibrated value" 
                   value="0" step="0.01"
                   onchange="WeightCalibration.updatePoint(${index}, 'calibrated', this.value)">
            <button onclick="WeightCalibration.removePoint(${index})">×</button>
        `;
        
        pointsContainer.appendChild(pointDiv);
    }

    // Update calibration point
    function updatePoint(index, field, value) {
        if (calibrationPoints[index]) {
            calibrationPoints[index][field] = parseFloat(value) || 0;
            saveCalibration();
        }
    }

    // Remove calibration point
    function removePoint(index) {
        calibrationPoints.splice(index, 1);
        renderCalibrationPoints();
        saveCalibration();
    }

    // Render calibration points in modal
    function renderCalibrationPoints() {
        const pointsContainer = document.getElementById('calibrationPoints');
        pointsContainer.innerHTML = '';
        
        calibrationPoints.forEach((point, index) => {
            const pointDiv = document.createElement('div');
            pointDiv.className = 'calibration-point';
            pointDiv.innerHTML = `
                <span>Point ${index + 1}:</span>
                <input type="number" placeholder="Raw value" 
                       value="${point.raw}" step="0.01"
                       onchange="WeightCalibration.updatePoint(${index}, 'raw', this.value)">
                <input type="number" placeholder="Calibrated value" 
                       value="${point.calibrated}" step="0.01"
                       onchange="WeightCalibration.updatePoint(${index}, 'calibrated', this.value)">
                <button onclick="WeightCalibration.removePoint(${index})">×</button>
            `;
            pointsContainer.appendChild(pointDiv);
        });
    }

    // Calculate calibrated weight using linear interpolation
    function calculateCalibratedWeight(rawValue) {
        if (calibrationPoints.length < 2) {
            // No calibration, return raw value
            return rawValue;
        }
        
        // Sort points by raw value
        const sortedPoints = [...calibrationPoints].sort((a, b) => a.raw - b.raw);
        
        // Find the segment containing the raw value
        for (let i = 0; i < sortedPoints.length - 1; i++) {
            const point1 = sortedPoints[i];
            const point2 = sortedPoints[i + 1];
            
            if (rawValue >= point1.raw && rawValue <= point2.raw) {
                // Linear interpolation
                const ratio = (rawValue - point1.raw) / (point2.raw - point1.raw);
                return point1.calibrated + ratio * (point2.calibrated - point1.calibrated);
            }
        }
        
        // If outside range, extrapolate using last segment
        if (rawValue < sortedPoints[0].raw) {
            const point1 = sortedPoints[0];
            const point2 = sortedPoints[1];
            const ratio = (rawValue - point1.raw) / (point2.raw - point1.raw);
            return point1.calibrated + ratio * (point2.calibrated - point1.calibrated);
        } else {
            const point1 = sortedPoints[sortedPoints.length - 2];
            const point2 = sortedPoints[sortedPoints.length - 1];
            const ratio = (rawValue - point1.raw) / (point2.raw - point1.raw);
            return point1.calibrated + ratio * (point2.calibrated - point1.calibrated);
        }
    }

    // Convert raw weight based on unit and calibration
    function convertWeight(rawValue, fromUnit = 'mV') {
        let calibratedValue = rawValue;
        
        // Convert to mV if needed
        if (fromUnit === 'tons') {
            // Convert tons to mV (assuming 1 ton = 1000 mV)
            calibratedValue = rawValue * 1000;
        } else if (fromUnit === 'kg') {
            // Convert kg to mV (assuming 1 kg = 1 mV)
            calibratedValue = rawValue;
        }
        
        // Apply calibration
        calibratedValue = calculateCalibratedWeight(calibratedValue);
        
        // Convert to selected unit
        if (weightUnit === 'tons') {
            return calibratedValue / 1000;
        } else if (weightUnit === 'kg') {
            return calibratedValue;
        } else {
            return calibratedValue;
        }
    }

    // Calibrate all weight values for a device
    function calibrateDeviceWeight(deviceId) {
        const device = GPSDashboard.appState.devices[deviceId];
        if (!device || !device.data) return;
        
        device.data.forEach(point => {
            if (point.weightRaw !== undefined) {
                // Determine source unit (check metadata or infer)
                const sourceUnit = device.metadata.weightUnit || 'mV';
                point.weightCalibrated = convertWeight(point.weightRaw, sourceUnit);
            }
        });
        
        // Update sensor stats
        if (device.sensorStats && device.sensorStats.weightCalibrated) {
            const values = device.data.map(p => p.weightCalibrated).filter(v => v != null);
            if (values.length > 0) {
                device.sensorStats.weightCalibrated = {
                    min: Math.min(...values),
                    max: Math.max(...values),
                    avg: values.reduce((a, b) => a + b, 0) / values.length,
                    latest: values[values.length - 1],
                    unit: weightUnit,
                    values: values
                };
            }
        }
    }

    // Show calibration modal
    function showCalibrationModal() {
        renderCalibrationPoints();
        document.getElementById('calibrationModal').classList.add('active');
    }

    // Hide calibration modal
    function hideCalibrationModal() {
        document.getElementById('calibrationModal').classList.remove('active');
    }

    // Apply calibration to all devices
    function applyCalibration() {
        Object.keys(GPSDashboard.appState.devices).forEach(deviceId => {
            calibrateDeviceWeight(deviceId);
        });
        
        // Update all charts
        Object.keys(GPSDashboard.appState.devices).forEach(deviceId => {
            if (GPSDashboard.appState.expandedDevices.has(deviceId)) {
                GPSDashboard.renderDeviceChart(deviceId);
                GPSDashboard.updateSensorSummary(deviceId);
            }
        });
        
        hideCalibrationModal();
        alert('Calibration applied to all devices!');
    }

    // Update weight unit display
    function updateWeightUnitDisplay() {
        const unitSelect = document.getElementById('weightUnit');
        if (unitSelect) {
            unitSelect.value = weightUnit;
        }
    }

    return {
        init,
        setWeightUnit,
        addCalibrationPoint,
        updatePoint,
        removePoint,
        calculateCalibratedWeight,
        convertWeight,
        calibrateDeviceWeight,
        showCalibrationModal,
        hideCalibrationModal,
        applyCalibration,
        getCalibrationPoints: () => calibrationPoints,
        getWeightUnit: () => weightUnit
    };
})();