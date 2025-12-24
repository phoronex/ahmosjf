// File Handler Module
const FileHandler = (function() {
    
    // Load JSON file from file input
    function loadJsonFile(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const jsonData = JSON.parse(e.target.result);
                GPSDashboard.processDeviceData(jsonData);
                showNotification('Data loaded successfully!', 'success');
                
                // Save file info for later saving
                GPSDashboard.appState.currentFileName = file.name;
                GPSDashboard.appState.fileContent = jsonData;
                
                // Update date range display
                if (jsonData.date) {
                    const dateRange = document.getElementById('dateRange');
                    if (dateRange) {
                        dateRange.innerHTML = `
                            <span>📅 From: ${jsonData.date.from || 'N/A'}</span>
                            <span>➡️</span>
                            <span>📅 To: ${jsonData.date.to || 'N/A'}</span>
                            <span>⏱️ Interval: ${jsonData.date.interval || 'N/A'}</span>
                        `;
                    }
                }
            } catch (error) {
                console.error('Error parsing JSON:', error);
                showNotification('Error parsing JSON file: ' + error.message, 'error');
            }
        };
        
        reader.readAsText(file);
    }
    
    // Save dashboard as HTML file with embedded data
    function saveDashboard() {
        if (Object.keys(GPSDashboard.appState.devices).length === 0) {
            showNotification('No data to save!', 'warning');
            return;
        }
        
        // Create a blob with the current state
        const state = {
            devices: GPSDashboard.appState.devices,
            selectedSensors: Array.from(GPSDashboard.appState.selectedSensors || []),
            expandedDevices: Array.from(GPSDashboard.appState.expandedDevices || []),
            calibrationPoints: WeightCalibration.getCalibrationPoints(),
            weightUnit: WeightCalibration.getWeightUnit(),
            timestamp: new Date().toISOString(),
            version: '1.0.0'
        };
        
        // Create HTML content with embedded data
        const htmlContent = generateEmbeddedHtml(state);
        
        // Create blob and download
        const blob = new Blob([htmlContent], { type: 'text/html' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        
        // Generate filename with timestamp
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const deviceCount = Object.keys(GPSDashboard.appState.devices).length;
        a.download = `gps-dashboard-${deviceCount}devices-${timestamp}.html`;
        
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        showNotification('Dashboard saved successfully!', 'success');
    }
    
    // Generate HTML with embedded data
    function generateEmbeddedHtml(state) {
        return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>GPS Dashboard - ${Object.keys(state.devices).length} Devices - ${new Date(state.timestamp).toLocaleDateString()}</title>
    <link rel="stylesheet" href="style.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-adapter-date-fns/dist/chartjs-adapter-date-fns.bundle.min.js"></script>
    <style>
        /* Embedded CSS from style.css */
        ${getCssContent()}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>🚚 GPS Devices Dashboard (Saved)</h1>
            <p class="subtitle">Saved on ${new Date(state.timestamp).toLocaleString()} | ${Object.keys(state.devices).length} devices</p>
        </header>
        <div class="devices-container" id="devicesContainer"></div>
    </div>
    
    <script>
        // Embedded JavaScript data
        const embeddedData = ${JSON.stringify(state, null, 2)};
        
        // Embedded app logic
        ${getJsContent()}
        
        // Initialize with embedded data
        window.addEventListener('DOMContentLoaded', function() {
            WeightCalibration.init(embeddedData.calibrationPoints);
            WeightCalibration.setWeightUnit(embeddedData.weightUnit);
            GPSDashboard.appState.devices = embeddedData.devices;
            GPSDashboard.appState.selectedSensors = new Map(Object.entries(embeddedData.selectedSensors));
            GPSDashboard.appState.expandedDevices = new Set(embeddedData.expandedDevices);
            GPSDashboard.renderDevices();
        });
    </script>
</body>
</html>`;
    }
    
    // Get CSS content (simplified version for embedded)
    function getCssContent() {
        // This would contain the essential CSS from style.css
        // For brevity, returning a minimal version
        return `
            body { font-family: sans-serif; margin: 20px; }
            .container { max-width: 1200px; margin: 0 auto; }
            .device-card { border: 1px solid #ddd; margin: 10px 0; padding: 15px; }
            .chart-container { height: 300px; }
        `;
    }
    
    // Get JS content (simplified version for embedded)
    function getJsContent() {
        // This would contain essential JavaScript
        // For brevity, returning a minimal version
        return `
            const GPSDashboard = {
                appState: {
                    devices: {},
                    selectedSensors: new Map(),
                    expandedDevices: new Set()
                },
                renderDevices: function() {
                    // Simplified render function
                    console.log('Rendering devices');
                }
            };
            
            const WeightCalibration = {
                init: function() {},
                setWeightUnit: function() {}
            };
        `;
    }
    
    // Show notification
    function showNotification(message, type = 'info') {
        // Create notification element
        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.innerHTML = `
            <span>${message}</span>
            <button onclick="this.parentElement.remove()">×</button>
        `;
        
        // Add styles if not exists
        if (!document.querySelector('#notification-styles')) {
            const styles = document.createElement('style');
            styles.id = 'notification-styles';
            styles.textContent = `
                .notification {
                    position: fixed;
                    top: 20px;
                    right: 20px;
                    padding: 15px 20px;
                    border-radius: 5px;
                    color: white;
                    z-index: 1000;
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    min-width: 300px;
                    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                    animation: slideIn 0.3s ease;
                }
                .notification-success { background: #28a745; }
                .notification-error { background: #dc3545; }
                .notification-warning { background: #ffc107; color: #333; }
                .notification-info { background: #17a2b8; }
                .notification button {
                    background: none;
                    border: none;
                    color: inherit;
                    font-size: 1.2em;
                    cursor: pointer;
                    margin-left: 10px;
                }
                @keyframes slideIn {
                    from { transform: translateX(100%); opacity: 0; }
                    to { transform: translateX(0); opacity: 1; }
                }
            `;
            document.head.appendChild(styles);
        }
        
        document.body.appendChild(notification);
        
        // Auto remove after 5 seconds
        setTimeout(() => {
            if (notification.parentElement) {
                notification.remove();
            }
        }, 5000);
    }
    
    // Export functions
    return {
        loadJsonFile,
        saveDashboard,
        showNotification
    };
})();