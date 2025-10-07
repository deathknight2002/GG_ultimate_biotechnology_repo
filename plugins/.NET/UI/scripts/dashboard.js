// SeleniumVBA Biotech Dashboard - Dynamic UI Controller
// Handles real-time updates and user interactions

class BiotechDashboard {
    constructor() {
        this.sampleData = [];
        this.updateInterval = null;
        this.init();
    }

    init() {
        console.log('Biotech Dashboard initialized');
        this.loadInitialData();
        this.startAutoRefresh();
        this.initializeChart();
    }

    loadInitialData() {
        // Simulate loading sample data
        this.sampleData = [
            { id: 'BT-001', type: 'DNA Sequencing', status: 'Active', progress: 75, results: 'Pending' },
            { id: 'BT-002', type: 'Protein Analysis', status: 'Complete', progress: 100, results: 'View' },
            { id: 'BT-003', type: 'Cell Culture', status: 'Processing', progress: 45, results: 'Pending' },
            { id: 'BT-004', type: 'Gene Expression', status: 'Active', progress: 60, results: 'Pending' },
            { id: 'BT-005', type: 'Microscopy', status: 'Complete', progress: 100, results: 'View' }
        ];
        
        this.updateMetrics();
        this.renderSampleTable();
    }

    updateMetrics() {
        const activeCount = this.sampleData.filter(s => s.status === 'Active').length;
        const processingCount = this.sampleData.filter(s => s.status === 'Processing').length;
        const totalProgress = this.sampleData.reduce((sum, s) => sum + s.progress, 0);
        const avgProgress = Math.round(totalProgress / this.sampleData.length);

        this.animateCounter('sampleCount', activeCount);
        this.animateCounter('analysisCount', processingCount);
        this.animateCounter('completionRate', avgProgress, '%');
    }

    animateCounter(elementId, targetValue, suffix = '') {
        const element = document.getElementById(elementId);
        if (!element) return;

        const startValue = parseInt(element.textContent) || 0;
        const duration = 1000;
        const startTime = performance.now();

        const animate = (currentTime) => {
            const elapsed = currentTime - startTime;
            const progress = Math.min(elapsed / duration, 1);
            
            const currentValue = Math.floor(startValue + (targetValue - startValue) * progress);
            element.textContent = currentValue + suffix;

            if (progress < 1) {
                requestAnimationFrame(animate);
            }
        };

        requestAnimationFrame(animate);
    }

    renderSampleTable() {
        const tbody = document.getElementById('sampleTableBody');
        if (!tbody) return;

        tbody.innerHTML = this.sampleData.map(sample => `
            <tr>
                <td>${sample.id}</td>
                <td>${sample.type}</td>
                <td>
                    <span class="status-indicator ${this.getStatusClass(sample.status)}"></span>
                    ${sample.status}
                </td>
                <td>
                    <div class="progress-bar">
                        <div class="progress-fill" style="width: ${sample.progress}%"></div>
                    </div>
                </td>
                <td>
                    ${sample.results === 'View' 
                        ? '<button class="btn-glass" style="padding: 6px 12px; font-size: 0.875rem;" onclick="viewResults(\'' + sample.id + '\')">View</button>'
                        : sample.results
                    }
                </td>
            </tr>
        `).join('');
    }

    getStatusClass(status) {
        const statusMap = {
            'Active': 'status-active',
            'Processing': 'status-processing',
            'Complete': 'status-active',
            'Error': 'status-error',
            'Warning': 'status-warning'
        };
        return statusMap[status] || 'status-active';
    }

    initializeChart() {
        // Placeholder for chart initialization
        // In a real implementation, use a library like Chart.js
        const canvas = document.getElementById('dataChart');
        if (!canvas) return;

        const ctx = canvas.getContext('2d');
        this.drawPlaceholderChart(ctx, canvas.width, canvas.height);
    }

    drawPlaceholderChart(ctx, width, height) {
        // Simple gradient background
        const gradient = ctx.createLinearGradient(0, 0, 0, height);
        gradient.addColorStop(0, 'rgba(0, 255, 136, 0.2)');
        gradient.addColorStop(1, 'rgba(0, 212, 255, 0.1)');
        
        ctx.fillStyle = gradient;
        ctx.fillRect(0, 0, width, height);

        // Draw grid
        ctx.strokeStyle = 'rgba(0, 255, 136, 0.1)';
        ctx.lineWidth = 1;
        
        for (let i = 0; i < 10; i++) {
            const y = (height / 10) * i;
            ctx.beginPath();
            ctx.moveTo(0, y);
            ctx.lineTo(width, y);
            ctx.stroke();
        }

        // Draw sample data line
        ctx.strokeStyle = 'rgba(0, 255, 136, 0.8)';
        ctx.lineWidth = 2;
        ctx.beginPath();
        
        const points = 50;
        for (let i = 0; i < points; i++) {
            const x = (width / points) * i;
            const y = height / 2 + Math.sin(i / 5) * (height / 4) + Math.random() * 20;
            
            if (i === 0) {
                ctx.moveTo(x, y);
            } else {
                ctx.lineTo(x, y);
            }
        }
        ctx.stroke();
    }

    startAutoRefresh() {
        this.updateInterval = setInterval(() => {
            this.simulateDataUpdate();
        }, 5000); // Update every 5 seconds
    }

    simulateDataUpdate() {
        // Simulate progress updates
        this.sampleData.forEach(sample => {
            if (sample.status === 'Active' && sample.progress < 100) {
                sample.progress = Math.min(sample.progress + Math.random() * 5, 100);
                if (sample.progress >= 100) {
                    sample.status = 'Complete';
                    sample.results = 'View';
                }
            }
        });

        this.updateMetrics();
        this.renderSampleTable();
    }

    stopAutoRefresh() {
        if (this.updateInterval) {
            clearInterval(this.updateInterval);
            this.updateInterval = null;
        }
    }
}

// Global functions for button handlers
function refreshData() {
    console.log('Refreshing data...');
    dashboard.loadInitialData();
}

function startAnalysis() {
    console.log('Starting new analysis...');
    const newSample = {
        id: `BT-${String(dashboard.sampleData.length + 1).padStart(3, '0')}`,
        type: 'New Analysis',
        status: 'Active',
        progress: 0,
        results: 'Pending'
    };
    dashboard.sampleData.push(newSample);
    dashboard.updateMetrics();
    dashboard.renderSampleTable();
}

function exportData() {
    console.log('Exporting data...');
    const dataStr = JSON.stringify(dashboard.sampleData, null, 2);
    const dataBlob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(dataBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'biotech-data-export.json';
    link.click();
}

function generateReport() {
    console.log('Generating report...');
    alert('Report generation initiated. This would integrate with the .NET plugin system.');
}

function viewProtocols() {
    console.log('Viewing protocols...');
    alert('Protocols viewer would open here with biotech-specific protocols.');
}

function viewResults(sampleId) {
    console.log(`Viewing results for ${sampleId}`);
    alert(`Results viewer for ${sampleId} would open here.`);
}

// Initialize dashboard on page load
let dashboard;
document.addEventListener('DOMContentLoaded', () => {
    dashboard = new BiotechDashboard();
});

// Cleanup on page unload
window.addEventListener('beforeunload', () => {
    if (dashboard) {
        dashboard.stopAutoRefresh();
    }
});
