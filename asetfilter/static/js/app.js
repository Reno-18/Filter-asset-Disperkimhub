/**
 * AsetFilter - Main JavaScript
 * Handles UI interactions, sidebar toggle, and dynamic features
 */

document.addEventListener('DOMContentLoaded', function () {
    initSidebar();
    initTooltips();
    initAutoHideAlerts();
});

/**
 * Initialize sidebar toggle for mobile
 */
function initSidebar() {
    const sidebar = document.getElementById('sidebar');
    const toggleBtn = document.querySelector('.sidebar-toggle');

    if (toggleBtn && sidebar) {
        toggleBtn.addEventListener('click', function () {
            sidebar.classList.toggle('show');
        });

        // Close sidebar when clicking outside on mobile
        document.addEventListener('click', function (e) {
            if (window.innerWidth < 768) {
                if (!sidebar.contains(e.target) && !toggleBtn.contains(e.target)) {
                    sidebar.classList.remove('show');
                }
            }
        });
    }
}

/**
 * Initialize Bootstrap tooltips
 */
function initTooltips() {
    const tooltipTriggerList = document.querySelectorAll('[data-bs-toggle="tooltip"]');
    tooltipTriggerList.forEach(function (tooltipTriggerEl) {
        new bootstrap.Tooltip(tooltipTriggerEl);
    });
}

/**
 * Auto-hide alerts after 5 seconds
 */
function initAutoHideAlerts() {
    const alerts = document.querySelectorAll('.alert:not(.alert-permanent)');
    alerts.forEach(function (alert) {
        setTimeout(function () {
            const bsAlert = new bootstrap.Alert(alert);
            bsAlert.close();
        }, 5000);
    });
}

/**
 * Show loading overlay
 */
function showLoading() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) {
        overlay.classList.remove('d-none');
    }
}

/**
 * Hide loading overlay
 */
function hideLoading() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) {
        overlay.classList.add('d-none');
    }
}

/**
 * Format number with thousand separators
 * @param {number} num - Number to format
 * @returns {string} Formatted number string
 */
function formatNumber(num) {
    return new Intl.NumberFormat('id-ID').format(num);
}

/**
 * Format currency (Indonesian Rupiah)
 * @param {number} num - Number to format
 * @returns {string} Formatted currency string
 */
function formatCurrency(num) {
    return new Intl.NumberFormat('id-ID', {
        style: 'currency',
        currency: 'IDR',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(num);
}

/**
 * Debounce function for search input
 * @param {Function} func - Function to debounce
 * @param {number} wait - Wait time in milliseconds
 * @returns {Function} Debounced function
 */
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

/**
 * Export table data (triggered by export buttons)
 * @param {string} format - Export format ('csv' or 'excel')
 */
function exportData(format) {
    const currentUrl = new URL(window.location.href);
    const exportUrl = format === 'csv' ? '/export-csv' : '/export-excel';

    // Copy filter parameters
    const params = new URLSearchParams(currentUrl.search);

    window.location.href = exportUrl + '?' + params.toString();
}

/**
 * Confirm and clear all data
 */
function confirmClearData() {
    if (confirm('Apakah Anda yakin ingin menghapus semua data? Tindakan ini tidak dapat dibatalkan.')) {
        document.getElementById('clear-data-form').submit();
    }
}

/**
 * Update URL parameters without page reload
 * @param {string} key - Parameter key
 * @param {string} value - Parameter value
 */
function updateUrlParam(key, value) {
    const url = new URL(window.location.href);

    if (value) {
        url.searchParams.set(key, value);
    } else {
        url.searchParams.delete(key);
    }

    window.history.pushState({}, '', url.toString());
}

/**
 * Fetch and update stats in sidebar
 */
async function updateStats() {
    try {
        const response = await fetch('/api/stats');
        const data = await response.json();

        const totalRecords = document.getElementById('total-records');
        if (totalRecords) {
            totalRecords.textContent = formatNumber(data.total);
        }
    } catch (error) {
        console.error('Error fetching stats:', error);
    }
}

// Update stats on page load
updateStats();
