document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('uploadForm');
    
    if (form) {
        form.addEventListener('submit', function(e) {
            const statusElement = document.getElementById('statusMessage');
            const submitBtn = document.getElementById('submitBtn');
            const terminalElement = document.querySelector('.terminal');
            
            if (terminalElement) {
                terminalElement.innerHTML = '';
                terminalElement.style.display = 'none';
            }
            
            statusElement.classList.remove('hidden');
            statusElement.classList.remove('status-success', 'status-error', 'status-warning');
            statusElement.classList.add('status-loading');
            statusElement.textContent = '⏳ Google Sheets update in progress';
            
            submitBtn.disabled = true;
            submitBtn.textContent = 'Processing...';
            submitBtn.style.opacity = '0.7';
        });
    }
    
    // Mobile optimization
    optimizeForMobile();
    
    // Check operation status if operation_id exists
    const operationIdElement = document.querySelector('[data-operation-id]');
    if (operationIdElement) {
        const operationId = operationIdElement.getAttribute('data-operation-id');
        checkOperationStatus(operationId);
    }
});

function optimizeForMobile() {
    const terminal = document.querySelector('.terminal');
    const MOBILE_USER_AGENTS = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i;
    const isMobile = MOBILE_USER_AGENTS.test(navigator.userAgent);
    
    if (terminal && isMobile) {
        const viewportWidth = window.innerWidth;
        const viewportHeight = window.innerHeight;
        const headerHeight = document.querySelector('.header')?.offsetHeight || 100;
        const formHeight = document.querySelector('.form-container')?.offsetHeight || 150;
        const availableHeight = viewportHeight - headerHeight - formHeight - 50;
        
        terminal.style.width = '100%';
        terminal.style.maxWidth = '100%';
        terminal.style.maxHeight = Math.max(availableHeight, 300) + 'px';
        terminal.style.fontSize = viewportWidth < 400 ? '11px' : '12px';
        terminal.style.lineHeight = '1.3';
    }
}

function checkOperationStatus(operationId) {
    fetch('/status/' + operationId)
        .then(response => response.json())
        .then(data => {
            const statusElement = document.getElementById('statusMessage');
            if (statusElement) {
                statusElement.textContent = data.status;

                if (data.status.includes('✅')) {
                    statusElement.className = 'status status-success';
                } else if (data.status.includes('❌')) {
                    statusElement.className = 'status status-error';
                } else if (data.status.includes('⏳')) {
                    statusElement.className = 'status status-loading';
                    setTimeout(() => checkOperationStatus(operationId), 5000);
                } else if (data.status.includes('⚠️')) {
                    statusElement.className = 'status status-warning';
                }
            }
        })
        .catch(error => {
            console.error('Error checking status:', error);
        });
}

// Handle window resize
window.addEventListener('resize', function() {
    optimizeForMobile();
});