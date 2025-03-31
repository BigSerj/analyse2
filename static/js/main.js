// Add this code to your JavaScript file that manages warehouse processing

// Option 1: Use a more resilient approach with Web Workers
function processWarehousesInBackground(warehouseIds) {
    // Create a simple worker that can run in background
    const workerCode = `
        self.onmessage = function(e) {
            const warehouseIds = e.data;
            let currentIndex = 0;
            
            function processNextWarehouse() {
                if (currentIndex >= warehouseIds.length) {
                    self.postMessage({type: 'complete'});
                    return;
                }
                
                const warehouseId = warehouseIds[currentIndex++];
                self.postMessage({type: 'processing', warehouseId: warehouseId});
                
                fetch('/process', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/x-www-form-urlencoded'},
                    body: 'store_id=' + warehouseId + '&other_params_here'
                })
                .then(response => response.json())
                .then(data => {
                    self.postMessage({type: 'warehouseComplete', warehouseId: warehouseId});
                    // Process next warehouse after a short delay
                    setTimeout(processNextWarehouse, 1000);
                })
                .catch(error => {
                    self.postMessage({type: 'error', warehouseId: warehouseId, error: error});
                    // Still try next warehouse
                    setTimeout(processNextWarehouse, 1000);
                });
            }
            
            processNextWarehouse();
        };
    `;
    
    // Create a blob and worker from the code
    const blob = new Blob([workerCode], {type: 'application/javascript'});
    const worker = new Worker(URL.createObjectURL(blob));
    
    // Handle messages from worker
    worker.onmessage = function(e) {
        const message = e.data;
        switch(message.type) {
            case 'processing':
                console.log('Started processing warehouse:', message.warehouseId);
                break;
            case 'warehouseComplete':
                console.log('Finished processing warehouse:', message.warehouseId);
                break;
            case 'complete':
                console.log('All warehouses processed');
                worker.terminate();
                break;
            case 'error':
                console.error('Error processing warehouse:', message.warehouseId, message.error);
                break;
        }
    };
    
    // Start processing
    worker.postMessage(warehouseIds);
}

// Option 2: Server-side batch processing
function processWarehousesBatch(warehouseIds) {
    // Send all warehouse IDs to the server at once
    fetch('/process-multiple', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            store_ids: warehouseIds,
            // Other parameters
            start_date: document.getElementById('start_date').value,
            end_date: document.getElementById('end_date').value,
            planning_days: document.getElementById('planning_days').value,
            search_days: document.getElementById('search_days').value
        })
    })
    .then(response => response.json())
    .then(data => {
        console.log('Batch processing complete', data);
    })
    .catch(error => {
        console.error('Error in batch processing', error);
    });
} 