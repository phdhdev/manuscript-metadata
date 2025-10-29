Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('saveBtn').onclick = saveMetadata;
        document.getElementById('loadBtn').onclick = loadMetadata;
        document.getElementById('clearBtn').onclick = clearMetadata;
        
        // Check for selection when panel loads
        checkSelection();
        
        // Monitor selection changes
        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            onSelectionChanged
        );
    }
});

function onSelectionChanged() {
    checkSelection();
}

async function checkSelection() {
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            const tables = selection.tables;
            
            context.load(tables);
            await context.sync();
            
            if (tables.items.length > 0) {
                // User has selected something in a table
                document.getElementById('noSelectionWarning').style.display = 'none';
                document.getElementById('cellInfo').style.display = 'block';
                document.getElementById('cellLocation').textContent = 'Table cell';
            } else {
                // No table cell selected
                document.getElementById('noSelectionWarning').style.display = 'block';
                document.getElementById('cellInfo').style.display = 'none';
            }
        });
    } catch (error) {
        console.log('Error checking selection:', error);
    }
}

async function saveMetadata() {
    const linkUrl = document.getElementById('linkUrl').value;
    const references = document.getElementById('references').value;
    const altTags = document.getElementById('altTags').value;
    const functionality = document.getElementById('functionality').value;
    
    const metadata = {
        link: linkUrl,
        references: references,
        altTags: altTags,
        functionality: functionality,
        timestamp: new Date().toISOString()
    };
    
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            const tables = selection.tables;
            
            context.load(tables);
            await context.sync();
            
            if (tables.items.length === 0) {
                showStatus('Please select a table cell first.', 'error');
                return;
            }
            
            // Get the first cell in the selection
            const table = tables.items[0];
            const range = selection;
            
            // Store metadata as a custom XML part
            const metadataJson = JSON.stringify(metadata);
            
            // Use content control to mark the cell
            const contentControl = range.insertContentControl();
            contentControl.tag = 'cellMetadata';
            contentControl.title = 'Cell with Metadata';
            
            // Store the metadata in the content control's properties
            // We'll use a custom property approach
            contentControl.appearance = 'BoundingBox'; // Make it visible but subtle
            
            // Save metadata to document settings with a unique key based on content control ID
            await context.sync();
            
            // Get the content control ID
            context.load(contentControl, 'id');
            await context.sync();
            
            const key = `cellMetadata_${contentControl.id}`;
            
            // Save to document custom XML or settings
            Office.context.document.settings.set(key, metadataJson);
            await Office.context.document.settings.saveAsync();
            
            showStatus('Metadata saved successfully!', 'success');
        });
    } catch (error) {
        console.error('Error saving metadata:', error);
        showStatus('Error saving metadata: ' + error.message, 'error');
    }
}

async function loadMetadata() {
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            const contentControls = selection.contentControls;
            
            context.load(contentControls);
            await context.sync();
            
            // Look for a content control with cellMetadata tag
            let foundMetadata = false;
            
            for (let i = 0; i < contentControls.items.length; i++) {
                const cc = contentControls.items[i];
                context.load(cc, ['id', 'tag']);
                await context.sync();
                
                if (cc.tag === 'cellMetadata') {
                    const key = `cellMetadata_${cc.id}`;
                    const metadataJson = Office.context.document.settings.get(key);
                    
                    if (metadataJson) {
                        const metadata = JSON.parse(metadataJson);
                        
                        document.getElementById('linkUrl').value = metadata.link || '';
                        document.getElementById('references').value = metadata.references || '';
                        document.getElementById('altTags').value = metadata.altTags || '';
                        document.getElementById('functionality').value = metadata.functionality || '';
                        
                        showStatus('Metadata loaded successfully!', 'success');
                        foundMetadata = true;
                        break;
                    }
                }
            }
            
            if (!foundMetadata) {
                showStatus('No metadata found for this cell.', 'error');
                clearForm();
            }
        });
    } catch (error) {
        console.error('Error loading metadata:', error);
        showStatus('Error loading metadata: ' + error.message, 'error');
    }
}

async function clearMetadata() {
    if (!confirm('Are you sure you want to clear the metadata for this cell?')) {
        return;
    }
    
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            const contentControls = selection.contentControls;
            
            context.load(contentControls);
            await context.sync();
            
            let cleared = false;
            
            for (let i = 0; i < contentControls.items.length; i++) {
                const cc = contentControls.items[i];
                context.load(cc, ['id', 'tag']);
                await context.sync();
                
                if (cc.tag === 'cellMetadata') {
                    const key = `cellMetadata_${cc.id}`;
                    Office.context.document.settings.remove(key);
                    await Office.context.document.settings.saveAsync();
                    
                    // Remove the content control but keep the content
                    cc.delete(true); // true = keep content
                    
                    cleared = true;
                    break;
                }
            }
            
            await context.sync();
            
            if (cleared) {
                clearForm();
                showStatus('Metadata cleared successfully!', 'success');
            } else {
                showStatus('No metadata found to clear.', 'error');
            }
        });
    } catch (error) {
        console.error('Error clearing metadata:', error);
        showStatus('Error clearing metadata: ' + error.message, 'error');
    }
}

function clearForm() {
    document.getElementById('linkUrl').value = '';
    document.getElementById('references').value = '';
    document.getElementById('altTags').value = '';
    document.getElementById('functionality').value = '';
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
    
    setTimeout(() => {
        statusDiv.className = 'status';
        statusDiv.style.display = 'none';
    }, 5000);
}
