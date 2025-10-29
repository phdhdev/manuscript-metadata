Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById('saveBtn').onclick = saveMetadata;
        document.getElementById('clearBtn').onclick = clearMetadata;
        
        // Check for selection when panel loads
        checkSelection();
        // Auto-load metadata if any exists
        loadMetadata();
        
        // Monitor selection changes
        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            onSelectionChanged
        );
    }
});

function onSelectionChanged() {
    checkSelection();
    // Automatically try to load metadata when selection changes
    loadMetadata();
}

async function checkSelection() {
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            
            context.load(selection, 'text');
            await context.sync();
            
            if (selection.text && selection.text.trim() !== '') {
                // User has selected some text
                document.getElementById('noSelectionWarning').style.display = 'none';
                document.getElementById('cellInfo').style.display = 'block';
                
                // Show a preview of the selected text (first 50 chars)
                const preview = selection.text.length > 50 
                    ? selection.text.substring(0, 47) + '...' 
                    : selection.text;
                document.getElementById('cellLocation').textContent = `"${preview}"`;
            } else {
                // No text selected
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
            const contentControls = selection.contentControls;
            
            // Check if there's actually text selected
            context.load(selection, 'text');
            context.load(contentControls);
            await context.sync();
            
            if (!selection.text || selection.text.trim() === '') {
                showStatus('Please select some text in a cell first.', 'error');
                return;
            }
            
            const metadataJson = JSON.stringify(metadata);
            let contentControl = null;
            let isExisting = false;
            
            // Check if there's already a content control (editing existing metadata)
            for (let i = 0; i < contentControls.items.length; i++) {
                const cc = contentControls.items[i];
                context.load(cc, ['id', 'tag']);
                await context.sync();
                
                if (cc.tag === 'cellMetadata') {
                    contentControl = cc;
                    isExisting = true;
                    break;
                }
            }
            
            // If no existing content control, create a new one
            if (!contentControl) {
                contentControl = selection.insertContentControl();
                contentControl.tag = 'cellMetadata';
                contentControl.title = 'View Metadata';
                contentControl.appearance = 'BoundingBox';
                await context.sync();
            }
            
            // Get or load the content control ID
            context.load(contentControl, 'id');
            await context.sync();
            
            const key = `cellMetadata_${contentControl.id}`;
            
            // Save to document settings
            Office.context.document.settings.set(key, metadataJson);
            
            // Save settings using proper callback
            await new Promise((resolve, reject) => {
                Office.context.document.settings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve();
                    } else {
                        reject(result.error);
                    }
                });
            });
            
            if (isExisting) {
                showStatus('Metadata updated successfully!', 'success');
            } else {
                showStatus('Metadata saved successfully!', 'success');
            }
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
                        
                        foundMetadata = true;
                        break;
                    }
                }
            }
            
            if (!foundMetadata) {
                // Silently clear the form - no error message needed
                clearForm();
            }
        });
    } catch (error) {
        console.error('Error loading metadata:', error);
        // Silently fail - don't show error to user
        clearForm();
    }
}

async function clearMetadata() {
    if (!confirm('Are you sure you want to clear the metadata for this text?')) {
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
                    
                    // Save settings synchronously using callback
                    await new Promise((resolve, reject) => {
                        Office.context.document.settings.saveAsync((result) => {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                resolve();
                            } else {
                                reject(result.error);
                            }
                        });
                    });
                    
                    // Remove the content control but keep the text (true = keep content)
                    cc.delete(true);
                    
                    cleared = true;
                    break;
                }
            }
            
            await context.sync();
            
            if (cleared) {
                clearForm();
                showStatus('Metadata cleared successfully!', 'success');
            } else {
                // If no metadata found, just clear the form anyway
                clearForm();
                showStatus('Form cleared.', 'success');
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
