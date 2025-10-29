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
            
            context.load(selection, ['text', 'inlinePictures']);
            await context.sync();
            
            const hasText = selection.text && selection.text.trim() !== '';
            const hasImages = selection.inlinePictures && selection.inlinePictures.items.length > 0;
            
            if (hasText || hasImages) {
                // User has selected text or images
                document.getElementById('noSelectionWarning').style.display = 'none';
                document.getElementById('cellInfo').style.display = 'block';
                
                if (hasImages && !hasText) {
                    document.getElementById('cellLocation').textContent = `Image selected`;
                } else if (hasText) {
                    // Show a preview of the selected text (first 50 chars)
                    const preview = selection.text.length > 50 
                        ? selection.text.substring(0, 47) + '...' 
                        : selection.text;
                    document.getElementById('cellLocation').textContent = `"${preview}"`;
                }
            } else {
                // Nothing selected
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
            
            // Load selection properties
            context.load(selection, ['text', 'inlinePictures']);
            context.load(contentControls);
            await context.sync();
            
            // Check if there's text or images selected
            const hasText = selection.text && selection.text.trim() !== '';
            const hasImages = selection.inlinePictures && selection.inlinePictures.items.length > 0;
            
            if (!hasText && !hasImages) {
                showStatus('Please select text or an image in a cell first.', 'error');
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
                contentControl.color = '#1B9FFF'; // Bright blue background
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
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            
            // Get all content controls that intersect with the selection
            const contentControls = selection.contentControls;
            contentControls.load(['items', 'tag', 'id', 'text']);
            
            await context.sync();
            
            showStatus(`Found ${contentControls.items.length} content controls in selection`, 'success');
            
            if (contentControls.items.length === 0) {
                // Try getting content controls from the entire document
                const allContentControls = context.document.contentControls;
                allContentControls.load(['items', 'tag', 'id']);
                await context.sync();
                
                showStatus(`Total content controls in document: ${allContentControls.items.length}`, 'success');
                
                let foundMetadataControl = false;
                
                // Check each content control to see if it's a metadata control
                for (let i = 0; i < allContentControls.items.length; i++) {
                    const cc = allContentControls.items[i];
                    if (cc.tag === 'cellMetadata') {
                        foundMetadataControl = true;
                        
                        const key = `cellMetadata_${cc.id}`;
                        Office.context.document.settings.remove(key);
                        
                        await new Promise((resolve, reject) => {
                            Office.context.document.settings.saveAsync((result) => {
                                if (result.status === Office.AsyncResultStatus.Succeeded) {
                                    resolve();
                                } else {
                                    reject(result.error);
                                }
                            });
                        });
                        
                        cc.delete(true); // Keep the text
                        await context.sync();
                        
                        clearForm();
                        showStatus('Metadata cleared successfully!', 'success');
                        return;
                    }
                }
                
                if (!foundMetadataControl) {
                    showStatus('No metadata controls found in document', 'error');
                }
                
                clearForm();
                return;
            }
            
            // Found content controls in selection
            let cleared = false;
            let foundCellMetadata = false;
            
            for (let i = 0; i < contentControls.items.length; i++) {
                const cc = contentControls.items[i];
                
                showStatus(`Checking control with tag: ${cc.tag}`, 'success');
                
                if (cc.tag === 'cellMetadata') {
                    foundCellMetadata = true;
                    
                    const key = `cellMetadata_${cc.id}`;
                    
                    Office.context.document.settings.remove(key);
                    
                    // Save settings
                    await new Promise((resolve, reject) => {
                        Office.context.document.settings.saveAsync((result) => {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                resolve();
                            } else {
                                reject(result.error);
                            }
                        });
                    });
                    
                    // Delete the content control but keep the text
                    cc.delete(true);
                    await context.sync();
                    
                    cleared = true;
                    break;
                }
            }
            
            if (!foundCellMetadata) {
                showStatus('No cellMetadata tag found in selected controls', 'error');
            }
            
            await context.sync();
            
            if (cleared) {
                clearForm();
                showStatus('Metadata cleared successfully!', 'success');
            } else {
                clearForm();
                showStatus('Form cleared.', 'success');
            }
        });
    } catch (error) {
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
    
    // Clear any existing timeout first
    if (window.statusTimeout) {
        clearTimeout(window.statusTimeout);
    }
    
    // Reset and show the message
    statusDiv.style.display = 'none'; // Hide first
    statusDiv.className = 'status'; // Reset class
    
    // Force a reflow to ensure the animation/display triggers
    void statusDiv.offsetHeight;
    
    // Now set the new message
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
    statusDiv.style.display = 'block';
    
    // Set timeout to hide after 5 seconds
    window.statusTimeout = setTimeout(() => {
        statusDiv.className = 'status';
        statusDiv.style.display = 'none';
    }, 5000);
}
