/* global Office, Excel, $ */

let treeInitialized = false;
let eventHandlersRegistered = false;
let saveTimeout = null;
let copiedNodes = [];

// Message bar helper functions
function showMessage(message, type = 'info') {
    const messageBar = document.getElementById('message-bar');
    const messageText = document.getElementById('message-text');
    
    messageText.textContent = message;
    messageBar.className = 'message-bar ' + type;
    messageBar.style.display = 'flex';
    
    // Auto-hide after 5 seconds for success messages
    if (type === 'success') {
        setTimeout(() => {
            messageBar.style.display = 'none';
        }, 5000);
    }
}

function hideMessage() {
    document.getElementById('message-bar').style.display = 'none';
}

// Custom prompt function
function showPrompt(message, defaultValue = '', callback) {
    const promptDiv = document.createElement('div');
    promptDiv.className = 'custom-prompt-overlay';
    promptDiv.innerHTML = `
        <div class="custom-prompt">
            <p>${message}</p>
            <input type="text" id="prompt-input" value="${defaultValue}" />
            <div class="prompt-buttons">
                <button id="prompt-ok" class="control-btn">OK</button>
                <button id="prompt-cancel" class="control-btn" style="background-color: #8a8886;">Cancel</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(promptDiv);
    
    const input = document.getElementById('prompt-input');
    input.focus();
    input.select();
    
    document.getElementById('prompt-ok').onclick = () => {
        const value = input.value;
        document.body.removeChild(promptDiv);
        callback(value);
    };
    
    document.getElementById('prompt-cancel').onclick = () => {
        document.body.removeChild(promptDiv);
        callback(null);
    };
    
    input.onkeypress = (e) => {
        if (e.key === 'Enter') {
            document.getElementById('prompt-ok').click();
        }
    };
}

// Custom confirm function
function showConfirm(message, callback) {
    const confirmDiv = document.createElement('div');
    confirmDiv.className = 'custom-prompt-overlay';
    confirmDiv.innerHTML = `
        <div class="custom-prompt">
            <p>${message}</p>
            <div class="prompt-buttons">
                <button id="confirm-yes" class="control-btn">Yes</button>
                <button id="confirm-no" class="control-btn" style="background-color: #8a8886;">No</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(confirmDiv);
    
    document.getElementById('confirm-yes').onclick = () => {
        document.body.removeChild(confirmDiv);
        callback(true);
    };
    
    document.getElementById('confirm-no').onclick = () => {
        document.body.removeChild(confirmDiv);
        callback(false);
    };
}


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log('Office.js ready');
    
    // Button event listeners
    document.getElementById("add-folder").onclick = addFolder;
    document.getElementById("add-sheet").onclick = addSheet;
    document.getElementById("refresh-tree").onclick = refreshTree;
    document.getElementById("save-structure").onclick = saveStructure;
    document.getElementById("message-close").onclick = hideMessage;
    document.getElementById("hide-others").onclick = hideOtherSheets;
    
    // Initialize the tree
    initializeTree();
    
    // Load saved structure or create default
    setTimeout(() => {
      loadStructure();
    }, 500);
  }
});


// Initialize jsTree with drag & drop
function initializeTree() {
  console.log('Initializing tree...');
  
  $('#sheet-tree').jstree({
    'core': {
      'data': [],
      'check_callback': true,
      'themes': {
        'name': 'default',          
        'responsive': false,
        'dots': false,
        'icons': true,               
        'stripes': false         
      }
    },
    'plugins': ['dnd', 'contextmenu', 'types'],
    'types': {
      'folder': {
        'icon': 'jstree-folder'
      },
      'sheet': {
        'icon': 'jstree-file'
      }
    },
    'dnd': {
      'is_draggable': function(nodes) {
        return true;
      },
      'check_while_dragging': true,
      'large_drop_target': true,
      'large_drag_target': true,
      'always_copy': false,
      'inside_pos': 'last',
      'touch': true,
      'drag_selection': true,
      'copy': false
    },
    'contextmenu': {
      'items': customContextMenu
    }
  }).on('ready.jstree', function() {
    treeInitialized = true;
    console.log('Tree initialized and ready');
    
  });

  // Handle node click
  $('#sheet-tree').on('select_node.jstree', function (e, data) {
    if (data.node.data && data.node.data.isWorksheet && data.node.data.sheetName) {
      navigateToSheet(data.node.data.sheetName);
    }
  });

  // Handle drag and drop
  $('#sheet-tree').on('move_node.jstree', function (e, data) {
    console.log('Node moved, saving structure...');
    
    // Clear any pending save
    if (saveTimeout) {
      clearTimeout(saveTimeout);
    }
    
    // Schedule a new save (debounced - only saves once after all moves complete)
    saveTimeout = setTimeout(() => {
      saveStructureToStorage()
        .catch(err => {
          console.error('Save failed after move:', err);
          showMessage('Moved but save failed: ' + err.message, 'warning');
        });
      saveTimeout = null;
    }, 200);
  });

  // Handle rename
  $('#sheet-tree').on('rename_node.jstree', function (e, data) {
    console.log('Node renamed');
    if (data.node.data && data.node.data.isWorksheet) {
      const oldName = data.node.data.sheetName;
      const newName = data.text;
      if (oldName !== newName) {
        // Check if new name already exists
        checkSheetNameExists(newName).then(exists => {
          if (exists) {
            // Revert the rename in the tree
            const tree = $('#sheet-tree').jstree(true);
            tree.rename_node(data.node, oldName);
            showMessage(`A sheet named "${newName}" already exists. Please choose a different name.`, 'error');
          } else {
            // Name is available, proceed with rename
            renameSheet(oldName, newName).then(() => {
              data.node.data.sheetName = newName;
              saveStructureToStorage();
              showMessage(`Sheet renamed to "${newName}" successfully!`, 'success');
            }).catch(err => {
              console.error('Rename failed:', err);
              // Revert on error
              const tree = $('#sheet-tree').jstree(true);
              tree.rename_node(data.node, oldName);
              showMessage('Error renaming sheet: ' + err.message, 'error');
            });
          }
        });
      }
    } else {
      // For folders, just save (no Excel validation needed)
      saveStructureToStorage();
    }
  });

  // Handle clicking on empty space to deselect
  $('#sheet-tree').on('click', function(e) {
    // Check if the click was directly on the container (empty space)
    // and not on a node or its children
    if (e.target.id === 'sheet-tree' || $(e.target).hasClass('jstree-container-ul')) {
      const tree = $('#sheet-tree').jstree(true);
      tree.deselect_all();
      console.log('Clicked empty space - deselected all nodes');
    }
  });

  // Handle Delete key press
  $(document).on('keydown', function(e) {
    // Check if Delete key was pressed
    if (e.key === 'Delete' || e.keyCode === 46) {
      const tree = $('#sheet-tree').jstree(true);
      const selected = tree.get_selected(true); // Get selected nodes as objects
      
      if (selected.length === 0) {
        return; // Nothing selected
      }
      
      // Separate sheets and folders
      const sheets = selected.filter(node => node.data && node.data.isWorksheet);
      const folders = selected.filter(node => !node.data || !node.data.isWorksheet);
      
      // Build confirmation message
      let message = 'Delete the following items?\n\n';
      if (sheets.length > 0) {
        message += `Sheets (${sheets.length}):\n`;
        sheets.forEach(node => message += `  • ${node.text}\n`);
      }
      if (folders.length > 0) {
        message += `\nFolders (${folders.length}):\n`;
        folders.forEach(node => message += `  • ${node.text}\n`);
      }
      
      showConfirm(message, async (confirmed) => {
        if (confirmed) {
          try {
            // Delete all sheets from Excel first
            for (const node of sheets) {
              try {
                await deleteSheet(node.data.sheetName);
                console.log(`Deleted sheet: ${node.text}`);
              } catch (error) {
                console.error(`Failed to delete sheet ${node.text}:`, error);
                showMessage(`Error deleting sheet "${node.text}": ${error.message}`, 'error');
              }
            }
            
            // Delete all nodes from tree (both sheets and folders)
            for (const node of selected) {
              tree.delete_node(node);
            }
            
            // Save structure
            await saveStructureToStorage();
            
            // Show success message
            const totalDeleted = selected.length;
            showMessage(`Successfully deleted ${totalDeleted} item(s)`, 'success');
            
          } catch (error) {
            console.error('Delete operation failed:', error);
            showMessage('Error during deletion: ' + error.message, 'error');
          }
        }
      });
    }
  });

  // Handle Copy/Paste keyboard shortcuts
  $(document).on('keydown', function(e) {
    // Ctrl+C or Cmd+C (Mac)
    if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
      e.preventDefault();
      copyNodes();
    }
    
    // Ctrl+V or Cmd+V (Mac)
    if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
      e.preventDefault();
      pasteNodes();
    }
  });

}

// Add a new folder
function addFolder() {
  console.log('Add folder clicked');
  
  if (!treeInitialized) {
    showMessage('Tree is still loading, please wait...', 'warning');
    return;
  }

  try {
    const tree = $('#sheet-tree').jstree(true);
    if (!tree) {
      console.error('Tree instance not found');
      showMessage('Tree not ready', 'error');
      return;
    }

    const selected = tree.get_selected();
    const parent = selected.length > 0 ? selected[0] : '#';

    if (parent === '#') {
      console.log('Adding folder to root (no selection)');
    } else {
      console.log('Adding folder under selected node:', parent);
    }
    
    showPrompt('Enter folder name:', 'New Folder', (folderName) => {
      if (folderName && folderName.trim() !== '') {
        console.log('Creating folder:', folderName);
        
        const nodeId = 'folder_' + Date.now();
        const newNode = tree.create_node(parent, {
          id: nodeId,
          text: folderName,
          type: 'folder',
          icon: 'jstree-folder',
          data: {
            isWorksheet: false,
            nodeType: 'folder'
          }
        }, 'last');
        
        console.log('Node created:', newNode);
        
        if (newNode) {
          tree.open_node(parent);
          console.log('Saving structure...');
          saveStructureToStorage()
            .then(() => showMessage('Folder created successfully!', 'success'))
            .catch(err => {
              console.error('Save failed:', err);
              showMessage('Folder created but could not save structure: ' + err.message, 'error');
            });
        } else {
          console.error('Failed to create node');
          showMessage('Failed to create folder node', 'error');
        }
      }
    });
  } catch (error) {
    console.error('Error in addFolder:', error);
    showMessage('Error adding folder: ' + error.message, 'error');
  }
}


// Add a new sheet
async function addSheet() {
  console.log('Add sheet clicked');
  
  if (!treeInitialized) {
    showMessage('Tree is still loading, please wait...', 'warning');
    return;
  }

  const tree = $('#sheet-tree').jstree(true);
  if (!tree) {
    showMessage('Tree not ready', 'error');
    return;
  }

  const selected = tree.get_selected();
  const parent = selected.length > 0 ? selected[0] : '#';
  
  showPrompt('Enter sheet name:', 'New Sheet', async (sheetName) => {
    if (sheetName && sheetName.trim() !== '') {
      try {
        console.log('Creating sheet:', sheetName);
        
        await Excel.run(async (context) => {
          // Get all existing sheet names
          const sheets = context.workbook.worksheets;
          sheets.load("items/name");
          await context.sync();
          
          const existingNames = sheets.items.map(s => s.name.toLowerCase());
          
          // Check if the name exists and find next available name
          let finalName = sheetName;
          let counter = 2;
          
          while (existingNames.includes(finalName.toLowerCase())) {
            finalName = `${sheetName} (${counter})`;
            counter++;
          }
          
          // Log if name was changed
          if (finalName !== sheetName) {
            console.log(`Sheet name "${sheetName}" already exists, using "${finalName}" instead`);
          }
          
          // Create the sheet with the final name
          const sheet = context.workbook.worksheets.add(finalName);
          sheet.activate();
          await context.sync();
          
          console.log('Sheet created in Excel:', finalName);
          
          const nodeId = 'sheet_' + Date.now();
          const newNode = tree.create_node(parent, {
            id: nodeId,
            text: finalName,
            type: 'sheet',
            icon: 'jstree-file',
            data: {
              isWorksheet: true,
              sheetName: finalName,
              nodeType: 'sheet'
            }
          }, 'last');
          
          if (newNode) {
            tree.open_node(parent);
            await saveStructureToStorage();
            
            // Show appropriate message
            if (finalName !== sheetName) {
              showMessage(`Sheet "${sheetName}" already exists. Created "${finalName}" instead.`, 'success');
            } else {
              showMessage('Sheet created successfully!', 'success');
            }
          }
        });
      } catch (error) {
        console.error("Error creating sheet:", error);
        showMessage("Error creating sheet: " + error.message, 'error');
      }
    }
  });
}


// Refresh tree from Excel sheets
async function refreshTree() {
  if (!treeInitialized) {
    showMessage('Tree is still loading, please wait...', 'warning');
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name, items/visibility");
      await context.sync();

      const tree = $('#sheet-tree').jstree(true);
      const currentNodes = tree.get_json('#', { flat: true });
      
      const existingSheetNames = currentNodes
        .filter(node => node.data && node.data.isWorksheet)
        .map(node => node.data.sheetName);
      
      const excelSheetNames = sheets.items.map(sheet => sheet.name);
      
      // Add missing sheets to root
      excelSheetNames.forEach(sheetName => {
        if (!existingSheetNames.includes(sheetName)) {
          const nodeId = 'sheet_' + Date.now() + '_' + Math.random();
          tree.create_node('#', {
            id: nodeId,
            text: sheetName,
            type: 'sheet',
            icon: 'jstree-file',
            data: {
              isWorksheet: true,
              sheetName: sheetName,
              nodeType: 'sheet'
            }
          }, 'last');
        }
      });
      
      // Remove nodes for deleted sheets
      currentNodes.forEach(node => {
        if (node.data && node.data.isWorksheet && !excelSheetNames.includes(node.data.sheetName)) {
          tree.delete_node(node.id);
        }
      });
      
      await saveStructureToStorage();
      showMessage('Tree refreshed successfully!', 'success');
    });
  } catch (error) {
    console.error("Error refreshing tree:", error);
    showMessage("Error refreshing: " + error.message, 'error');
  }
}

// Navigate to a specific sheet
async function navigateToSheet(sheetName) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      sheet.load("visibility");
      await context.sync();
      
      // If sheet is hidden, unhide it first
      if (sheet.visibility === Excel.SheetVisibility.hidden) {
        console.log(`Sheet "${sheetName}" is hidden, unhiding...`);
        sheet.visibility = Excel.SheetVisibility.visible;
        await context.sync();
      }
      
      // Now activate/select the sheet
      sheet.activate();
      await context.sync();
      
      console.log(`Navigated to sheet: ${sheetName}`);
    });
  } catch (error) {
    console.error("Error navigating to sheet:", error);
    showMessage("Sheet '" + sheetName + "' not found. It may have been deleted.", 'error');
  }
}

// Rename a sheet
async function renameSheet(oldName, newName) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(oldName);
      sheet.name = newName;
      await context.sync();
    });
  } catch (error) {
    console.error("Error renaming sheet:", error);
    throw error;
  }
}

// Custom context menu
function customContextMenu(node) {
  const tree = $('#sheet-tree').jstree(true);
  
  const menu = {
    "addFolder": {
      "label": "Add Folder Here",
      "action": function () {
        showPrompt('Enter folder name:', 'New Folder', (folderName) => {
          if (folderName && folderName.trim() !== '') {
            const nodeId = 'folder_' + Date.now();
            tree.create_node(node.id, {
              id: nodeId,
              text: folderName,
              type: 'folder',
              icon: 'jstree-folder',
              data: {
                isWorksheet: false,
                nodeType: 'folder'
              }
            }, 'last');
            tree.open_node(node.id);
            saveStructureToStorage().catch(err => console.error('Save failed:', err));
          }
        });
      }
    },
    "addSheet": {
      "label": "Add Sheet Here",
      "action": async function () {
        showPrompt('Enter sheet name:', 'New Sheet', async (sheetName) => {
          if (sheetName && sheetName.trim() !== '') {
            try {
              await Excel.run(async (context) => {
                // Get all existing sheet names
                const sheets = context.workbook.worksheets;
                sheets.load("items/name");
                await context.sync();
                
                const existingNames = sheets.items.map(s => s.name.toLowerCase());
                
                // Check if the name exists and find next available name
                let finalName = sheetName;
                let counter = 2;
                
                while (existingNames.includes(finalName.toLowerCase())) {
                  finalName = `${sheetName} (${counter})`;
                  counter++;
                }
                
                // Log if name was changed
                if (finalName !== sheetName) {
                  console.log(`Sheet name "${sheetName}" already exists, using "${finalName}" instead`);
                }
                
                // Create the sheet with the final name
                const sheet = context.workbook.worksheets.add(finalName);
                sheet.activate();
                await context.sync();
                
                const nodeId = 'sheet_' + Date.now();
                tree.create_node(node.id, {
                  id: nodeId,
                  text: finalName,
                  type: 'sheet',
                  icon: 'jstree-file',
                  data: {
                    isWorksheet: true,
                    sheetName: finalName,
                    nodeType: 'sheet'
                  }
                }, 'last');
                tree.open_node(node.id);
                await saveStructureToStorage();
                
                // Show appropriate message
                if (finalName !== sheetName) {
                  showMessage(`Sheet "${sheetName}" already exists. Created "${finalName}" instead.`, 'success');
                } else {
                  showMessage('Sheet created successfully!', 'success');
                }
              });
            } catch (error) {
              console.error("Error creating sheet:", error);
              showMessage("Error creating sheet: " + error.message, 'error');
            }
          }
        });
      }
    },
    "copy": {
      "label": "Copy",
      "action": function () {
        // Select this node if not already selected
        const tree = $('#sheet-tree').jstree(true);
        if (!tree.is_selected(node)) {
          tree.deselect_all();
          tree.select_node(node);
        }
        copyNodes();
      }
    },
    "paste": {
      "label": "Paste",
      "action": function () {
        // Select this node as the paste target
        const tree = $('#sheet-tree').jstree(true);
        tree.deselect_all();
        tree.select_node(node);
        pasteNodes();
      },
      "_disabled": function() {
        // Disable if nothing is copied
        return copiedNodes.length === 0;
      }
    },
    "moveToRoot": {
      "label": "Move to Root",
      "action": function () {
        try {
          // Get fresh reference to the node
          const nodeToMove = tree.get_node(node.id);
          
          if (!nodeToMove) {
            showMessage('Node not found', 'error');
            return;
          }
          
          // Check if already at root
          if (nodeToMove.parent === '#') {
            showMessage(`"${node.text}" is already at root level`, 'info');
            return;
          }
          
          // Move to root
          const moved = tree.move_node(nodeToMove, '#', 'last');
          
          if (moved) {
            // Add a small delay before saving to let jsTree update
            setTimeout(() => {
              saveStructureToStorage()
                .then(() => {
                  showMessage(`"${node.text}" moved to root level`, 'success');
                })
                .catch(err => {
                  console.error('Save failed:', err);
                  showMessage('Moved but save failed: ' + err.message, 'warning');
                });
            }, 100);
          } else {
            showMessage('Failed to move node', 'error');
          }
        } catch (error) {
          console.error('Error moving to root:', error);
          showMessage('Error moving node: ' + error.message, 'error');
        }
      }
    },
    "rename": {
      "label": "Rename",
      "action": function () {
        tree.edit(node);
      }
    }
  };
  
  // Add hide option (only for sheets)
  if (node.data && node.data.isWorksheet) {
    menu.hide = {
      "label": "Hide",
      "action": async function () {
        try {
          await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(node.data.sheetName);
            sheet.visibility = Excel.SheetVisibility.hidden;
            await context.sync();
            
            showMessage(`Sheet "${node.text}" hidden successfully!`, 'success');
          });
        } catch (error) {
          console.error("Error hiding sheet:", error);
          
          // Check if it's the "can't hide all sheets" error
          if (error.message && error.message.includes("InvalidOperation")) {
            showMessage(`Cannot hide "${node.text}" - at least one sheet must remain visible in Excel.`, 'warning');
          } else {
            showMessage("Error hiding sheet: " + error.message, 'error');
          }
        }
      }
    };
  }
  
  // Add delete option
  menu.delete = {
    "label": "Delete",
    "action": async function () {
      // Check if multiple nodes are selected
      const selected = tree.get_selected(true);
      
      if (selected.length > 1) {
        // Multiple items selected - delete all
        const sheets = selected.filter(n => n.data && n.data.isWorksheet);
        const folders = selected.filter(n => !n.data || !n.data.isWorksheet);
        
        // Build confirmation message
        let message = 'Delete the following items?\n\n';
        if (sheets.length > 0) {
          message += `Sheets (${sheets.length}):\n`;
          sheets.forEach(n => message += `  • ${n.text}\n`);
        }
        if (folders.length > 0) {
          message += `\nFolders (${folders.length}):\n`;
          folders.forEach(n => message += `  • ${n.text}\n`);
        }
        
        showConfirm(message, async (confirmed) => {
          if (confirmed) {
            try {
              // Delete all sheets from Excel first
              for (const n of sheets) {
                try {
                  await deleteSheet(n.data.sheetName);
                  console.log(`Deleted sheet: ${n.text}`);
                } catch (error) {
                  console.error(`Failed to delete sheet ${n.text}:`, error);
                  showMessage(`Error deleting sheet "${n.text}": ${error.message}`, 'error');
                }
              }
              
              // Delete all nodes from tree
              for (const n of selected) {
                tree.delete_node(n);
              }
              
              // Save structure
              await saveStructureToStorage();
              
              // Show success message
              showMessage(`Successfully deleted ${selected.length} item(s)`, 'success');
              
            } catch (error) {
              console.error('Delete operation failed:', error);
              showMessage('Error during deletion: ' + error.message, 'error');
            }
          }
        });
      } else {
        // Single item - use original logic
        if (node.data && node.data.isWorksheet) {
          showConfirm(`Delete sheet "${node.text}" from Excel?`, async (confirmed) => {
            if (confirmed) {
              try {
                await deleteSheet(node.data.sheetName);
                tree.delete_node(node);
                await saveStructureToStorage();
                showMessage('Sheet deleted successfully!', 'success');
              } catch (error) {
                console.error('Delete failed:', error);
                showMessage('Error deleting: ' + error.message, 'error');
              }
            }
          });
        } else {
          showConfirm(`Delete folder "${node.text}" and all its contents?`, (confirmed) => {
            if (confirmed) {
              tree.delete_node(node);
              saveStructureToStorage().catch(err => console.error('Save failed:', err));
              showMessage('Folder deleted successfully!', 'success');
            }
          });
        }
      }
    }
  };
  
  return menu;
}


// Delete a sheet from Excel
async function deleteSheet(sheetName) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      sheet.delete();
      await context.sync();
    });
  } catch (error) {
    console.error("Error deleting sheet:", error);
    throw error;
  }
}

// Save structure
async function saveStructure() {
  try {
    await saveStructureToStorage();
    showMessage('Structure saved successfully!', 'success');
  } catch (error) {
    console.error('Save structure failed:', error);
    showMessage('Error saving: ' + error.message, 'error');
  }
}

// Save structure to Document Settings
async function saveStructureToStorage() {
  if (!treeInitialized) {
    console.log('Tree not ready, skipping save');
    return;
  }

  console.log('Saving structure to Document Settings...');

  try {
    const tree = $('#sheet-tree').jstree(true);
    if (!tree) {
      console.error('Tree instance not available');
      return;
    }

    const treeData = tree.get_json('#', { flat: false });
    console.log('Tree data to save:', treeData); // DEBUG
    
    const treeDataString = JSON.stringify(treeData);
    
    // Check size (optional warning)
    const dataSize = treeDataString.length;
    console.log(`Tree data size: ${(dataSize / 1024).toFixed(2)} KB`);
    
    if (dataSize > 1900000) { // ~1.9MB safety margin
      console.warn('Tree structure is very large, may hit storage limits');
      showMessage('Warning: Tree structure is very large', 'warning');
    }
    
    console.log('About to call Excel.run...'); // DEBUG
    
    await Excel.run(async (context) => {
      console.log('Inside Excel.run context'); // DEBUG
      
      const settings = context.workbook.settings;
      
      // Remove old setting if it exists
      const oldSetting = settings.getItemOrNullObject("treeStructure");
      await context.sync();
      console.log('After first sync'); // DEBUG
      
      if (!oldSetting.isNullObject) {
        console.log('Removing old tree structure setting');
        oldSetting.delete();
        await context.sync();
        console.log('After delete sync'); // DEBUG
      }
      
      // Add new setting
      console.log('Adding new tree structure setting');
      settings.add("treeStructure", treeDataString);
      await context.sync();
      console.log('After add sync'); // DEBUG
      
      console.log('Structure saved successfully to Document Settings');
    });
  } catch (error) {
    console.error("Error in saveStructureToStorage:", error);
    console.error("Error message:", error.message); // DEBUG
    console.error("Error code:", error.code); // DEBUG
    if (error.debugInfo) {
      console.error("Error details:", JSON.stringify(error.debugInfo, null, 2)); // DEBUG
    }
    throw error;
  }
}

// Load structure from Document Settings
async function loadStructure() {
  console.log('Loading structure from Document Settings...');
  
  try {
    await Excel.run(async (context) => {
      const settings = context.workbook.settings;
      const treeSetting = settings.getItemOrNullObject("treeStructure");
      treeSetting.load("value");
      await context.sync();
      
      if (!treeSetting.isNullObject && treeSetting.value) {
        console.log('Found saved structure in Document Settings');
        const treeData = JSON.parse(treeSetting.value);
        const tree = $('#sheet-tree').jstree(true);
        tree.settings.core.data = treeData;
        tree.refresh();

        $('#sheet-tree').one('refresh.jstree', function() {
          console.log('Tree refresh completed');
          registerWorksheetEventHandlers();
        });

        console.log('Structure loaded successfully');
        return;
      }
      
      console.log('No saved structure found, loading default');
      await loadDefaultStructure();
    });
  } catch (error) {
    console.error("Error loading structure:", error);
    if (error.debugInfo) {
      console.error("Error details:", error.debugInfo);
    }
    showMessage('Error loading structure, using default', 'warning');
    await loadDefaultStructure();
  }
}

// Load default structure from Excel sheets
async function loadDefaultStructure() {
  console.log('Loading default structure from sheets...');
  
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name, items/visibility");
      await context.sync();

      const treeData = sheets.items
        .filter(sheet => sheet.visibility === Excel.SheetVisibility.visible)
        .map((sheet, index) => ({
          id: `sheet_${index}_${Date.now()}`,
          text: sheet.name,
          type: 'sheet',
          icon: 'jstree-file',
          data: {
            sheetName: sheet.name,
            isWorksheet: true,
            nodeType: 'sheet'
          }
        }));

      console.log('Default tree data:', treeData);

      const tree = $('#sheet-tree').jstree(true);
      tree.settings.core.data = treeData;
      tree.refresh();

      $('#sheet-tree').one('refresh.jstree', function() {
        console.log('Tree refresh completed');
        registerWorksheetEventHandlers();
      });
      
      // Save the default structure
      await saveStructureToStorage();
    });
  } catch (error) {
    console.error("Error loading default structure:", error);
    showMessage('Error loading sheets', 'error');
  }
}


// Register event handlers - call this after tree is loaded
async function registerWorksheetEventHandlers() {
  if (eventHandlersRegistered) {
    console.log('Event handlers already registered');
    return;
  }
  
  console.log('Attempting to register worksheet event handlers...');
  
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      
      // Check if event handlers are supported
      if (typeof sheets.onAdded === 'undefined') {
        console.log('Event handlers not supported in this Excel version');
        return false;
      }
      
      // Register the onAdded event handler
      console.log('Registering onAdded...');
      sheets.onAdded.add(onSheetAdded);
      
      console.log('Registering onNameChanged...');
      sheets.onNameChanged.add(onSheetRenamed);
      
      console.log('Registering onDeleted...');
      sheets.onDeleted.add(onSheetDeleted);
      
      await context.sync();
      
      eventHandlersRegistered = true;
      console.log('✓ All event handlers registered successfully');
      return true;
    });
  } catch (error) {
    console.error('Event handlers registration failed:', error.message);
    if (error.debugInfo) {
      console.error('Error details:', error.debugInfo);
    }
    return false;
  }
}


// Handler function when a new sheet is added in Excel
async function onSheetAdded(event) {
  console.log('Sheet added event triggered:', event);
  
  if (!treeInitialized) {
    console.log('Tree not initialized yet, skipping event');
    return;
  }
  
  try {
    await Excel.run(async (context) => {
      
      let sheet;
      try {
        sheet = context.workbook.worksheets.getItem(event.worksheetId);
      } catch (e) {
        // Fallback: refresh the entire tree if we can't get the specific sheet
        console.log('Could not get sheet by ID, refreshing tree instead');
        await refreshTree();
        return;
      }
      
      sheet.load("name, visibility");
      await context.sync();
      
      console.log('New sheet detected:', sheet.name, 'Visibility:', sheet.visibility);
      
      // Only add visible sheets to the tree
      if (sheet.visibility !== Excel.SheetVisibility.visible) {
        console.log('Sheet is hidden, not adding to tree');
        return;
      }
      
      const tree = $('#sheet-tree').jstree(true);
      
      if (!tree) {
        console.error('Tree instance not available');
        return;
      }
      
      // Check if sheet already exists in tree (avoid duplicates)
      const currentNodes = tree.get_json('#', { flat: true });
      const exists = currentNodes.some(node => 
        node.data && 
        node.data.isWorksheet && 
        node.data.sheetName === sheet.name
      );
      
      if (exists) {
        console.log('Sheet already exists in tree, skipping');
        return;
      }
      
      // Add the new sheet to the tree (at root level)
      console.log('Adding new sheet to tree:', sheet.name);
      const nodeId = 'sheet_' + Date.now() + '_' + Math.random();
      const newNode = tree.create_node('#', {
        id: nodeId,
        text: sheet.name,
        type: 'sheet',
        icon: 'jstree-file',
        data: {
          isWorksheet: true,
          sheetName: sheet.name,
          nodeType: 'sheet'
        }
      }, 'last');
      
      if (newNode) {
        console.log('Sheet node created successfully:', nodeId);
        
        // Save the updated tree structure to document storage
        await saveStructureToStorage();
        
        // Show success message to user
        showMessage(`Sheet "${sheet.name}" added to tree`, 'success');
      } else {
        console.error('Failed to create tree node for sheet:', sheet.name);
      }
    });
  } catch (error) {
    console.error('Error in onSheetAdded handler:', error);
    if (error.debugInfo) {
      console.error('Error details:', error.debugInfo);
    }
  }
}

// Handler function when a sheet is renamed in Excel
async function onSheetRenamed(event) {
  console.log('========================================');
  console.log('Sheet renamed event triggered!');
  console.log('Event object:', event);
  console.log('Event worksheetId:', event.worksheetId);
  console.log('Event nameAfter:', event.nameAfter);
  console.log('Event nameBefore:', event.nameBefore);
  console.log('========================================');
  
  if (!treeInitialized) {
    console.log('Tree not initialized yet, skipping event');
    return;
  }
  
  try {
    await Excel.run(async (context) => {
      // Get the renamed sheet
      let sheet;
      try {
        console.log('Attempting to get sheet with ID:', event.worksheetId);
        sheet = context.workbook.worksheets.getItem(event.worksheetId);
        sheet.load("name, id");
        await context.sync();
        console.log('Successfully loaded sheet. Name:', sheet.name, 'ID:', sheet.id);
      } catch (e) {
        console.error('Could not get sheet by ID:', e);
        console.log('Falling back to refreshTree()');
        await refreshTree();
        return;
      }
      
      const newName = sheet.name;
      console.log('New sheet name:', newName);
      
      const tree = $('#sheet-tree').jstree(true);
      
      if (!tree) {
        console.error('Tree instance not available');
        return;
      }
      
      // Get all current nodes
      const currentNodes = tree.get_json('#', { flat: true });
      console.log('Current tree nodes:', currentNodes.length);
      console.log('Tree nodes with sheet data:', currentNodes.filter(n => n.data && n.data.isWorksheet).map(n => ({
        id: n.id,
        text: n.text,
        sheetName: n.data.sheetName
      })));
      
      // Get all current Excel sheet names
      const sheets = context.workbook.worksheets;
      sheets.load("items/name, items/id");
      await context.sync();
      
      const excelSheetNames = sheets.items.map(s => s.name);
      console.log('Current Excel sheet names:', excelSheetNames);
      
      // Find the orphaned node (one that doesn't match any current Excel sheet)
      const orphanedNode = currentNodes.find(node =>
        node.data &&
        node.data.isWorksheet &&
        !excelSheetNames.includes(node.data.sheetName)
      );
      
      if (orphanedNode) {
        console.log('Found orphaned node to update:');
        console.log('  - Node ID:', orphanedNode.id);
        console.log('  - Old name:', orphanedNode.data.sheetName);
        console.log('  - New name:', newName);
        
        // Update the node text and data
        tree.rename_node(orphanedNode.id, newName);
        orphanedNode.data.sheetName = newName;
        
        console.log('Node renamed successfully');
        
        // Save the updated structure
        await saveStructureToStorage();
        console.log('Structure saved');
        
        showMessage(`Sheet renamed to "${newName}" in tree`, 'success');
      } else {
        console.warn('Could not find orphaned node!');
        console.log('This might mean:');
        console.log('  1. The sheet was already updated in the tree');
        console.log('  2. The sheet is not in the tree');
        console.log('  3. There is a timing issue');
        console.log('Refreshing tree as fallback...');
        await refreshTree();
      }
    });
  } catch (error) {
    console.error('Error in onSheetRenamed handler:', error);
    if (error.debugInfo) {
      console.error('Error details:', error.debugInfo);
    }
  }
}


// Handler function when a sheet is deleted
async function onSheetDeleted(event) {
  console.log('Sheet deleted event triggered:', event);
  
  if (!treeInitialized) {
    console.log('Tree not initialized yet, skipping event');
    return;
  }
  
  try {
    // Simply refresh the tree to remove deleted sheets
    await refreshTree();
    showMessage('Sheet deleted, tree updated', 'info');
  } catch (error) {
    console.error('Error in onSheetDeleted handler:', error);
  }
}


// Test function to check event handler status
async function testEventHandlers() {
  console.log('=== EVENT HANDLER TEST ===');
  console.log('eventHandlersRegistered flag:', eventHandlersRegistered);
  console.log('treeInitialized flag:', treeInitialized);
  
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      
      console.log('sheets.onAdded type:', typeof sheets.onAdded);
      console.log('sheets.onNameChanged type:', typeof sheets.onNameChanged);
      console.log('sheets.onDeleted type:', typeof sheets.onDeleted);
      
      // Try to manually trigger a test
      console.log('Attempting to register a test handler...');
      
      if (typeof sheets.onNameChanged !== 'undefined') {
        sheets.onNameChanged.add(function(event) {
          console.log('🔥 TEST RENAME EVENT FIRED! 🔥', event);
        });
        await context.sync();
        console.log('✓ Test handler registered. Now try renaming a sheet.');
      } else {
        console.log('❌ onNameChanged is not available in this Excel version');
      }
    });
  } catch (error) {
    console.error('Test failed:', error);
  }
}


async function hideOtherSheets() {
  console.log('Hide others clicked');
  
  if (!treeInitialized) {
    showMessage('Tree is still loading, please wait...', 'warning');
    return;
  }

  try {
    await Excel.run(async (context) => {
      // Get the active sheet
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name");
      
      // Get all sheets
      const sheets = context.workbook.worksheets;
      sheets.load("items/name, items/visibility");
      await context.sync();
      
      const activeSheetName = activeSheet.name;
      let hiddenCount = 0;
      
      // Hide all sheets except the active one
      sheets.items.forEach(sheet => {
        if (sheet.name !== activeSheetName && sheet.visibility === Excel.SheetVisibility.visible) {
          sheet.visibility = Excel.SheetVisibility.hidden;
          hiddenCount++;
        }
      });
      
      await context.sync();
      
      console.log(`Hidden ${hiddenCount} sheets, kept "${activeSheetName}" visible`);
      
      // Refresh the tree to reflect changes
      await refreshTree();
      
      showMessage(`Hidden ${hiddenCount} sheet(s). Only "${activeSheetName}" is visible.`, 'success');
    });
  } catch (error) {
    console.error("Error hiding sheets:", error);
    showMessage("Error hiding sheets: " + error.message, 'error');
  }
}


// Check if a sheet name already exists in Excel
async function checkSheetNameExists(sheetName) {
  try {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      
      const exists = sheets.items.some(sheet => 
        sheet.name.toLowerCase() === sheetName.toLowerCase()
      );
      
      return exists;
    });
  } catch (error) {
    console.error("Error checking sheet name:", error);
    return false;
  }
}

// Copy selected nodes to clipboard
function copyNodes() {
  const tree = $('#sheet-tree').jstree(true);
  const selected = tree.get_selected(true);
  
  if (selected.length === 0) {
    showMessage('No items selected to copy', 'warning');
    return;
  }
  
  // Store copies of the selected nodes
  copiedNodes = selected.map(node => {
    // Get the full node data including children
    const nodeJson = tree.get_json(node.id);
    
    return {
      text: node.text,
      type: node.type,
      data: node.data ? JSON.parse(JSON.stringify(node.data)) : null,
      children: nodeJson && nodeJson.children ? nodeJson.children : []
    };
  });
  
  console.log('Copied nodes:', copiedNodes);
  showMessage(`Copied ${copiedNodes.length} item(s)`, 'success');
}

// Paste copied nodes
async function pasteNodes() {
  if (copiedNodes.length === 0) {
    showMessage('Nothing to paste', 'warning');
    return;
  }
  
  const tree = $('#sheet-tree').jstree(true);
  const selected = tree.get_selected();
  const parent = selected.length > 0 ? selected[0] : '#';
  
  console.log('Pasting', copiedNodes.length, 'nodes to parent:', parent);
  
  try {
    for (const nodeToCopy of copiedNodes) {
      if (nodeToCopy.data && nodeToCopy.data.isWorksheet) {
        // It's a sheet - duplicate it in Excel
        await duplicateSheet(nodeToCopy.data.sheetName, parent, tree);
      } else {
        // It's a folder - copy the folder structure
        await copyFolderStructure(nodeToCopy, parent, tree);
      }
    }
    
    await saveStructureToStorage();
    showMessage(`Pasted ${copiedNodes.length} item(s) successfully`, 'success');
    
  } catch (error) {
    console.error('Paste failed:', error);
    showMessage('Error pasting: ' + error.message, 'error');
  }
}

// Duplicate a sheet in Excel and add to tree
async function duplicateSheet(sheetName, parentNode, tree) {
  return await Excel.run(async (context) => {
    // Get the source sheet
    const sourceSheet = context.workbook.worksheets.getItem(sheetName);
    sourceSheet.load("name");
    await context.sync();
    
    // Get all existing sheet names to find a unique name
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    
    const existingNames = sheets.items.map(s => s.name.toLowerCase());
    
    // Find a unique name for the copy
    let copyName = `${sheetName} Copy`;
    let counter = 2;
    
    while (existingNames.includes(copyName.toLowerCase())) {
      copyName = `${sheetName} Copy (${counter})`;
      counter++;
    }
    
    // Copy the sheet
    const copiedSheet = sourceSheet.copy(Excel.WorksheetPositionType.end);
    copiedSheet.name = copyName;
    await context.sync();
    
    console.log(`Duplicated sheet "${sheetName}" as "${copyName}"`);
    
    // Add to tree
    const nodeId = 'sheet_' + Date.now() + '_' + Math.random();
    tree.create_node(parentNode, {
      id: nodeId,
      text: copyName,
      type: 'sheet',
      icon: 'jstree-file',
      data: {
        isWorksheet: true,
        sheetName: copyName,
        nodeType: 'sheet'
      }
    }, 'last');
    
    tree.open_node(parentNode);
  });
}

// Copy folder structure recursively
async function copyFolderStructure(folderNode, parentNode, tree) {
  // Create the folder copy
  const nodeId = 'folder_' + Date.now() + '_' + Math.random();
  const newFolderId = tree.create_node(parentNode, {
    id: nodeId,
    text: folderNode.text + ' Copy',
    type: 'folder',
    icon: 'jstree-folder',
    data: {
      isWorksheet: false,
      nodeType: 'folder'
    }
  }, 'last');
  
  tree.open_node(parentNode);
  
  // Recursively copy children
  if (folderNode.children && folderNode.children.length > 0) {
    for (const child of folderNode.children) {
      if (child.data && child.data.isWorksheet) {
        // Child is a sheet
        await duplicateSheet(child.data.sheetName, newFolderId, tree);
      } else {
        // Child is a folder
        await copyFolderStructure(child, newFolderId, tree);
      }
    }
  }
}