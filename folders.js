// Separate function to load folders
function loadFolders(source = 'sidebar') {
  // Only process sidebar requests - ignore 'main' source
  if (source !== 'sidebar') return;
  
  // Get the folder display element
  const folderDisplay = document.getElementById('folderContainerSidebar');
  
  if (!folderDisplay) {
    console.error('Error: folder display element not found in loadFolders');
    return; // Exit if element is not found
  }
  
  // Clear the current content first
  folderDisplay.innerHTML = '';
  
  // Create a loading indicator using DOM methods
  const loadingDiv = document.createElement('div');
  loadingDiv.className = 'loading';
  loadingDiv.id = 'folderLoadingIndicator';
  
  const loadingIcon = document.createElement('i');
  loadingIcon.className = 'material-icons';
  loadingIcon.textContent = 'sync';
  
  loadingDiv.appendChild(loadingIcon);
  loadingDiv.appendChild(document.createTextNode(' Loading...'));
  
  folderDisplay.appendChild(loadingDiv);
  
  // Add timeout to prevent infinite loading
  let timeoutId = setTimeout(() => {
    // Check if loading indicator still exists after timeout
    if (document.getElementById('folderLoadingIndicator')) {
      console.log('Folder data loading timeout triggered');
      folderDisplay.innerHTML = '';
      
      const errorBox = document.createElement('div');
      errorBox.className = 'error-box';
      
      const errorStrong = document.createElement('strong');
      errorStrong.textContent = 'Loading Error:';
      
      errorBox.appendChild(errorStrong);
      errorBox.appendChild(document.createTextNode(' Request timed out. '));
      
      const refreshButton = document.createElement('button');
      refreshButton.className = 'btn-small';
      refreshButton.style.marginLeft = '10px';
      refreshButton.innerHTML = '<i class="material-icons">refresh</i> Retry';
      refreshButton.addEventListener('click', () => loadFolders(source));
      
      errorBox.appendChild(refreshButton);
      folderDisplay.appendChild(errorBox);
    }
  }, 15000); // 15 second timeout
  
  // Get the project data from the modal
  const modal = document.getElementById('projectDetailsModal');
  const folderId = modal ? modal.getAttribute('data-folder-id') : null;
  const projectName = modal ? modal.getAttribute('data-project-name') : 'Project';
  
  console.log('loadFolders: Found folderId:', folderId);
  console.log('loadFolders: Project name:', projectName);
  
  google.script.run
    .withSuccessHandler(data => {
      // Clear the timeout since we got a response
      clearTimeout(timeoutId);
      
      // Check for error message
      if (data.error) {
        // Clear the container
        folderDisplay.innerHTML = '';
        
        // Create error message using DOM methods
        const errorBox = document.createElement('div');
        errorBox.className = 'error-box';
        
        const errorStrong = document.createElement('strong');
        errorStrong.textContent = 'Error:';
        
        errorBox.appendChild(errorStrong);
        errorBox.appendChild(document.createTextNode(' ' + data.error));
        
        folderDisplay.appendChild(errorBox);
        return;
      }
      
      if (!data.files) {
        // Clear the container
        folderDisplay.innerHTML = '';
        
        // Create empty state using DOM methods
        const emptyState = document.createElement('div');
        emptyState.className = 'empty-state';
        
        const emptyIcon = document.createElement('i');
        emptyIcon.className = 'material-icons';
        emptyIcon.textContent = 'folder_off';
        
        const emptyText = document.createElement('p');
        emptyText.textContent = 'No files found. Please configure a folder in Settings.';
        
        emptyState.appendChild(emptyIcon);
        emptyState.appendChild(emptyText);
        
        folderDisplay.appendChild(emptyState);
        return;
      }
      
      // Use DocumentFragment to build the entire structure once
      const fragment = document.createDocumentFragment();
      const folderEntries = Object.entries(data.files);
      
      folderEntries.forEach(([folder, files], index) => {
        const secId = 'sec' + index;
        
        // Create folder header
        const toggleHeader = document.createElement('div');
        toggleHeader.className = 'toggle-header';
        
        const folderIcon = document.createElement('i');
        folderIcon.className = 'material-icons';
        folderIcon.textContent = 'folder';
        
        const chevronIcon = document.createElement('i');
        chevronIcon.className = 'material-icons chevron';
        chevronIcon.style.marginLeft = 'auto';
        chevronIcon.textContent = 'chevron_right';
        
        toggleHeader.appendChild(folderIcon);
        toggleHeader.appendChild(document.createTextNode(' ' + folder));
        toggleHeader.appendChild(chevronIcon);
        
        toggleHeader.addEventListener('click', function() {
          const fileList = document.getElementById(secId);
          fileList.style.display = fileList.style.display === 'block' ? 'none' : 'block';
          
          // Toggle the folder icon
          folderIcon.textContent = fileList.style.display === 'block' ? 'folder_open' : 'folder';
          
          // Toggle the chevron icon
          chevronIcon.textContent = fileList.style.display === 'block' ? 'expand_more' : 'chevron_right';
        });
        
        fragment.appendChild(toggleHeader);
        
        // Create file list container
        const fileList = document.createElement('div');
        fileList.className = 'file-list';
        fileList.id = secId;
        
        // Group files by type if they have categories
        const filesByType = {};
        
        files.forEach(file => {
          const type = file.type || 'Other';
          if (!filesByType[type]) {
            filesByType[type] = [];
          }
          filesByType[type].push(file);
        });
        
        // Add files by category
        Object.entries(filesByType).forEach(([type, typeFiles]) => {
          // Add type label if there are multiple types
          if (Object.keys(filesByType).length > 1) {
            const typeLabel = document.createElement('div');
            typeLabel.className = 'section-label';
            typeLabel.textContent = type;
            fileList.appendChild(typeLabel);
          }
          
          // Add each file
          typeFiles.forEach(file => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file';
            
            // Choose appropriate icon based on file type
            let iconName = 'insert_drive_file';
            if (file.name.toLowerCase().endsWith('.pdf')) {
              iconName = 'picture_as_pdf';
            } else if (file.name.toLowerCase().match(/\.(jpg|jpeg|png|gif)$/)) {
              iconName = 'image';
            } else if (file.name.toLowerCase().match(/\.(doc|docx)$/)) {
              iconName = 'description';
            } else if (file.name.toLowerCase().match(/\.(xls|xlsx)$/)) {
              iconName = 'table_chart';
            } else if (type === 'Boards') {
              iconName = 'dashboard';
            } else if (type === 'Briefs') {
              iconName = 'description';
            }
            
            const fileIcon = document.createElement('i');
            fileIcon.className = 'material-icons';
            fileIcon.textContent = iconName;
            
            const fileLink = document.createElement('a');
            fileLink.href = file.url;
            fileLink.target = '_blank';
            fileLink.textContent = file.name;
            
            fileItem.appendChild(fileIcon);
            fileItem.appendChild(fileLink);
            fileList.appendChild(fileItem);
          });
        });
        
        fragment.appendChild(fileList);
      });
      
      // Clear the container and add the entire fragment at once
      folderDisplay.innerHTML = '';
      folderDisplay.appendChild(fragment);
    })
    .withFailureHandler(error => {
      // Clear the timeout since we got a response
      clearTimeout(timeoutId);
      
      // Clear the container
      folderDisplay.innerHTML = '';
      
      // Create error message using DOM methods
      const errorBox = document.createElement('div');
      errorBox.className = 'error-box';
      
      const errorStrong = document.createElement('strong');
      errorStrong.textContent = 'Error loading data:';
      
      errorBox.appendChild(errorStrong);
      errorBox.appendChild(document.createTextNode(' ' + error));
      
      const refreshButton = document.createElement('button');
      refreshButton.className = 'btn-small';
      refreshButton.style.marginLeft = '10px';
      refreshButton.innerHTML = '<i class="material-icons">refresh</i> Retry';
      refreshButton.addEventListener('click', () => loadFolders(source));
      
      errorBox.appendChild(refreshButton);
      folderDisplay.appendChild(errorBox);
    })
    .getDashboardData(folderId, projectName);
}