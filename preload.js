const { ipcRenderer } = require('electron');

// Listen for the 'update-treeview' event from the main process
ipcRenderer.on('update-treeview', (event, nodes) => {
    updateTreeView(nodes);
});

// Function to update the treeview with matched nodes
function updateTreeView(nodes) {
    const treeview = document.getElementById('treeview');
    const headNode = document.createElement('li');
    headNode.innerText = 'Projects';
    treeview.appendChild(headNode);

    nodes.forEach(node => {
        const subNode = document.createElement('li');
        subNode.innerText = node;
        subNode.onclick = nodeClicked; // Assign the node click handler
        headNode.appendChild(subNode);
    });
}

// Function to handle node clicks
function nodeClicked() {
    const clickedNode = event.target; // Get the clicked node
    const nodeCaption = clickedNode.innerText.trim();

    // Extract the number before the "-" sign
    const number = nodeCaption.split('-')[0].trim();

    // Now perform the search logic
    findMatchingRowInSrpnSheet(number);
}
