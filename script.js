const { ipcRenderer } = require('electron');

// Function to populate the treeview
function populateTreeview(data) {
    const treeviewContainer = document.getElementById('treeview');
    treeviewContainer.innerHTML = ''; // Clear existing tree nodes

    // Create head node "Projects"
    const headNode = document.createElement('li');
    headNode.textContent = 'Projects';
    treeviewContainer.appendChild(headNode);

    // Add sub-nodes under the "Projects" node
    const subTree = document.createElement('ul');
    subTree.classList.add('nested'); // Nested items are initially hidden
    headNode.appendChild(subTree);

    // Loop through the data and create sub-nodes
    data.forEach((project, index) => {
        const subNode = document.createElement('li');
        subNode.textContent = project;
        subNode.addEventListener('click', () => {
            alert(`Hello, I'm node no: ${index + 1}`);
        });
        subTree.appendChild(subNode);
    });
}

// Load Excel data and populate treeview on page load
window.onload = () => {
    ipcRenderer.invoke('read-excel').then((data) => {
        console.log(data);
        populateTreeview(data);
    }).catch((error) => {
        console.error('Error fetching Excel data:', error);
    });
};






