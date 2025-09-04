// D3.js Organization Chart Visualization
const OrgChart = (function() {
  'use strict';
  
  // Private variables
  let svg = null;
  let g = null;
  let root = null;
  let treemap = null;
  let zoom = null;
  let i = 0;
  let orgData = [];
  let licensedEmails = new Set();
  let highlightedDepartment = null;
  let highlightedNodes = null;
  let licenseHighlightActive = false;
  
  // Configuration
  const config = {
    horizontalSpacing: 200,
    verticalSpacing: 1,
    nodeRadius: 10,
    duration: 750,
    margin: { top: 20, right: 120, bottom: 20, left: 120 },
    showEngagementStars: true
  };
  
  // Color scales
  let departmentColorScale = null;
  let levelColorScale = null;
  
  // Public API
  return {
    // Initialize the visualization module
    initialize: function() {
      // Set up event listeners
      const horizontalSlider = document.getElementById('horizontalSpacing');
      if (horizontalSlider) {
        horizontalSlider.addEventListener('input', (e) => {
          config.horizontalSpacing = parseInt(e.target.value);
          document.getElementById('horizontalValue').textContent = e.target.value;
        });
      }
      
      const verticalSlider = document.getElementById('verticalSpacing');
      if (verticalSlider) {
        verticalSlider.addEventListener('input', (e) => {
          config.verticalSpacing = parseFloat(e.target.value);
          document.getElementById('verticalValue').textContent = e.target.value;
        });
      }

      const showStarsCheckbox = document.getElementById('showStars');
      if (showStarsCheckbox) {
        config.showEngagementStars = showStarsCheckbox.checked;
        showStarsCheckbox.addEventListener('change', (e) => {
          config.showEngagementStars = e.target.checked;
          if (root) this.update(root);
        });
      }
    },
    
    // Create the visualization
    createVisualization: function(hierarchyData, userData, licenses) {
      if (!hierarchyData) {
        console.error('No hierarchy data provided');
        return;
      }
      
      // Reset any existing search or department highlights
      highlightedDepartment = null;
      highlightedNodes = null;
      licenseHighlightActive = false;
      document.querySelectorAll('.dept-row').forEach(row => row.classList.remove('active'));
      const searchInput = document.getElementById('searchInput');
      if (searchInput) searchInput.value = '';
      const clearBtn = document.getElementById('clearHighlightBtn');
      if (clearBtn) clearBtn.style.display = 'none';
      const licenseBtn = document.getElementById('licenseToggleBtn');
      if (licenseBtn) licenseBtn.classList.remove('active');

      orgData = userData;
      licensedEmails = licenses;

      // Set up color scales
      this.setupColorScales();
      
      // Get container dimensions
      const container = document.getElementById('treeContainer');
      const width = container.clientWidth - config.margin.left - config.margin.right;
      const height = container.clientHeight - config.margin.top - config.margin.bottom;
      
      // Clear previous visualization
      d3.select('#treeContainer').selectAll('*').remove();
      
      // Create SVG
      svg = d3.select('#treeContainer')
        .append('svg')
        .attr('width', container.clientWidth)
        .attr('height', container.clientHeight);
      
      // Create group for zoom
      g = svg.append('g')
        .attr('transform', `translate(${config.margin.left},${height / 2})`);
      
      // Set up zoom behavior
      zoom = d3.zoom()
        .scaleExtent([0.1, 3])
        .on('zoom', (event) => {
          g.attr('transform', event.transform);
        });
      
      svg.call(zoom);
      
      // Create tree layout
      treemap = d3.tree().size([height, width]);
      
      // Create hierarchy
      root = d3.hierarchy(hierarchyData);
      root.x0 = height / 2;
      root.y0 = 0;
      
      // Store original positions
      root.descendants().forEach(d => {
        d.x0 = d.x;
        d.y0 = d.y;
      });
      
      // Collapse after the second level initially
      if (root.children) {
        root.children.forEach(this.collapse);
      }
      
      // Initial render
      this.update(root);
      
      // Show controls and panels
      document.getElementById('legend').style.display = 'block';
      document.getElementById('statsPanel').style.display = 'block';
      document.getElementById('bottomControls').style.display = 'flex';
      
      this.updateLegend();
      this.updateStatsPanel();
    },
    
    // Setup color scales
    setupColorScales: function() {
      const departments = [...new Set(orgData.map(u => u.department))];
      // Use a color scheme that avoids reds to prevent confusion with license borders
      const colorPalette = [
        '#1f77b4', // blue
        '#2ca02c', // green
        '#ff7f0e', // orange
        '#9467bd', // purple
        '#8c564b', // brown
        '#e377c2', // pink
        '#7f7f7f', // gray
        '#bcbd22', // olive
        '#17becf', // cyan
        '#aec7e8', // light blue
        '#ffbb78', // light orange
        '#98df8a', // light green
        '#c5b0d5', // light purple
        '#c49c94', // light brown
        '#f7b6d2', // light pink
        '#c7c7c7', // light gray
        '#dbdb8d', // light olive
        '#9edae5'  // light cyan
      ];
      
      departmentColorScale = d3.scaleOrdinal()
        .domain(departments)
        .range(colorPalette);
      
      levelColorScale = d3.scaleSequential()
        .domain([0, 5])
        .interpolator(d3.interpolateBlues);
    },
    
    // Update the tree
    update: function(source) {
      if (!root || !treemap) return;
      
      // Compute the new tree layout
      const treeData = treemap(root);
      const nodes = treeData.descendants();
      const links = treeData.links();
      
      // Normalize for fixed-depth and apply spacing
      nodes.forEach(d => {
        d.y = d.depth * config.horizontalSpacing;
        d.x = d.x * config.verticalSpacing;
      });
      
      // -------------------- NODES --------------------
      const node = g.selectAll('g.node')
        .data(nodes, d => d.id || (d.id = ++i));
      
      // Enter new nodes at the parent's previous position
      const nodeEnter = node.enter().append('g')
        .attr('class', 'node')
        .attr('transform', d => `translate(${source.y0 || 0},${source.x0 || 0})`)
        .on('click', (event, d) => this.click(event, d))
        .on('mouseover', (event, d) => this.showTooltip(event, d))
        .on('mouseout', () => this.hideTooltip());
      
      // Add circles for nodes
      nodeEnter.append('circle')
        .attr('r', 1e-6)
        .style('fill', d => this.getNodeColor(d))
        .style('stroke', d => this.getNodeStroke(d))
        .style('opacity', d => this.getNodeOpacity(d));
      
      // Add labels for nodes - include star rating if available
      nodeEnter.append('text')
        .attr('dy', '.35em')
        .attr('x', d => d.children || d._children ? -13 : 13)
        .attr('text-anchor', d => d.children || d._children ? 'end' : 'start')
        .text(d => this.getNodeLabel(d))
        .style('fill-opacity', 1e-6);
      
      // Add title (job title) as second line
      nodeEnter.append('text')
        .attr('class', 'title')
        .attr('dy', '1.5em')
        .attr('x', d => d.children || d._children ? -13 : 13)
        .attr('text-anchor', d => d.children || d._children ? 'end' : 'start')
        .text(d => d.data.title || '')
        .style('fill-opacity', 1e-6);
      
      // UPDATE
      const nodeUpdate = nodeEnter.merge(node);
      
      // Transition to the proper position for the node
      nodeUpdate.transition()
        .duration(config.duration)
        .attr('transform', d => `translate(${d.y},${d.x})`);
      
      // Update the node attributes and style
      nodeUpdate.select('circle')
        .attr('r', config.nodeRadius)
        .style('fill', d => this.getNodeColor(d))
        .style('stroke', d => this.getNodeStroke(d))
        .style('opacity', d => this.getNodeOpacity(d))
        .attr('cursor', 'pointer');
      
      nodeUpdate.select('text')
        .text(d => this.getNodeLabel(d))
        .style('fill-opacity', d => this.getNodeOpacity(d));
      
      nodeUpdate.select('text.title')
        .style('fill-opacity', d => this.getNodeOpacity(d) * 0.7);
      
      // EXIT
      const nodeExit = node.exit().transition()
        .duration(config.duration)
        .attr('transform', d => `translate(${source.y},${source.x})`)
        .remove();
      
      nodeExit.select('circle')
        .attr('r', 1e-6);
      
      nodeExit.select('text')
        .style('fill-opacity', 1e-6);
      
      // -------------------- LINKS --------------------
      const link = g.selectAll('path.link')
        .data(links, d => d.target.id);
      
      // Enter new links at the parent's previous position
      const linkEnter = link.enter().insert('path', 'g')
        .attr('class', 'link')
        .attr('d', d => {
          const o = { x: source.x0 || 0, y: source.y0 || 0 };
          return this.diagonal(o, o);
        });
      
      // UPDATE
      const linkUpdate = linkEnter.merge(link);
      
      // Transition back to the parent element position
      linkUpdate.transition()
        .duration(config.duration)
        .attr('d', d => this.diagonal(d.source, d.target))
        .style('opacity', d => this.getLinkOpacity(d));
      
      // EXIT
      link.exit().transition()
        .duration(config.duration)
        .attr('d', d => {
          const o = { x: source.x, y: source.y };
          return this.diagonal(o, o);
        })
        .remove();
      
      // Store the old positions for transition
      nodes.forEach(d => {
        d.x0 = d.x;
        d.y0 = d.y;
      });
    },
    
    // Creates a curved (diagonal) path from parent to child nodes
    diagonal: function(source, target) {
      // Handle both cases: when called with link object or separate source/target
      const s = source.source ? source.source : source;
      const d = source.target ? source.target : target;
      
      // Safety check for valid coordinates
      if (!s || !d || 
          s.y === undefined || s.x === undefined || 
          d.y === undefined || d.x === undefined) {
        return '';
      }
      
      return `M ${s.y} ${s.x}
              C ${(s.y + d.y) / 2} ${s.x},
                ${(s.y + d.y) / 2} ${d.x},
                ${d.y} ${d.x}`;
    },
    
    // Toggle children on click
    click: function(event, d) {
      if (d.children) {
        d._children = d.children;
        d.children = null;
      } else {
        d.children = d._children;
        d._children = null;
      }
      this.update(d);
    },
    
    // Collapse node
    collapse: function(d) {
      if (d.children) {
        d._children = d.children;
        d._children.forEach(OrgChart.collapse);
        d.children = null;
      }
    },
    
    // Expand all nodes
    expandAll: function() {
      if (!root) return;
      
      function expand(d) {
        if (d._children) {
          d.children = d._children;
          d._children = null;
        }
        if (d.children) {
          d.children.forEach(expand);
        }
      }
      
      expand(root);
      this.update(root);
    },
    
    // Collapse all nodes
    collapseAll: function() {
      if (!root) return;
      
      if (root.children) {
        root.children.forEach(this.collapse);
      }
      this.update(root);
    },
    
    // Reset view to center
    resetView: function() {
      if (!svg || !root) return;

      const nodes = root.descendants();
      let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
      nodes.forEach(d => {
        minX = Math.min(minX, d.x);
        maxX = Math.max(maxX, d.x);
        minY = Math.min(minY, d.y);
        maxY = Math.max(maxY, d.y);
      });

      const container = document.getElementById('treeContainer');
      const width = container.clientWidth - config.margin.left - config.margin.right;
      const height = container.clientHeight - config.margin.top - config.margin.bottom;

      const dx = maxX - minX;
      const dy = maxY - minY;
      const scale = Math.min(width / (dy || width), height / (dx || height));
      const tx = -minY * scale + (width - dy * scale) / 2 + config.margin.left;
      const ty = -minX * scale + (height - dx * scale) / 2 + config.margin.top;

      svg.transition()
        .duration(750)
        .call(zoom.transform, d3.zoomIdentity.translate(tx, ty).scale(scale));
    },

    // Download current graph data as CSV
    downloadCSV: function() {
      if (!orgData || orgData.length === 0) return;

      // Determine which users to export
      let dataToExport = orgData;
      if (highlightedNodes && highlightedNodes.size > 0) {
        dataToExport = orgData.filter(u => highlightedNodes.has(u.email));
      }

      const header = ['Name', 'Title', 'Department', 'Location', 'Email', 'License', 'AI Usage'];
      const rows = dataToExport.map(u => {
        const rating = u.aiEngagement !== undefined ? GraphAPI.getStarRating(u.aiEngagement) : null;
        const daily = u.aiEngagement !== undefined ? Math.round(u.aiEngagement * 3) : '';
        const aiUsage = rating ? `${rating.label} (~${daily}/day)` : '';
        return [
          u.name || '',
          u.title || '',
          u.department || '',
          u.location || '',
          u.email || '',
          u.hasLicense ? 'Yes' : 'No',
          aiUsage
        ];
      });

      const csv = [header, ...rows]
        .map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(','))
        .join('\n');

      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'org-data.csv';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    },

    // Update spacing
    updateSpacing: function(type, value) {
      if (type === 'horizontal') {
        config.horizontalSpacing = parseInt(value);
        document.getElementById('horizontalValue').textContent = value;
      } else if (type === 'vertical') {
        config.verticalSpacing = parseFloat(value);
        document.getElementById('verticalValue').textContent = value;
      }
      
      if (root) {
        this.update(root);
      }
    },
    
    // Update visualization with new color scheme
    updateVisualization: function() {
      if (root) {
        this.update(root);
        this.updateLegend();
      }
    },
    
    // Highlight department nodes
    highlightDepartment: function(department) {
      // If clicking the same department, clear the highlight
      if (highlightedDepartment === department) {
        this.clearHighlight();
      } else {
        highlightedDepartment = department;
        highlightedNodes = new Set(orgData
          .filter(u => u.department === department)
          .map(u => u.email));
        highlightedNodes = highlightedNodes.size > 0 ? highlightedNodes : null;
        const input = document.getElementById('searchInput');
        if (input) input.value = '';
        const searchClearBtn = document.getElementById('clearSearchBtn');
        if (searchClearBtn) searchClearBtn.style.display = 'none';
        licenseHighlightActive = false;
        const licenseBtn = document.getElementById('licenseToggleBtn');
        if (licenseBtn) licenseBtn.classList.remove('active');

        // Update active class on rows
        document.querySelectorAll('.dept-row').forEach(row => {
          if (row.dataset.department === department) {
            row.classList.add('active');
          } else {
            row.classList.remove('active');
          }
        });

        // Show clear highlight button
        const clearBtn = document.getElementById('clearHighlightBtn');
        if (clearBtn) {
          clearBtn.style.display = 'flex';
        }
      }

      // Update the visualization
      if (root) {
        this.update(root);
      }
    },

    // Toggle highlight of licensed users
    toggleLicenseHighlight: function() {
      if (licenseHighlightActive) {
        this.clearHighlight();
      } else {
        highlightedDepartment = null;
        highlightedNodes = new Set(orgData.filter(u => u.hasLicense).map(u => u.email));
        highlightedNodes = highlightedNodes.size > 0 ? highlightedNodes : null;
        licenseHighlightActive = true;
        document.querySelectorAll('.dept-row').forEach(row => row.classList.remove('active'));
        const input = document.getElementById('searchInput');
        if (input) input.value = '';
        const clearBtn = document.getElementById('clearHighlightBtn');
        if (clearBtn) clearBtn.style.display = 'flex';
        const searchClearBtn = document.getElementById('clearSearchBtn');
        if (searchClearBtn) searchClearBtn.style.display = 'none';
        const licenseBtn = document.getElementById('licenseToggleBtn');
        if (licenseBtn) licenseBtn.classList.add('active');
        if (root) this.update(root);
      }
    },

    // Search by name or title
    searchByNameTitle: function() {
      const input = document.getElementById('searchInput');
      if (!input) return;
      const raw = input.value.trim().toLowerCase();
      if (!raw) {
        this.clearHighlight();
        return;
      }

      // Split on quoted phrases or individual words
      const terms = raw.match(/"[^"]+"|\S+/g) || [];

      highlightedDepartment = null;
      highlightedNodes = new Set(orgData.filter(u => {
        const name = (u.name || '').toLowerCase();
        const title = (u.title || '').toLowerCase();
        return terms.some(t => {
          const token = t.replace(/^"|"$/g, '');
          return name.includes(token) || title.includes(token);
        });
      }).map(u => u.email));

      highlightedNodes = highlightedNodes.size > 0 ? highlightedNodes : null;
      licenseHighlightActive = false;
      const licenseBtn = document.getElementById('licenseToggleBtn');
      if (licenseBtn) licenseBtn.classList.remove('active');

      // Remove active class from department rows
      document.querySelectorAll('.dept-row').forEach(row => row.classList.remove('active'));

      // Show clear search button if we found matches
      const clearHighlightBtn = document.getElementById('clearHighlightBtn');
      if (clearHighlightBtn) {
        clearHighlightBtn.style.display = 'none';
      }
      const clearSearchBtn = document.getElementById('clearSearchBtn');
      if (clearSearchBtn) {
        clearSearchBtn.style.display = highlightedNodes && highlightedNodes.size > 0 ? 'flex' : 'none';
      }

      if (root) {
        this.update(root);
      }
    },

    // Clear highlight (department or search)
    clearHighlight: function() {
      highlightedDepartment = null;
      highlightedNodes = null;
      licenseHighlightActive = false;

      // Remove active class from all rows
      document.querySelectorAll('.dept-row').forEach(row => {
        row.classList.remove('active');
      });

      // Clear search input
      const input = document.getElementById('searchInput');
      if (input) input.value = '';

      // Hide clear buttons
      const clearHighlightBtn = document.getElementById('clearHighlightBtn');
      if (clearHighlightBtn) {
        clearHighlightBtn.style.display = 'none';
      }
      const clearSearchBtn = document.getElementById('clearSearchBtn');
      if (clearSearchBtn) {
        clearSearchBtn.style.display = 'none';
      }
      const licenseBtn = document.getElementById('licenseToggleBtn');
      if (licenseBtn) licenseBtn.classList.remove('active');

      // Update the visualization
      if (root) {
        this.update(root);
      }
    },

    // Get node opacity based on highlight
    getNodeOpacity: function(d) {
      if (!highlightedNodes) {
        return 1;
      }
      return highlightedNodes.has(d.data.email) ? 1 : 0.2;
    },

    // Get link opacity based on highlight
    getLinkOpacity: function(d) {
      if (!highlightedNodes) {
        return 1;
      }
      // Show link if either source or target is in highlighted set
      return (highlightedNodes.has(d.source.data.email) ||
              highlightedNodes.has(d.target.data.email)) ? 0.6 : 0.1;
    },

    // Get node label with optional engagement stars
    getNodeLabel: function(d) {
      const name = d.data.name || 'Unknown';
      if (config.showEngagementStars && d.data.aiEngagement !== undefined && d.data.aiEngagement > 0) {
        const rating = GraphAPI.getStarRating(d.data.aiEngagement);
        return `${name} ${rating.stars}`;
      }
      return name;
    },

    // Get node color based on current color scheme
    getNodeColor: function(d) {
      const colorBy = document.getElementById('colorBy').value;
      
      if (colorBy === 'license') {
        return d.data.hasLicense ? '#00a854' : '#f0f0f0';
      } else if (colorBy === 'level') {
        return levelColorScale(Math.min(d.depth, 5));
      } else { // department
        return departmentColorScale(d.data.department || 'Unknown');
      }
    },
    
    // Get node stroke color
    getNodeStroke: function(d) {
      // Red border for no license, green for has license
      if (d.data.hasLicense) {
        return '#00a854'; // Green for licensed users
      }
      return '#dc3545'; // Red for unlicensed users
    },
    
    // Show tooltip
    showTooltip: function(event, d) {
      const tooltip = document.getElementById('tooltip');
      const licenseText = d.data.hasLicense ? 
        '<span style="color: #00a854; font-weight: bold;">✔ Has ChatGPT License</span>' : 
        '<span style="color: #dc3545; font-weight: bold;">✗ No ChatGPT License</span>';
      
      // Calculate AI engagement display
      let engagementText = '';
      if (d.data.aiEngagement !== undefined) {
        const rating = GraphAPI.getStarRating(d.data.aiEngagement);
        const dailyEngagements = Math.round(d.data.aiEngagement * 3);
        engagementText = `
          <div style="margin-top: 0.5rem; padding-top: 0.5rem; border-top: 1px solid #ddd;">
            <strong>AI Engagement:</strong> ${rating.stars}<br>
            <span style="color: var(--text-light); font-size: 0.85rem;">
              ${rating.label} (~${dailyEngagements} engagements/day)
            </span>
          </div>
        `;
      }
      
      tooltip.innerHTML = `
        <strong>${d.data.name || 'Unknown'}</strong><br>
        ${d.data.title ? d.data.title + '<br>' : ''}
        ${d.data.department || 'No department'}<br>
        ${d.data.location ? d.data.location + '<br>' : ''}
        ${d.data.email ? d.data.email + '<br>' : ''}
        <div style="margin-top: 0.5rem; padding-top: 0.5rem; border-top: 1px solid #ddd;">
          ${licenseText}
        </div>
        ${engagementText}
      `;
      
      tooltip.style.left = (event.pageX + 10) + 'px';
      tooltip.style.top = (event.pageY - 10) + 'px';
      tooltip.classList.add('visible');
    },
    
    // Hide tooltip
    hideTooltip: function() {
      document.getElementById('tooltip').classList.remove('visible');
    },
    
    // Update legend based on color scheme
    updateLegend: function() {
      const legend = document.getElementById('legend');
      const colorBy = document.getElementById('colorBy').value;
      
      let html = '<h4>Legend</h4>';
      
      // Always show license border legend
      html += `
        <div style="margin-bottom: 1rem; padding-bottom: 0.5rem; border-bottom: 1px solid var(--border);">
          <strong>License Status (Border)</strong>
          <div class="legend-item">
            <div class="legend-color" style="background: white; border: 3px solid #00a854;"></div>
            <span>Has ChatGPT License</span>
          </div>
          <div class="legend-item">
            <div class="legend-color" style="background: white; border: 3px solid #dc3545;"></div>
            <span>No License</span>
          </div>
        </div>
      `;
      
      // Add color-specific legend
      html += '<strong>Node Color</strong>';
      
      if (colorBy === 'license') {
        html += `
          <div class="legend-item">
            <div class="legend-color" style="background: #00a854;"></div>
            <span>Has License (Fill)</span>
          </div>
          <div class="legend-item">
            <div class="legend-color" style="background: #f0f0f0;"></div>
            <span>No License (Fill)</span>
          </div>
        `;
      } else if (colorBy === 'level') {
        for (let i = 0; i <= 4; i++) {
          html += `
            <div class="legend-item">
              <div class="legend-color" style="background: ${levelColorScale(i)};"></div>
              <span>Level ${i}</span>
            </div>
          `;
        }
      } else { // department
        const departments = [...new Set(orgData.map(u => u.department))].slice(0, 10);
        departments.forEach(dept => {
          html += `
            <div class="legend-item">
              <div class="legend-color" style="background: ${departmentColorScale(dept)};"></div>
              <span>${dept}</span>
            </div>
          `;
        });
        if (departments.length < [...new Set(orgData.map(u => u.department))].length) {
          html += '<div class="legend-item"><small>...and more</small></div>';
        }
      }
      
      legend.innerHTML = html;
    },
    
    // Update statistics panel
    updateStatsPanel: function() {
      const stats = document.getElementById('statsPanel');
      const licensed = orgData.filter(u => u.hasLicense).length;
      const departments = [...new Set(orgData.map(u => u.department))].length;
      const maxDepth = root ? d3.max(root.descendants(), d => d.depth) : 0;
      
      stats.innerHTML = `
        <div class="stat-item">
          <span class="stat-label">Total Users:</span>
          <span class="stat-value">${orgData.length}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Licensed:</span>
          <span class="stat-value">${licensed} (${Math.round(licensed/orgData.length*100)}%)</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Departments:</span>
          <span class="stat-value">${departments}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Max Depth:</span>
          <span class="stat-value">${maxDepth}</span>
        </div>
      `;
    }
  };
})();

// Expose OrgChart to global scope for inline event handlers
window.OrgChart = OrgChart;