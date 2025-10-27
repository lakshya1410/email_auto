// Import Chart.js
import Chart from 'chart.js/auto';

// Global variables
let categoryChart, priorityChart, sentimentChart, activityChart;

// Wait for Office to be ready before initializing
Office.onReady((info) => {
  console.log('Office.js ready!', info);
  
  // Make sure we're in the right context
  if (info.host === Office.HostType.Outlook) {
    console.log('Running in Outlook');
  }
  
  // Initialize UI
  const createTicketBtn = document.getElementById("createTicket");
  const copySummaryBtn = document.getElementById("copySummary");
  const copyReplyBtn = document.getElementById("copyReply");
  const refreshTicketsBtn = document.getElementById("refreshTickets");
  const statusFilter = document.getElementById("statusFilter");
  const categoryFilter = document.getElementById("categoryFilter");
  const priorityFilter = document.getElementById("priorityFilter");
  
  if (createTicketBtn) {
    createTicketBtn.onclick = createTicketFromEmail;
  }
  
  if (copySummaryBtn) {
    copySummaryBtn.onclick = () => copyToClipboard('summary');
  }
  
  if (copyReplyBtn) {
    copyReplyBtn.onclick = () => copyToClipboard('reply');
  }
  
  if (refreshTicketsBtn) {
    refreshTicketsBtn.onclick = loadTickets;
  }
  
  if (statusFilter) {
    statusFilter.onchange = loadTickets;
  }
  
  if (categoryFilter) {
    categoryFilter.onchange = loadTickets;
  }
  
  if (priorityFilter) {
    priorityFilter.onchange = loadTickets;
  }
  
  // Setup tab navigation
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      const viewName = tab.dataset.view;
      switchView(viewName);
    });
  });
  
  setStatus("✅ Ready");
}).catch(err => {
  console.error('Office.onReady failed:', err);
  setStatus("❌ Error initializing Office.js", "error");
});

function setStatus(message, type = "normal") {
  const statusEl = document.getElementById("status");
  statusEl.innerText = message;
  
  // Update status styling based on type
  statusEl.className = "status-bar";
  if (type === "error") {
    statusEl.className = "status-bar error";
  } else if (type === "success") {
    statusEl.className = "status-bar success";
  }
}

async function createTicketFromEmail() {
  try {
    // Disable button and show loading
    const btn = document.getElementById("createTicket");
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner"></span> Creating Ticket...';
    
    // Hide previous results
    document.getElementById("badgesSection").style.display = "none";
    document.getElementById("keyPointsSection").style.display = "none";
    document.getElementById("copySummary").style.display = "none";
    document.getElementById("copyReply").style.display = "none";
    document.getElementById("ticketCreatedSection").style.display = "none";
    
    setStatus("📧 Fetching email content...");
    const item = Office.context.mailbox.item;

    // Get email body text (plain text)
    const body = await new Promise((resolve, reject) => {
      item.body.getAsync("text", (res) => {
        if (res.status === Office.AsyncResultStatus.Failed) {
          reject(res.error);
        } else {
          resolve(res.value || "");
        }
      });
    });

    setStatus("🎫 Creating support ticket...");
    
    // Get sender and subject
    const senderInfo = item.from?.emailAddress || item.sender?.emailAddress || "unknown@email.com";
    const senderName = item.from?.displayName || item.sender?.displayName || "Unknown User";
    const subject = item.subject || "No Subject";
    
    // Using HTTP for localhost development (HTTPS certificates don't work in Outlook Desktop)
    const resp = await fetch("http://localhost:8001/api/tickets/create", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ 
        sender_email: senderInfo,
        sender_name: senderName,
        subject: subject,
        body: body
      })
    });

    if (!resp.ok) {
      const t = await resp.text();
      throw new Error(`Backend error: ${resp.status} ${t}`);
    }

    const data = await resp.json();
    
    // Extract ticket data (the analysis is inside the ticket object)
    const ticket = data.ticket || data;
    
    // Display ticket created section
    const ticketCreatedSection = document.getElementById("ticketCreatedSection");
    const ticketNumberDisplay = document.getElementById("ticketNumberDisplay");
    const ticketStatusDisplay = document.getElementById("ticketStatusDisplay");
    const confirmationStatus = document.getElementById("confirmationStatus");
    
    ticketNumberDisplay.textContent = data.ticket_number || ticket.ticket_number;
    ticketStatusDisplay.textContent = `Status: ${(data.status || ticket.status).toUpperCase()}`;
    confirmationStatus.innerHTML = data.confirmation_sent ? '✓ Confirmation email sent to customer' : '⚠️ Ticket created (email not sent)';
    ticketCreatedSection.style.display = "block";
    
    // Display all analysis results (summary and reply are in the ticket object)
    document.getElementById("summary").innerText = ticket.summary || "No summary generated";
    document.getElementById("reply").innerText = ticket.suggested_reply || "No reply generated";
    
    // Display key points (need to parse JSON string)
    const keyPointsEl = document.getElementById("keyPoints");
    const keyPointsSection = document.getElementById("keyPointsSection");
    
    let keyPoints = [];
    if (ticket.key_points) {
      try {
        // key_points is stored as JSON string in database
        keyPoints = typeof ticket.key_points === 'string' ? JSON.parse(ticket.key_points) : ticket.key_points;
      } catch (e) {
        console.error("Failed to parse key points:", e);
      }
    }
    
    if (keyPoints && keyPoints.length > 0) {
      const ul = document.createElement("ul");
      keyPoints.forEach(point => {
        const li = document.createElement("li");
        li.textContent = point;
        ul.appendChild(li);
      });
      keyPointsEl.innerHTML = "";
      keyPointsEl.appendChild(ul);
      keyPointsSection.style.display = "block";
    }
    
    // Display badges
    const badgesSection = document.getElementById("badgesSection");
    const priorityBadge = document.getElementById("priorityBadge");
    const categoryBadge = document.getElementById("categoryBadge");
    const sentimentBadge = document.getElementById("sentimentBadge");
    
    if (ticket.priority) {
      const priorityClass = `badge-${ticket.priority.toLowerCase()}`;
      priorityBadge.className = `badge ${priorityClass}`;
      const priorityEmoji = ticket.priority === "High" ? "🔴" : ticket.priority === "Medium" ? "🟡" : "🟢";
      priorityBadge.innerHTML = `${priorityEmoji} ${ticket.priority}`;
      priorityBadge.style.display = "inline-block";
    }
    
    if (ticket.category) {
      const catIcon = getCategoryIcon(ticket.category);
      categoryBadge.innerHTML = `${catIcon} ${ticket.category}`;
      categoryBadge.style.display = "inline-block";
    }
    
    if (ticket.sentiment_tone) {
      const sentimentIcon = getSentimentIcon(ticket.sentiment_tone);
      sentimentBadge.innerHTML = `${sentimentIcon} ${ticket.sentiment_tone}`;
      sentimentBadge.style.display = "inline-block";
    }
    
    badgesSection.style.display = "flex";
    
    // Show copy buttons
    document.getElementById("copySummary").style.display = "inline-block";
    document.getElementById("copyReply").style.display = "inline-block";
    
    // Reset button
    btn.disabled = false;
    btn.innerHTML = '<span style="font-size: 1.2em;">➕</span> Create Ticket from Email';
    
    setStatus("✅ Ticket created successfully!", "success");
    
  } catch (error) {
    console.error("Error creating ticket:", error);
    setStatus(`❌ Error: ${error.message}`, "error");
    
    // Re-enable button
    const btn = document.getElementById("createTicket");
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span style="font-size: 1.2em;">➕</span> Create Ticket from Email';
    }
  }
}

async function copyToClipboard(type) {
  try {
    const textElement = document.getElementById(type);
    const text = textElement.innerText;
    
    // Copy to clipboard
    await navigator.clipboard.writeText(text);
    
    // Visual feedback
    const btn = document.getElementById(`copy${type.charAt(0).toUpperCase() + type.slice(1)}`);
    const originalHTML = btn.innerHTML;
    
    btn.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
        <path d="M13.854 3.646a.5.5 0 0 1 0 .708l-7 7a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L6.5 10.293l6.646-6.647a.5.5 0 0 1 .708 0z"/>
      </svg>
      Copied!
    `;
    btn.style.background = "#107c10";
    
    // Reset after 2 seconds
    setTimeout(() => {
      btn.innerHTML = originalHTML;
      btn.style.background = "";
    }, 2000);
    
  } catch (err) {
    console.error('Copy failed:', err);
    setStatus("❌ Failed to copy to clipboard", "error");
  }
}

// ============================================
// DASHBOARD FUNCTIONS
// ============================================

function switchView(viewName) {
  // Update tabs
  document.querySelectorAll('.tab').forEach(tab => {
    if (tab.dataset.view === viewName) {
      tab.classList.add('active');
    } else {
      tab.classList.remove('active');
    }
  });
  
  // Update views
  document.querySelectorAll('.view').forEach(view => {
    view.classList.remove('active');
  });
  document.getElementById(`${viewName}View`).classList.add('active');
  
  // Load data for the view
  if (viewName === 'dashboard') {
    loadDashboard();
  } else if (viewName === 'tickets') {
    loadTickets();
  }
}

async function loadDashboard() {
  try {
    const resp = await fetch("http://localhost:8001/api/tickets/stats/dashboard");
    if (!resp.ok) throw new Error(`Stats API error: ${resp.status}`);
    
    const stats = await resp.json();
    
    // Update stat cards
    document.getElementById('totalAnalyzed').textContent = stats.total_tickets || 0;
    
    // Calculate this week count
    const thisWeekCount = (stats.recent_activity || []).reduce((sum, day) => sum + day.count, 0);
    document.getElementById('thisWeek').textContent = thisWeekCount;
    
    // Render charts
    renderCategoryChart(stats.by_category || {});
    renderPriorityChart(stats.by_priority || {});
    renderSentimentChart(stats.by_status || {});  // Changed from by_sentiment to by_status
    renderActivityChart(stats.recent_activity || []);
    
  } catch (err) {
    console.error('Failed to load dashboard:', err);
  }
}

function renderCategoryChart(data) {
  const ctx = document.getElementById('categoryChart');
  if (!ctx) return;
  
  // Destroy existing chart
  if (categoryChart) {
    categoryChart.destroy();
  }
  
  const labels = Object.keys(data);
  const values = Object.values(data);
  
  if (labels.length === 0) {
    ctx.getContext('2d').clearRect(0, 0, ctx.width, ctx.height);
    return;
  }
  
  categoryChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Emails',
        data: values,
        backgroundColor: ['#0078d4', '#107c10', '#faa300', '#d13438', '#8764b8']
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: {
          display: false
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: {
            stepSize: 1
          }
        }
      }
    }
  });
}

function renderPriorityChart(data) {
  const ctx = document.getElementById('priorityChart');
  if (!ctx) return;
  
  if (priorityChart) {
    priorityChart.destroy();
  }
  
  const labels = Object.keys(data);
  const values = Object.values(data);
  
  if (labels.length === 0) {
    ctx.getContext('2d').clearRect(0, 0, ctx.width, ctx.height);
    return;
  }
  
  priorityChart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{
        data: values,
        backgroundColor: ['#d13438', '#faa300', '#107c10']
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: {
          position: 'bottom'
        }
      }
    }
  });
}

function renderSentimentChart(data) {
  const ctx = document.getElementById('sentimentChart');
  if (!ctx) return;
  
  if (sentimentChart) {
    sentimentChart.destroy();
  }
  
  const labels = Object.keys(data);
  const values = Object.values(data);
  
  if (labels.length === 0) {
    ctx.getContext('2d').clearRect(0, 0, ctx.width, ctx.height);
    return;
  }
  
  // Color mapping for ticket status
  const statusColors = {
    'open': '#22c55e',
    'in-progress': '#f59e0b',
    'closed': '#6b7280'
  };
  const colors = labels.map(label => statusColors[label.toLowerCase()] || '#0078d4');
  
  sentimentChart = new Chart(ctx, {
    type: 'pie',
    data: {
      labels: labels.map(l => l.charAt(0).toUpperCase() + l.slice(1).replace('-', ' ')),
      datasets: [{
        data: values,
        backgroundColor: colors
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: {
          position: 'bottom'
        }
      }
    }
  });
}

function renderActivityChart(data) {
  const ctx = document.getElementById('activityChart');
  if (!ctx) return;
  
  if (activityChart) {
    activityChart.destroy();
  }
  
  if (data.length === 0) {
    ctx.getContext('2d').clearRect(0, 0, ctx.width, ctx.height);
    return;
  }
  
  const labels = data.map(d => {
    const date = new Date(d.date);
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  });
  const values = data.map(d => d.count);
  
  activityChart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [{
        label: 'Analyses',
        data: values,
        borderColor: '#0078d4',
        backgroundColor: 'rgba(0, 120, 212, 0.1)',
        tension: 0.3,
        fill: true
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: {
          display: false
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: {
            stepSize: 1
          }
        }
      }
    }
  });
}

async function loadTickets() {
  try {
    const statusFilter = document.getElementById('statusFilter')?.value || '';
    const categoryFilter = document.getElementById('categoryFilter')?.value || '';
    const priorityFilter = document.getElementById('priorityFilter')?.value || '';
    
    let url = "http://localhost:8001/api/tickets?limit=50";
    if (statusFilter) url += `&status=${statusFilter}`;
    if (categoryFilter) url += `&category=${categoryFilter}`;
    if (priorityFilter) url += `&priority=${priorityFilter}`;
    
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Tickets API error: ${resp.status}`);
    
    const data = await resp.json();
    const ticketsList = document.getElementById('ticketsList');
    
    // Backend returns 'items' not 'tickets'
    const tickets = data.items || data.tickets || [];
    
    if (tickets.length === 0) {
      ticketsList.innerHTML = `
        <div class="empty-state">
          <div class="empty-icon">🎫</div>
          <div class="empty-text">No tickets yet</div>
          <div class="empty-subtext">Create a ticket from an email to get started</div>
        </div>
      `;
      return;
    }
    
    ticketsList.innerHTML = tickets.map(ticket => {
      const date = new Date(ticket.created_at);
      const dateStr = date.toLocaleDateString('en-US', { 
        month: 'short', 
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });
      
      const statusColors = {
        'open': '#22c55e',
        'in-progress': '#f59e0b',
        'closed': '#6b7280'
      };
      const statusColor = statusColors[ticket.status] || '#6b7280';
      
      const priorityClass = `badge-${ticket.priority?.toLowerCase() || 'medium'}`;
      const priorityEmoji = ticket.priority === "High" ? "🔴" : ticket.priority === "Medium" ? "🟡" : "🟢";
      
      const categoryEmojis = {
        "Sales": "💰",
        "Support": "🛠️",
        "General": "📧",
        "Marketing": "📢",
        "HR": "👥"
      };
      const categoryEmoji = categoryEmojis[ticket.category] || "📁";
      
      return `
        <div class="history-item" onclick="viewTicket('${ticket.ticket_number}')">
          <div class="history-header">
            <div>
              <strong style="color: #0078d4;">🎫 ${ticket.ticket_number}</strong>
              <span style="display: inline-block; padding: 2px 8px; border-radius: 12px; background: ${statusColor}; color: white; font-size: 0.75em; margin-left: 8px;">
                ${ticket.status.toUpperCase()}
              </span>
            </div>
            <div class="history-date">${dateStr}</div>
          </div>
          <div class="history-subject"><strong>From:</strong> ${ticket.sender_name || ticket.sender_email}</div>
          <div class="history-subject"><strong>Subject:</strong> ${ticket.subject}</div>
          <div class="history-summary">${ticket.summary || 'No summary available'}</div>
          <div class="history-badges">
            <span class="badge ${priorityClass}">${priorityEmoji} ${ticket.priority || 'Medium'}</span>
            <span class="badge">${categoryEmoji} ${ticket.category || 'General'}</span>
          </div>
        </div>
      `;
    }).join('');
    
  } catch (error) {
    console.error("Error loading tickets:", error);
    const ticketsList = document.getElementById('ticketsList');
    ticketsList.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">❌</div>
        <div class="empty-text">Error loading tickets</div>
        <div class="empty-subtext">${error.message}</div>
      </div>
    `;
  }
}

function viewTicket(ticketNumber) {
  // Switch to create view and populate with ticket data
  fetch(`http://localhost:8001/api/tickets/${ticketNumber}`)
    .then(resp => resp.json())
    .then(ticket => {
      // Switch to create view
      switchView('create');
      
      // Show ticket created section
      const ticketCreatedSection = document.getElementById("ticketCreatedSection");
      const ticketNumberDisplay = document.getElementById("ticketNumberDisplay");
      const ticketStatusDisplay = document.getElementById("ticketStatusDisplay");
      
      ticketNumberDisplay.textContent = ticket.ticket_number;
      ticketStatusDisplay.textContent = `Status: ${ticket.status.toUpperCase()}`;
      ticketCreatedSection.style.display = "block";
      
      // Populate the fields
      document.getElementById("summary").innerText = ticket.summary || "";
      document.getElementById("reply").innerText = ticket.suggested_reply || "";
      
      // Display key points
      const keyPointsEl = document.getElementById("keyPoints");
      const keyPointsSection = document.getElementById("keyPointsSection");
      if (ticket.key_points && ticket.key_points.length > 0) {
        const ul = document.createElement("ul");
        ticket.key_points.forEach(point => {
          const li = document.createElement("li");
          li.textContent = point;
          ul.appendChild(li);
        });
        keyPointsEl.innerHTML = "";
        keyPointsEl.appendChild(ul);
        keyPointsSection.style.display = "block";
      }
      
      // Display badges
      const badgesSection = document.getElementById("badgesSection");
      const priorityBadge = document.getElementById("priorityBadge");
      const categoryBadge = document.getElementById("categoryBadge");
      const sentimentBadge = document.getElementById("sentimentBadge");
      
      if (ticket.priority) {
        const priorityClass = `badge-${ticket.priority.toLowerCase()}`;
        priorityBadge.className = `badge ${priorityClass}`;
        const priorityEmoji = ticket.priority === "High" ? "🔴" : ticket.priority === "Medium" ? "🟡" : "🟢";
        priorityBadge.innerHTML = `${priorityEmoji} ${ticket.priority}`;
        priorityBadge.style.display = "inline-block";
      }
      
      if (ticket.category) {
        const categoryIcon = getCategoryIcon(ticket.category);
        categoryBadge.innerHTML = `${categoryIcon} ${ticket.category}`;
        categoryBadge.style.display = "inline-block";
      }
      
      if (ticket.sentiment) {
        const sentimentIcon = getSentimentIcon(ticket.sentiment);
        sentimentBadge.innerHTML = `${sentimentIcon} ${ticket.sentiment}`;
        sentimentBadge.style.display = "inline-block";
      }
      
      badgesSection.style.display = "flex";
      
      // Show copy buttons
      document.getElementById("copySummary").style.display = "inline-block";
      document.getElementById("copyReply").style.display = "inline-block";
      
      setStatus(`🎫 Viewing ticket ${ticket.ticket_number} from ${new Date(ticket.created_at).toLocaleString()}`, "success");
    })
    .catch(err => {
      console.error("Error loading ticket:", err);
      setStatus(`❌ Error loading ticket: ${err.message}`, "error");
    });
}

function getCategoryIcon(category) {
  const categoryEmojis = {
    "Sales": "💰",
    "Support": "🛠️",
    "General": "📧",
    "Marketing": "📢",
    "HR": "👥"
  };
  return categoryEmojis[category] || "📁";
}

function getSentimentIcon(sentiment) {
  const sentimentEmojis = {
    "Positive": "😊",
    "Neutral": "😐",
    "Negative": "😟",
    "Urgent": "⚠️"
  };
  return sentimentEmojis[sentiment] || "😐";
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}
