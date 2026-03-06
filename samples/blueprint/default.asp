<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="includes/includes.asp" -->
<%
Response.CodePage = 65001
Response.Charset = "UTF-8"

Dim db, bpSvc, sortBy, search, category, status, rs, totalCount
Set db = New cls_db
db.Open
Set bpSvc = New cls_blueprint

sortBy = Request.QueryString("sort")
search = Trim("" & Request.QueryString("q"))
category = Trim("" & Request.QueryString("category"))
status = Trim("" & Request.QueryString("status"))

If sortBy = "" Then sortBy = "date"

totalCount = 0
Set rs = bpSvc.GetAllBlueprints(db, sortBy, search, category, status)
Do While Not rs.EOF
    totalCount = totalCount + 1
    rs.MoveNext
Loop
rs.MoveFirst

' Don't close db here - will close at end of page
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Idea Blueprint</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500&display=swap" rel="stylesheet">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Inter', sans-serif;
            background: #fafafa;
            color: #1a1a1a;
            min-height: 100vh;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 24px;
        }
        
        header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 32px;
        }
        
        .logo {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 24px;
            font-weight: 700;
            color: #1a1a1a;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .logo-icon {
            width: 36px;
            height: 36px;
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
        }
        
        .nav-links {
            display: flex;
            gap: 16px;
        }
        
        .nav-links a {
            text-decoration: none;
            color: #6b7280;
            font-size: 14px;
            font-weight: 500;
            padding: 8px 16px;
            border-radius: 8px;
            transition: all 0.2s;
        }
        
        .nav-links a:hover, .nav-links a.active {
            background: #f3f4f6;
            color: #1a1a1a;
        }
        
        .toolbar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 24px;
            flex-wrap: wrap;
            gap: 16px;
        }
        
        .search-box {
            flex: 1;
            max-width: 400px;
            position: relative;
        }
        
        .search-box input {
            width: 100%;
            padding: 12px 16px 12px 44px;
            border: 2px solid #e5e7eb;
            border-radius: 10px;
            font-size: 14px;
            font-family: inherit;
            transition: border-color 0.2s;
            background: #fff;
        }
        
        .search-box input:focus {
            outline: none;
            border-color: #6366f1;
        }
        
        .search-box::before {
            content: "🔍";
            position: absolute;
            left: 14px;
            top: 50%;
            transform: translateY(-50%);
            font-size: 16px;
        }
        
        .filters {
            display: flex;
            gap: 12px;
        }
        
        .filters select {
            padding: 10px 14px;
            border: 2px solid #e5e7eb;
            border-radius: 8px;
            font-size: 13px;
            font-family: inherit;
            background: #fff;
            cursor: pointer;
        }
        
        .add-btn {
            padding: 12px 24px;
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border: none;
            border-radius: 10px;
            color: #fff;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .add-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 15px rgba(99, 102, 241, 0.4);
        }
        
        .stats-bar {
            display: flex;
            gap: 24px;
            margin-bottom: 24px;
            padding: 16px 20px;
            background: #fff;
            border-radius: 12px;
            border: 1px solid #e5e7eb;
        }
        
        .stat-item {
            text-align: center;
        }
        
        .stat-item .value {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 24px;
            font-weight: 700;
            color: #1a1a1a;
        }
        
        .stat-item .label {
            font-size: 12px;
            color: #6b7280;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .blueprints-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
            gap: 20px;
        }
        
        .blueprint-card {
            background: #fff;
            border: 1px solid #e5e7eb;
            border-radius: 14px;
            padding: 20px;
            cursor: pointer;
            transition: all 0.2s;
            text-decoration: none;
            color: inherit;
            display: block;
        }
        
        .blueprint-card:hover {
            border-color: #6366f1;
            box-shadow: 0 4px 20px rgba(99, 102, 241, 0.15);
            transform: translateY(-2px);
        }
        
        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 12px;
        }
        
        .card-category {
            font-size: 12px;
            color: #6366f1;
            font-weight: 500;
        }
        
        .card-spark {
            font-size: 14px;
            letter-spacing: 2px;
        }
        
        .card-name {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 8px;
            color: #1a1a1a;
        }
        
        .card-liner {
            font-size: 13px;
            color: #6b7280;
            line-height: 1.5;
            margin-bottom: 12px;
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            overflow: hidden;
        }
        
        .card-footer {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .card-status {
            font-size: 12px;
            padding: 4px 10px;
            border-radius: 20px;
            background: #f3f4f6;
            color: #6b7280;
        }
        
        .card-date {
            font-size: 11px;
            color: #9ca3af;
        }
        
        .empty-state {
            text-align: center;
            padding: 80px 20px;
            color: #6b7280;
        }
        
        .empty-state .icon {
            font-size: 64px;
            margin-bottom: 20px;
            opacity: 0.5;
        }
        
        .empty-state h3 {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 20px;
            margin-bottom: 8px;
            color: #1a1a1a;
        }
        
        .modal-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            z-index: 1000;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        
        .modal-overlay.visible {
            display: flex;
        }
        
        .modal {
            background: #fff;
            border-radius: 16px;
            width: 100%;
            max-width: 600px;
            max-height: 90vh;
            overflow-y: auto;
            animation: slideUp 0.3s ease-out;
        }
        
        @keyframes slideUp {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        .modal-header {
            padding: 24px;
            border-bottom: 1px solid #e5e7eb;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .modal-header h2 {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 20px;
            font-weight: 600;
        }
        
        .modal-close {
            width: 32px;
            height: 32px;
            border: none;
            background: #f3f4f6;
            border-radius: 8px;
            cursor: pointer;
            font-size: 18px;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        
        .modal-body {
            padding: 24px;
        }
        
        .form-group {
            margin-bottom: 20px;
        }
        
        .form-group label {
            display: block;
            font-size: 12px;
            font-weight: 600;
            color: #374151;
            margin-bottom: 6px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .form-group input, .form-group textarea, .form-group select {
            width: 100%;
            padding: 12px 14px;
            border: 2px solid #e5e7eb;
            border-radius: 10px;
            font-size: 14px;
            font-family: inherit;
            transition: border-color 0.2s;
        }
        
        .form-group input:focus, .form-group textarea:focus, .form-group select:focus {
            outline: none;
            border-color: #6366f1;
        }
        
        .form-group textarea {
            resize: vertical;
            min-height: 80px;
        }
        
        .form-hint {
            font-size: 11px;
            color: #9ca3af;
            margin-top: 4px;
        }
        
        .spark-select {
            display: flex;
            gap: 8px;
        }
        
        .spark-option {
            flex: 1;
            padding: 10px;
            border: 2px solid #e5e7eb;
            border-radius: 8px;
            text-align: center;
            cursor: pointer;
            transition: all 0.2s;
            font-size: 18px;
        }
        
        .spark-option:hover, .spark-option.selected {
            border-color: #6366f1;
            background: #eef2ff;
        }
        
        .form-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 16px;
        }
        
        .modal-footer {
            padding: 20px 24px;
            border-top: 1px solid #e5e7eb;
            display: flex;
            justify-content: flex-end;
            gap: 12px;
        }
        
        .btn {
            padding: 12px 24px;
            border-radius: 10px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            font-family: inherit;
        }
        
        .btn-secondary {
            background: #f3f4f6;
            border: none;
            color: #374151;
        }
        
        .btn-primary {
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border: none;
            color: #fff;
        }
        
        .btn-primary:hover {
            box-shadow: 0 4px 15px rgba(99, 102, 241, 0.4);
        }
        
        .btn-danger {
            background: #fee2e2;
            border: none;
            color: #dc2626;
        }
        
        .export-btns {
            display: flex;
            gap: 8px;
        }
        
        .export-btn {
            padding: 8px 12px;
            background: #f3f4f6;
            border: none;
            border-radius: 6px;
            font-size: 12px;
            cursor: pointer;
            color: #6b7280;
            transition: all 0.2s;
        }
        
        .export-btn:hover {
            background: #e5e7eb;
            color: #1a1a1a;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="logo">
                <div class="logo-icon">📐</div>
                <span>Idea Blueprint</span>
            </div>
            <nav class="nav-links">
                <a href="default.asp" class="active">Vault</a>
                <a href="timeline.asp">Timeline</a>
                <a href="stats.asp">Stats</a>
            </nav>
        </header>
        
        <div class="toolbar">
            <div class="search-box">
                <input type="text" id="searchInput" placeholder="Search blueprints..." value="<%=H(search)%>">
            </div>
            <div class="filters">
                <select id="categoryFilter">
                    <option value="all">All categories</option>
                    <option value="tool" <%If category="tool" Then Response.Write("selected")%>>🛠️ Tool</option>
                    <option value="game" <%If category="game" Then Response.Write("selected")%>>🎮 Game</option>
                    <option value="app" <%If category="app" Then Response.Write("selected")%>>📱 App</option>
                    <option value="web" <%If category="web" Then Response.Write("selected")%>>🌐 Web</option>
                    <option value="ai" <%If category="ai" Then Response.Write("selected")%>>🤖 AI</option>
                    <option value="business" <%If category="business" Then Response.Write("selected")%>>🏪 Business</option>
                    <option value="creative" <%If category="creative" Then Response.Write("selected")%>>🎨 Creative</option>
                    <option value="social" <%If category="social" Then Response.Write("selected")%>>🌍 Social</option>
                </select>
                <select id="sortFilter">
                    <option value="date" <%If sortBy="date" Then Response.Write("selected")%>>Newest</option>
                    <option value="spark" <%If sortBy="spark" Then Response.Write("selected")%>>Spark score</option>
                    <option value="name" <%If sortBy="name" Then Response.Write("selected")%>>Name</option>
                </select>
            </div>
            <button class="add-btn" id="addBtn">+ New Blueprint</button>
        </div>
        
        <div class="stats-bar">
            <div class="stat-item">
                <div class="value" id="totalCount"><%=totalCount%></div>
                <div class="label">Total Ideas</div>
            </div>
            <div class="stat-item">
                <div class="value">5</div>
                <div class="label">Categories</div>
            </div>
        </div>
        
        <div class="blueprints-grid" id="blueprintsGrid">
            <% If Not rs.EOF Then
                Do While Not rs.EOF %>
            <a href="detail.asp?id=<%=rs("id")%>" class="blueprint-card">
                <div class="card-header">
                    <span class="card-category"><%=GetCategoryLabel(rs("category"))%></span>
                    <span class="card-spark" style="color:<%=GetSparkColor(rs("spark_score"))%>"><%=GetSparkStars(rs("spark_score"))%></span>
                </div>
                <div class="card-name"><%=H(rs("name"))%></div>
                <div class="card-liner"><%=H(rs("one_liner"))%></div>
                <div class="card-footer">
                    <span class="card-status"><%=GetStatusLabel(rs("status"))%></span>
                    <span class="card-date"><%=FormatDate(rs("created_at"))%></span>
                </div>
            </a>
                <% rs.MoveNext
                Loop
            Else %>
            <div class="empty-state" style="grid-column: 1 / -1;">
                <div class="icon">💡</div>
                <h3>No blueprints yet</h3>
                <p>Create your first idea blueprint to get started.</p>
            </div>
            <% End If %>
            <% rs.Close: Set rs = Nothing %>
        </div>
    </div>
    
    <div class="modal-overlay" id="modalOverlay">
        <div class="modal">
            <div class="modal-header">
                <h2 id="modalTitle">New Blueprint</h2>
                <button class="modal-close" id="modalClose">×</button>
            </div>
            <div class="modal-body">
                <form id="blueprintForm">
                    <input type="hidden" id="blueprintId" value="">
                    
                    <div class="form-group">
                        <label>Name *</label>
                        <input type="text" id="bpName" maxlength="60" placeholder="What do you call it?" required>
                        <div class="form-hint">Max 60 characters</div>
                    </div>
                    
                    <div class="form-group">
                        <label>One-liner</label>
                        <input type="text" id="bpOneLiner" maxlength="120" placeholder="Explain it to a stranger in one sentence">
                        <div class="form-hint">Max 120 characters</div>
                    </div>
                    
                    <div class="form-group">
                        <label>The Problem</label>
                        <textarea id="bpProblem" maxlength="300" placeholder="What frustration does it solve?"></textarea>
                        <div class="form-hint">Max 300 characters</div>
                    </div>
                    
                    <div class="form-group">
                        <label>The User</label>
                        <input type="text" id="bpUserTarget" maxlength="200" placeholder="Who is this for, exactly?">
                    </div>
                    
                    <div class="form-group">
                        <label>The Hook</label>
                        <textarea id="bpHook" maxlength="300" placeholder="What makes it unique or memorable?"></textarea>
                    </div>
                    
                    <div class="form-group">
                        <label>The Risk</label>
                        <textarea id="bpRisk" maxlength="200" placeholder="What could kill it? Be honest."></textarea>
                    </div>
                    
                    <div class="form-group">
                        <label>Spark Score</label>
                        <div class="spark-select" id="sparkSelect">
                            <div class="spark-option" data-score="1">☆</div>
                            <div class="spark-option" data-score="2">☆☆</div>
                            <div class="spark-option selected" data-score="3">☆☆☆</div>
                            <div class="spark-option" data-score="4">☆☆☆☆</div>
                            <div class="spark-option" data-score="5">☆☆☆☆☆</div>
                        </div>
                    </div>
                    
                    <div class="form-row">
                        <div class="form-group">
                            <label>Category</label>
                            <select id="bpCategory">
                                <option value="app">📱 App</option>
                                <option value="tool">🛠️ Tool</option>
                                <option value="game">🎮 Game</option>
                                <option value="web">🌐 Web</option>
                                <option value="ai">🤖 AI</option>
                                <option value="business">🏪 Business</option>
                                <option value="creative">🎨 Creative</option>
                                <option value="social">🌍 Social</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label>Status</label>
                            <select id="bpStatus">
                                <option value="idea">💡 Idea</option>
                                <option value="exploring">🔨 Exploring</option>
                                <option value="building">🚀 Building</option>
                                <option value="shipped">📦 Shipped</option>
                                <option value="shelved">🗄️ Shelved</option>
                            </select>
                        </div>
                    </div>
                    
                    <div class="form-group">
                        <label>Tags</label>
                        <input type="text" id="bpTags" placeholder="maps, anonymous, real-time (comma separated)">
                    </div>
                    
                    <div class="form-group">
                        <label>Private Notes</label>
                        <textarea id="bpNotes" placeholder="Follow-up thinking, links, sketches..." style="min-height: 120px;"></textarea>
                    </div>
                </form>
            </div>
            <div class="modal-footer">
                <button class="btn btn-secondary" id="cancelBtn">Cancel</button>
                <button class="btn btn-primary" id="saveBtn">Save Blueprint</button>
            </div>
        </div>
    </div>
    
    <script>
        const modal = document.getElementById('modalOverlay');
        const form = document.getElementById('blueprintForm');
        let selectedSpark = 3;
        
        document.getElementById('addBtn').addEventListener('click', () => {
            document.getElementById('modalTitle').textContent = 'New Blueprint';
            document.getElementById('blueprintId').value = '';
            form.reset();
            selectedSpark = 3;
            updateSparkDisplay();
            modal.classList.add('visible');
        });
        
        document.getElementById('modalClose').addEventListener('click', () => {
            modal.classList.remove('visible');
        });
        
        document.getElementById('cancelBtn').addEventListener('click', () => {
            modal.classList.remove('visible');
        });
        
        modal.addEventListener('click', (e) => {
            if (e.target === modal) modal.classList.remove('visible');
        });
        
        document.querySelectorAll('.spark-option').forEach(opt => {
            opt.addEventListener('click', () => {
                selectedSpark = parseInt(opt.dataset.score);
                updateSparkDisplay();
            });
        });
        
        function updateSparkDisplay() {
            document.querySelectorAll('.spark-option').forEach(opt => {
                opt.classList.toggle('selected', parseInt(opt.dataset.score) <= selectedSpark);
            });
        }
        
        document.getElementById('saveBtn').addEventListener('click', async () => {
            const name = document.getElementById('bpName').value.trim();
            if (!name) {
                alert('Name is required');
                return;
            }
            
            const formData = new FormData();
            formData.append('name', name);
            formData.append('one_liner', document.getElementById('bpOneLiner').value);
            formData.append('problem', document.getElementById('bpProblem').value);
            formData.append('user_target', document.getElementById('bpUserTarget').value);
            formData.append('hook', document.getElementById('bpHook').value);
            formData.append('risk', document.getElementById('bpRisk').value);
            formData.append('spark_score', selectedSpark);
            formData.append('category', document.getElementById('bpCategory').value);
            formData.append('status', document.getElementById('bpStatus').value);
            formData.append('tags', document.getElementById('bpTags').value);
            formData.append('notes', document.getElementById('bpNotes').value);
            
            try {
                const response = await fetch('api/blueprints.asp', {
                    method: 'POST',
                    body: formData
                });
                
                const data = await response.json();
                
                if (data.error) {
                    alert(data.error);
                    return;
                }
                
                modal.classList.remove('visible');
                location.reload();
                
            } catch (err) {
                alert('Failed to save. Please try again.');
            }
        });
        
        const searchInput = document.getElementById('searchInput');
        const categoryFilter = document.getElementById('categoryFilter');
        const sortFilter = document.getElementById('sortFilter');
        
        function applyFilters() {
            const params = new URLSearchParams();
            if (searchInput.value) params.set('q', searchInput.value);
            if (categoryFilter.value !== 'all') params.set('category', categoryFilter.value);
            if (sortFilter.value !== 'date') params.set('sort', sortFilter.value);
            window.location.href = '?' + params.toString();
        }
        
        searchInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') applyFilters();
        });
        
        categoryFilter.addEventListener('change', applyFilters);
        sortFilter.addEventListener('change', applyFilters);
    </script>
</body>
</html>
<%
If Not rs Is Nothing Then rs.Close: Set rs = Nothing
If Not db Is Nothing Then db.Close: Set db = Nothing
%>
