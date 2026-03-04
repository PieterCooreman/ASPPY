<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="includes/includes.asp" -->
<%
Response.CodePage = 65001
Response.Charset = "UTF-8"

Dim db, bpSvc, blueprintId, rs, rsHistory
Set db = New cls_db
db.Open
Set bpSvc = New cls_blueprint

blueprintId = Request.QueryString("id")

If blueprintId = "" Then
    Response.Redirect "default.asp"
    Response.End
End If

Set rs = bpSvc.GetBlueprintById(db, blueprintId)

If rs.EOF Then
    Response.Redirect "default.asp"
    Response.End
End If

Set rsHistory = bpSvc.GetSparkHistory(db, blueprintId)

' Don't close db here - will close at end of page
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title><%=H(rs("name"))%> — Idea Blueprint</title>
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
            max-width: 900px;
            margin: 0 auto;
            padding: 24px;
        }
        
        header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 32px;
        }
        
        .back-link {
            text-decoration: none;
            color: #6b7280;
            font-size: 14px;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: color 0.2s;
        }
        
        .back-link:hover {
            color: #1a1a1a;
        }
        
        .actions {
            display: flex;
            gap: 8px;
        }
        
        .action-btn {
            padding: 8px 14px;
            background: #f3f4f6;
            border: none;
            border-radius: 8px;
            font-size: 13px;
            cursor: pointer;
            color: #374151;
            transition: all 0.2s;
        }
        
        .action-btn:hover {
            background: #e5e7eb;
        }
        
        .action-btn.danger:hover {
            background: #fee2e2;
            color: #dc2626;
        }
        
        .blueprint-header {
            background: #fff;
            border-radius: 16px;
            padding: 32px;
            margin-bottom: 24px;
            border: 1px solid #e5e7eb;
        }
        
        .bp-meta {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 16px;
        }
        
        .bp-category {
            font-size: 14px;
            color: #6366f1;
            font-weight: 500;
        }
        
        .bp-status {
            font-size: 13px;
            padding: 6px 14px;
            border-radius: 20px;
            background: #f3f4f6;
            color: #6b7280;
        }
        
        .bp-name {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 32px;
            font-weight: 700;
            margin-bottom: 12px;
            color: #1a1a1a;
        }
        
        .bp-liner {
            font-size: 18px;
            color: #6b7280;
            line-height: 1.6;
            margin-bottom: 20px;
        }
        
        .bp-spark {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        
        .bp-spark-label {
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: #9ca3af;
        }
        
        .bp-spark-stars {
            font-size: 24px;
            letter-spacing: 4px;
        }
        
        .blueprint-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
        }
        
        @media (max-width: 700px) {
            .blueprint-grid {
                grid-template-columns: 1fr;
            }
        }
        
        .bp-section {
            background: #fff;
            border-radius: 14px;
            padding: 24px;
            border: 1px solid #e5e7eb;
        }
        
        .bp-section.full-width {
            grid-column: 1 / -1;
        }
        
        .bp-section h3 {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 1px;
            color: #9ca3af;
            margin-bottom: 12px;
        }
        
        .bp-section p {
            font-size: 15px;
            line-height: 1.7;
            color: #374151;
            white-space: pre-wrap;
        }
        
        .bp-section.empty p {
            color: #d1d5db;
            font-style: italic;
        }
        
        .tags-list {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }
        
        .tag {
            padding: 6px 12px;
            background: #f3f4f6;
            border-radius: 20px;
            font-size: 13px;
            color: #6b7280;
        }
        
        .spark-history {
            margin-top: 16px;
        }
        
        .spark-history h4 {
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: #9ca3af;
            margin-bottom: 10px;
        }
        
        .spark-timeline {
            display: flex;
            flex-direction: column;
            gap: 8px;
        }
        
        .spark-entry {
            display: flex;
            align-items: center;
            gap: 12px;
            font-size: 13px;
        }
        
        .spark-entry .date {
            color: #9ca3af;
            min-width: 80px;
        }
        
        .spark-entry .stars {
            letter-spacing: 2px;
        }
        
        .revisit-btn {
            margin-top: 16px;
            padding: 10px 20px;
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border: none;
            border-radius: 8px;
            color: #fff;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }
        
        .revisit-btn:hover {
            box-shadow: 0 4px 15px rgba(99, 102, 241, 0.4);
        }
        
        .revisit-form {
            display: none;
            margin-top: 16px;
            padding: 16px;
            background: #f9fafb;
            border-radius: 10px;
        }
        
        .revisit-form.visible {
            display: block;
        }
        
        .spark-select {
            display: flex;
            gap: 8px;
            margin-bottom: 12px;
        }
        
        .spark-option {
            flex: 1;
            padding: 8px;
            border: 2px solid #e5e7eb;
            border-radius: 6px;
            text-align: center;
            cursor: pointer;
            transition: all 0.2s;
            font-size: 16px;
        }
        
        .spark-option:hover, .spark-option.selected {
            border-color: #6366f1;
            background: #eef2ff;
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
            max-width: 500px;
            padding: 24px;
            animation: slideUp 0.3s ease-out;
        }
        
        @keyframes slideUp {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        .modal h3 {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 18px;
            margin-bottom: 16px;
        }
        
        .modal-btns {
            display: flex;
            justify-content: flex-end;
            gap: 12px;
            margin-top: 20px;
        }
        
        .btn {
            padding: 10px 20px;
            border-radius: 8px;
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
        
        .btn-danger {
            background: #fee2e2;
            border: none;
            color: #dc2626;
        }
        
        .btn-primary {
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border: none;
            color: #fff;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <a href="default.asp" class="back-link">← Back to vault</a>
            <div class="actions">
                <button class="action-btn" id="editBtn">Edit</button>
                <button class="action-btn danger" id="deleteBtn">Delete</button>
            </div>
        </header>
        
        <div class="blueprint-header">
            <div class="bp-meta">
                <span class="bp-category"><%=GetCategoryLabel(rs("category"))%></span>
                <span class="bp-status"><%=GetStatusLabel(rs("status"))%></span>
            </div>
            <h1 class="bp-name"><%=H(rs("name"))%></h1>
            <% If rs("one_liner") <> "" Then %>
            <p class="bp-liner"><%=H(rs("one_liner"))%></p>
            <% End If %>
            <div class="bp-spark">
                <span class="bp-spark-label">Spark Score</span>
                <span class="bp-spark-stars" style="color:<%=GetSparkColor(rs("spark_score"))%>"><%=GetSparkStars(rs("spark_score"))%></span>
            </div>
        </div>
        
        <div class="blueprint-grid">
            <div class="bp-section <%If rs("problem") = "" Then Response.Write("empty")%>">
                <h3>The Problem</h3>
                <p><%If rs("problem") <> "" Then Response.Write(H(rs("problem"))) Else Response.Write("Not specified")End If%></p>
            </div>
            
            <div class="bp-section <%If rs("user_target") = "" Then Response.Write("empty")%>">
                <h3>The User</h3>
                <p><%If rs("user_target") <> "" Then Response.Write(H(rs("user_target"))) Else Response.Write("Not specified")End If%></p>
            </div>
            
            <div class="bp-section <%If rs("hook") = "" Then Response.Write("empty")%>">
                <h3>The Hook</h3>
                <p><%If rs("hook") <> "" Then Response.Write(H(rs("hook"))) Else Response.Write("Not specified")End If%></p>
            </div>
            
            <div class="bp-section <%If rs("risk") = "" Then Response.Write("empty")%>">
                <h3>The Risk</h3>
                <p><%If rs("risk") <> "" Then Response.Write(H(rs("risk"))) Else Response.Write("Not specified")End If%></p>
            </div>
            
            <div class="bp-section full-width">
                <h3>Private Notes</h3>
                <p><%If rs("notes") <> "" Then Response.Write(H(rs("notes"))) Else Response.Write("No notes yet. Click Edit to add some.")%></p>
            </div>
            
            <div class="bp-section">
                <h3>Tags</h3>
                <% If rs("tags") <> "" Then %>
                <div class="tags-list">
                    <% 
                    Dim tagsArr, t
                    tagsArr = Split(rs("tags"), ",")
                    For t = 0 To UBound(tagsArr)
                        If Trim(tagsArr(t)) <> "" Then %>
                    <span class="tag"><%=H(Trim(tagsArr(t)))%></span>
                        <% End If
                    Next %>
                </div>
                <% Else %>
                <p style="color:#d1d5db;font-style:italic;">No tags</p>
                <% End If %>
            </div>
            
            <div class="bp-section">
                <h3>Revisit History</h3>
                <div class="spark-history">
                    <% 
                    Dim historyCount
                    historyCount = 0
                    Do While Not rsHistory.EOF 
                        historyCount = historyCount + 1
                        If historyCount <= 5 Then %>
                    <div class="spark-entry">
                        <span class="date"><%=FormatDate(rsHistory("recorded_at"))%></span>
                        <span class="stars" style="color:<%=GetSparkColor(rsHistory("score"))%>"><%=GetSparkStars(rsHistory("score"))%></span>
                    </div>
                        <% End If
                        rsHistory.MoveNext
                    Loop
                    rsHistory.Close
                    Set rsHistory = Nothing
                    %>
                </div>
                
                <button class="revisit-btn" id="revisitBtn">Revisit Spark Score</button>
                
                <div class="revisit-form" id="revisitForm">
                    <div class="spark-select" id="newSparkSelect">
                        <div class="spark-option" data-score="1">☆</div>
                        <div class="spark-option" data-score="2">☆☆</div>
                        <div class="spark-option selected" data-score="3">☆☆☆</div>
                        <div class="spark-option" data-score="4">☆☆☆☆</div>
                        <div class="spark-option" data-score="5">☆☆☆☆☆</div>
                    </div>
                    <button class="btn btn-primary" id="saveRevisitBtn">Save</button>
                </div>
            </div>
        </div>
    </div>
    
    <div class="modal-overlay" id="deleteModal">
        <div class="modal">
            <h3>Delete Blueprint?</h3>
            <p>This action cannot be undone. Are you sure you want to delete "<%=H(rs("name"))%>"?</p>
            <div class="modal-btns">
                <button class="btn btn-secondary" id="cancelDelete">Cancel</button>
                <button class="btn btn-danger" id="confirmDelete">Delete</button>
            </div>
        </div>
    </div>
    
    <script>
        const blueprintId = '<%=blueprintId%>';
        let newSpark = <%=rs("spark_score")%>;
        
        document.getElementById('editBtn').addEventListener('click', () => {
            window.location.href = 'edit.asp?id=' + blueprintId;
        });
        
        document.getElementById('deleteBtn').addEventListener('click', () => {
            document.getElementById('deleteModal').classList.add('visible');
        });
        
        document.getElementById('cancelDelete').addEventListener('click', () => {
            document.getElementById('deleteModal').classList.remove('visible');
        });
        
        document.getElementById('confirmDelete').addEventListener('click', async () => {
            try {
                await fetch('api/blueprints.asp?id=' + blueprintId, {
                    method: 'DELETE'
                });
                window.location.href = 'default.asp';
            } catch (err) {
                alert('Failed to delete. Please try again.');
            }
        });
        
        document.getElementById('revisitBtn').addEventListener('click', () => {
            document.getElementById('revisitForm').classList.toggle('visible');
        });
        
        document.querySelectorAll('#newSparkSelect .spark-option').forEach(opt => {
            opt.addEventListener('click', () => {
                newSpark = parseInt(opt.dataset.score);
                document.querySelectorAll('#newSparkSelect .spark-option').forEach(o => {
                    o.classList.toggle('selected', parseInt(o.dataset.score) <= newSpark);
                });
            });
        });
        
        document.getElementById('saveRevisitBtn').addEventListener('click', async () => {
            try {
                const formData = new FormData();
                formData.append('score', newSpark);
                
                await fetch('api/blueprints.asp?action=revisit&id=' + blueprintId, {
                    method: 'POST',
                    body: formData
                });
                
                location.reload();
            } catch (err) {
                alert('Failed to save. Please try again.');
            }
        });
    </script>
</body>
</html>
<%
If Not rs Is Nothing Then rs.Close: Set rs = Nothing
If Not rsHistory Is Nothing Then rsHistory.Close: Set rsHistory = Nothing
If Not db Is Nothing Then db.Close: Set db = Nothing
%>
