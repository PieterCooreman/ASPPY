<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="includes/includes.asp" -->
<%
Response.CodePage = 65001
Response.Charset = "UTF-8"

Dim db, bpSvc, blueprintId, rs
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

If IsPost() Then
    Dim name, oneLiner, problem, userTarget, hook, risk, sparkScore, category, tags, status, notes
    
    name = Trim("" & Request.Form("name"))
    oneLiner = Trim("" & Request.Form("one_liner"))
    problem = Trim("" & Request.Form("problem"))
    userTarget = Trim("" & Request.Form("user_target"))
    hook = Trim("" & Request.Form("hook"))
    risk = Trim("" & Request.Form("risk"))
    sparkScore = ToInt(Request.Form("spark_score"), 3)
    category = Trim("" & Request.Form("category"))
    tags = Trim("" & Request.Form("tags"))
    status = Trim("" & Request.Form("status"))
    notes = Trim("" & Request.Form("notes"))
    
    If name <> "" Then
        bpSvc.UpdateBlueprint db, blueprintId, name, oneLiner, problem, userTarget, hook, risk, sparkScore, category, tags, status, notes
        Response.Redirect "detail.asp?id=" & blueprintId
        Response.End
    End If
End If

' Don't close db here - will close at end of page
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Edit <%=H(rs("name"))%> — Idea Blueprint</title>
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
            max-width: 700px;
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
        
        h1 {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 24px;
            font-weight: 600;
            margin-bottom: 24px;
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
            background: #fff;
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
        
        .form-actions {
            display: flex;
            justify-content: space-between;
            margin-top: 32px;
        }
        
        .btn {
            padding: 12px 24px;
            border-radius: 10px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            font-family: inherit;
            text-decoration: none;
        }
        
        .btn-secondary {
            background: #f3f4f6;
            border: none;
            color: #374151;
            display: inline-flex;
            align-items: center;
        }
        
        .btn-primary {
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            border: none;
            color: #fff;
        }
        
        .btn-primary:hover {
            box-shadow: 0 4px 15px rgba(99, 102, 241, 0.4);
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <a href="detail.asp?id=<%=blueprintId%>" class="back-link">← Back to blueprint</a>
        </header>
        
        <h1>Edit Blueprint</h1>
        
        <form method="post">
            <div class="form-group">
                <label>Name *</label>
                <input type="text" name="name" maxlength="60" value="<%=H(rs("name"))%>" required>
            </div>
            
            <div class="form-group">
                <label>One-liner</label>
                <input type="text" name="one_liner" maxlength="120" value="<%=H(rs("one_liner"))%>">
            </div>
            
            <div class="form-group">
                <label>The Problem</label>
                <textarea name="problem" maxlength="300"><%=H(rs("problem"))%></textarea>
            </div>
            
            <div class="form-group">
                <label>The User</label>
                <input type="text" name="user_target" maxlength="200" value="<%=H(rs("user_target"))%>">
            </div>
            
            <div class="form-group">
                <label>The Hook</label>
                <textarea name="hook" maxlength="300"><%=H(rs("hook"))%></textarea>
            </div>
            
            <div class="form-group">
                <label>The Risk</label>
                <textarea name="risk" maxlength="200"><%=H(rs("risk"))%></textarea>
            </div>
            
            <div class="form-group">
                <label>Spark Score</label>
                <div class="spark-select">
                    <div class="spark-option <%If rs("spark_score")=1 Then Response.Write("selected")%>" data-score="1">☆</div>
                    <div class="spark-option <%If rs("spark_score")=2 Then Response.Write("selected")%>" data-score="2">☆☆</div>
                    <div class="spark-option <%If rs("spark_score")=3 Then Response.Write("selected")%>" data-score="3">☆☆☆</div>
                    <div class="spark-option <%If rs("spark_score")=4 Then Response.Write("selected")%>" data-score="4">☆☆☆☆</div>
                    <div class="spark-option <%If rs("spark_score")=5 Then Response.Write("selected")%>" data-score="5">☆☆☆☆☆</div>
                </div>
                <input type="hidden" name="spark_score" id="sparkScore" value="<%=rs("spark_score")%>">
            </div>
            
            <div class="form-row">
                <div class="form-group">
                    <label>Category</label>
                    <select name="category">
                        <option value="app" <%If rs("category")="app" Then Response.Write("selected")%>>📱 App</option>
                        <option value="tool" <%If rs("category")="tool" Then Response.Write("selected")%>>🛠️ Tool</option>
                        <option value="game" <%If rs("category")="game" Then Response.Write("selected")%>>🎮 Game</option>
                        <option value="web" <%If rs("category")="web" Then Response.Write("selected")%>>🌐 Web</option>
                        <option value="ai" <%If rs("category")="ai" Then Response.Write("selected")%>>🤖 AI</option>
                        <option value="business" <%If rs("category")="business" Then Response.Write("selected")%>>🏪 Business</option>
                        <option value="creative" <%If rs("category")="creative" Then Response.Write("selected")%>>🎨 Creative</option>
                        <option value="social" <%If rs("category")="social" Then Response.Write("selected")%>>🌍 Social</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>Status</label>
                    <select name="status">
                        <option value="idea" <%If rs("status")="idea" Then Response.Write("selected")%>>💡 Idea</option>
                        <option value="exploring" <%If rs("status")="exploring" Then Response.Write("selected")%>>🔨 Exploring</option>
                        <option value="building" <%If rs("status")="building" Then Response.Write("selected")%>>🚀 Building</option>
                        <option value="shipped" <%If rs("status")="shipped" Then Response.Write("selected")%>>📦 Shipped</option>
                        <option value="shelved" <%If rs("status")="shelved" Then Response.Write("selected")%>>🗄️ Shelved</option>
                    </select>
                </div>
            </div>
            
            <div class="form-group">
                <label>Tags</label>
                <input type="text" name="tags" value="<%=H(rs("tags"))%>" placeholder="maps, anonymous, real-time">
            </div>
            
            <div class="form-group">
                <label>Private Notes</label>
                <textarea name="notes" style="min-height: 150px;"><%=H(rs("notes"))%></textarea>
            </div>
            
            <div class="form-actions">
                <a href="detail.asp?id=<%=blueprintId%>" class="btn btn-secondary">Cancel</a>
                <button type="submit" class="btn btn-primary">Save Changes</button>
            </div>
        </form>
    </div>
    
    <script>
        document.querySelectorAll('.spark-option').forEach(opt => {
            opt.addEventListener('click', () => {
                document.querySelectorAll('.spark-option').forEach(o => o.classList.remove('selected'));
                opt.classList.add('selected');
                document.getElementById('sparkScore').value = opt.dataset.score;
            });
        });
    </script>
</body>
</html>
<%
If Not rs Is Nothing Then rs.Close: Set rs = Nothing
If Not db Is Nothing Then db.Close: Set db = Nothing
%>
