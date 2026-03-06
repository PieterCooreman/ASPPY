<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="includes/includes.asp" -->
<%
Response.CodePage = 65001
Response.Charset = "UTF-8"

Dim db, bpSvc, stats, rsCold, rsTimeline, timelineData(), i
Set db = New cls_db
db.Open
Set bpSvc = New cls_blueprint

stats = bpSvc.GetStats(db)

i = 0
Set rsTimeline = bpSvc.GetTimelineData(db)
Do While Not rsTimeline.EOF
    ReDim Preserve timelineData(i)
    timelineData(i) = Array("" & rsTimeline("month"), CLng(Nz(rsTimeline("count"), 0)), CDbl(Nz(rsTimeline("avg_spark"), 0)))
    rsTimeline.MoveNext
    i = i + 1
Loop
rsTimeline.Close
Set rsTimeline = Nothing

Set rsCold = bpSvc.GetColdIdeas(db)

' Don't close db here - will close at end of page
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Stats — Idea Blueprint</title>
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
            margin-bottom: 40px;
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
        
        h1 {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 28px;
            font-weight: 700;
            margin-bottom: 8px;
        }
        
        .subtitle {
            color: #6b7280;
            margin-bottom: 32px;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }
        
        .stat-card {
            background: #fff;
            border: 1px solid #e5e7eb;
            border-radius: 14px;
            padding: 24px;
            text-align: center;
        }
        
        .stat-value {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 36px;
            font-weight: 700;
            color: #1a1a1a;
            margin-bottom: 4px;
        }
        
        .stat-label {
            font-size: 13px;
            color: #6b7280;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .section {
            margin-bottom: 40px;
        }
        
        .section h2 {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 18px;
            font-weight: 600;
            margin-bottom: 16px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .breakdown {
            display: flex;
            gap: 16px;
            flex-wrap: wrap;
        }
        
        .breakdown-item {
            background: #fff;
            border: 1px solid #e5e7eb;
            border-radius: 10px;
            padding: 16px 24px;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        
        .breakdown-icon {
            font-size: 24px;
        }
        
        .breakdown-count {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 24px;
            font-weight: 700;
        }
        
        .breakdown-label {
            font-size: 12px;
            color: #6b7280;
        }
        
        .cold-list {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }
        
        .cold-item {
            display: flex;
            align-items: center;
            gap: 16px;
            padding: 16px;
            background: #fff;
            border: 1px solid #e5e7eb;
            border-radius: 12px;
            text-decoration: none;
            color: inherit;
            transition: all 0.2s;
        }
        
        .cold-item:hover {
            border-color: #6366f1;
        }
        
        .cold-spark {
            width: 12px;
            height: 12px;
            border-radius: 50%;
        }
        
        .cold-content {
            flex: 1;
        }
        
        .cold-name {
            font-weight: 600;
            font-size: 14px;
        }
        
        .cold-date {
            font-size: 12px;
            color: #9ca3af;
        }
        
        .chart-container {
            background: #fff;
            border: 1px solid #e5e7eb;
            border-radius: 14px;
            padding: 24px;
            margin-bottom: 24px;
        }
        
        .chart {
            display: flex;
            align-items: flex-end;
            gap: 8px;
            height: 150px;
            padding-top: 20px;
        }
        
        .chart-bar {
            flex: 1;
            background: linear-gradient(to top, #6366f1, #8b5cf6);
            border-radius: 4px 4px 0 0;
            position: relative;
            min-height: 4px;
            transition: all 0.3s;
        }
        
        .chart-bar:hover {
            opacity: 0.8;
        }
        
        .chart-bar .tooltip {
            position: absolute;
            bottom: 100%;
            left: 50%;
            transform: translateX(-50%);
            background: #1a1a1a;
            color: #fff;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 11px;
            white-space: nowrap;
            opacity: 0;
            pointer-events: none;
            transition: opacity 0.2s;
        }
        
        .chart-bar:hover .tooltip {
            opacity: 1;
        }
        
        .chart-labels {
            display: flex;
            gap: 8px;
            margin-top: 8px;
        }
        
        .chart-label {
            flex: 1;
            text-align: center;
            font-size: 10px;
            color: #9ca3af;
        }
        
        .empty-state {
            text-align: center;
            padding: 40px;
            color: #6b7280;
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
                <a href="default.asp">Vault</a>
                <a href="timeline.asp">Timeline</a>
                <a href="stats.asp" class="active">Stats</a>
            </nav>
        </header>
        
        <h1>Stats</h1>
        <p class="subtitle">Your vault at a glance</p>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-value"><%=stats(0)%></div>
                <div class="stat-label">Total Ideas</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%=Round(CDbl(stats(1)), 1)%></div>
                <div class="stat-label">Avg Spark Score</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%=stats(3)%></div>
                <div class="stat-label">Ideas</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%=stats(4)%></div>
                <div class="stat-label">Shelved</div>
            </div>
            <div class="stat-card">
                <div class="stat-value"><%=stats(5)%></div>
                <div class="stat-label">Shipped</div>
            </div>
        </div>
        
        <div class="section">
            <h2>Status Breakdown</h2>
            <div class="breakdown">
                <div class="breakdown-item">
                    <span class="breakdown-icon">💡</span>
                    <div>
                        <div class="breakdown-count"><%=stats(3)%></div>
                        <div class="breakdown-label">Ideas</div>
                    </div>
                </div>
                <div class="breakdown-item">
                    <span class="breakdown-icon">🔨</span>
                    <div>
                        <div class="breakdown-count">-</div>
                        <div class="breakdown-label">Exploring</div>
                    </div>
                </div>
                <div class="breakdown-item">
                    <span class="breakdown-icon">🚀</span>
                    <div>
                        <div class="breakdown-count">-</div>
                        <div class="breakdown-label">Building</div>
                    </div>
                </div>
                <div class="breakdown-item">
                    <span class="breakdown-icon">📦</span>
                    <div>
                        <div class="breakdown-count"><%=stats(5)%></div>
                        <div class="breakdown-label">Shipped</div>
                    </div>
                </div>
                <div class="breakdown-item">
                    <span class="breakdown-icon">🗄️</span>
                    <div>
                        <div class="breakdown-count"><%=stats(4)%></div>
                        <div class="breakdown-label">Shelved</div>
                    </div>
                </div>
            </div>
        </div>
        
        <% 
        Dim hasTimelineData
        hasTimelineData = False
        On Error Resume Next
        If UBound(timelineData) >= 0 Then hasTimelineData = True
        On Error GoTo 0
        
        If hasTimelineData Then 
        %>
        <div class="section">
            <h2>Monthly Activity</h2>
            <div class="chart-container">
                <div class="chart">
                    <%
                    On Error Resume Next
                    For j = 0 To UBound(timelineData)
                        If timelineData(j)(1) > maxCount Then maxCount = timelineData(j)(1)
                    Next
                    
                    For j = 0 To UBound(timelineData)
                        If maxCount > 0 Then
                            barHeight = (timelineData(j)(1) / maxCount) * 100
                        Else
                            barHeight = 0
                        End If
                    %>
                    <div class="chart-bar" style="height:<%=barHeight%>%">
                        <span class="tooltip"><%=timelineData(j)(1)%> ideas</span>
                    </div>
                    <% Next 
                    On Error GoTo 0
                    %>
                </div>
                <div class="chart-labels">
                    <% For j = 0 To UBound(timelineData) %>
                    <div class="chart-label"><%=Right(timelineData(j)(0), 2)%></div>
                    <% Next %>
                </div>
            </div>
        </div>
        <% End If %>
        
        <%
        Dim coldCount
        coldCount = 0
        Do While Not rsCold.EOF
            coldCount = coldCount + 1
            rsCold.MoveNext
        Loop
        
        If coldCount > 0 Then
            rsCold.MoveFirst
        %>
        <div class="section">
            <h2>Getting Cold 🔒</h2>
            <p style="color:#6b7280;font-size:14px;margin-bottom:16px;">These ideas haven't been revisited in 30+ days</p>
            <div class="cold-list">
                <% 
                Do While Not rsCold.EOF 
                %>
                <a href="detail.asp?id=<%=rsCold("id")%>" class="cold-item">
                    <div class="cold-spark" style="background:<%=GetSparkColor(rsCold("spark_score"))%>"></div>
                    <div class="cold-content">
                        <div class="cold-name"><%=H(rsCold("name"))%></div>
                        <div class="cold-date">Last updated: <%=FormatDate(rsCold("updated_at"))%></div>
                    </div>
                    <span>→</span>
                </a>
                <% 
                    rsCold.MoveNext
                Loop
                rsCold.Close
                Set rsCold = Nothing
                %>
            </div>
        </div>
        <% End If %>
        
        <% If coldCount = 0 And stats(0) = 0 Then %>
        <div class="empty-state">
            <p>No stats yet. Create some blueprints to see your analytics!</p>
        </div>
        <% End If %>
    </div>
</body>
</html>
<%
If Not rsTimeline Is Nothing Then rsTimeline.Close: Set rsTimeline = Nothing
If Not rsCold Is Nothing Then rsCold.Close: Set rsCold = Nothing
If Not db Is Nothing Then db.Close: Set db = Nothing
%>
