<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="includes/includes.asp" -->
<%
Response.CodePage = 65001
Response.Charset = "UTF-8"

Dim db, bpSvc, rs
Set db = New cls_db
db.Open
Set bpSvc = New cls_blueprint

Set rs = bpSvc.GetAllBlueprints(db, "date", "", "", "")

' Don't close db here - will close at end of page
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Timeline — Idea Blueprint</title>
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
            max-width: 800px;
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
            margin-bottom: 40px;
        }
        
        .timeline {
            position: relative;
            padding-left: 30px;
        }
        
        .timeline::before {
            content: '';
            position: absolute;
            left: 8px;
            top: 0;
            bottom: 0;
            width: 2px;
            background: #e5e7eb;
        }
        
        .timeline-month {
            margin-bottom: 32px;
        }
        
        .month-label {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 14px;
            font-weight: 600;
            color: #6b7280;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 16px;
            position: relative;
        }
        
        .month-label::before {
            content: '';
            position: absolute;
            left: -26px;
            top: 50%;
            transform: translateY(-50%);
            width: 12px;
            height: 12px;
            background: #fff;
            border: 3px solid #6366f1;
            border-radius: 50%;
        }
        
        .timeline-entries {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }
        
        .timeline-entry {
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
        
        .timeline-entry:hover {
            border-color: #6366f1;
            box-shadow: 0 4px 15px rgba(99, 102, 241, 0.1);
        }
        
        .spark-dot {
            width: 16px;
            height: 16px;
            border-radius: 50%;
            flex-shrink: 0;
        }
        
        .entry-content {
            flex: 1;
        }
        
        .entry-name {
            font-weight: 600;
            font-size: 15px;
            margin-bottom: 4px;
        }
        
        .entry-liner {
            font-size: 13px;
            color: #6b7280;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        
        .entry-meta {
            display: flex;
            gap: 12px;
            align-items: center;
        }
        
        .entry-category {
            font-size: 12px;
            color: #6366f1;
        }
        
        .entry-status {
            font-size: 11px;
            padding: 3px 8px;
            border-radius: 12px;
            background: #f3f4f6;
            color: #6b7280;
        }
        
        .empty-state {
            text-align: center;
            padding: 60px 20px;
            color: #6b7280;
        }
        
        .empty-state .icon {
            font-size: 48px;
            margin-bottom: 16px;
            opacity: 0.5;
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
                <a href="timeline.asp" class="active">Timeline</a>
                <a href="stats.asp">Stats</a>
            </nav>
        </header>
        
        <h1>Timeline</h1>
        <p class="subtitle">Your idea journey over time</p>
        
        <% 
        Dim groupedData, monthKey, entries, bp
        Set groupedData = CreateObject("Scripting.Dictionary")
        
        Do While Not rs.EOF
            Dim createdDate, monthStr
            createdDate = rs("created_at")
            If createdDate <> "" Then
                monthStr = Year(CDate(createdDate)) & "-" & Right("0" & Month(CDate(createdDate)), 2)
                
                If Not groupedData.Exists(monthStr) Then
                    Set entries = CreateObject("Scripting.Dictionary")
                    groupedData.Add monthStr, entries
                Else
                    Set entries = groupedData(monthStr)
                End If
                
                Dim entryKey
                entryKey = CStr(rs("id"))
                entries.Add entryKey, rs
            End If
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        Dim allMonths, m
        allMonths = groupedData.Keys
        
        If groupedData.Count > 0 Then
            For Each monthStr In allMonths
                Set entries = groupedData(monthStr)
                Dim monthDate, monthLabel
                monthDate = CDate(monthStr & "-01")
                monthLabel = FormatDateTime(monthDate, 1)
                monthLabel = Left(monthLabel, InStr(monthLabel, " ") - 1) & " " & Year(monthDate)
        %>
        <div class="timeline-month">
            <div class="month-label"><%=monthLabel%></div>
            <div class="timeline-entries">
                <%
                Dim entry
                For Each entry In entries.Keys
                    Set bp = entries(entry)
                %>
                <a href="detail.asp?id=<%=bp("id")%>" class="timeline-entry">
                    <div class="spark-dot" style="background:<%=GetSparkColor(bp("spark_score"))%>"></div>
                    <div class="entry-content">
                        <div class="entry-name"><%=H(bp("name"))%></div>
                        <% If bp("one_liner") <> "" Then %>
                        <div class="entry-liner"><%=H(bp("one_liner"))%></div>
                        <% End If %>
                    </div>
                    <div class="entry-meta">
                        <span class="entry-category"><%=GetCategoryLabel(bp("category"))%></span>
                        <span class="entry-status"><%=GetStatusLabel(bp("status"))%></span>
                    </div>
                </a>
                <% Next %>
            </div>
        </div>
        <% 
            Next
        Else %>
        <div class="empty-state">
            <div class="icon">📅</div>
            <p>No blueprints yet. Create your first one!</p>
        </div>
        <% End If %>
    </div>
</body>
</html>
<%
If Not rs Is Nothing Then rs.Close: Set rs = Nothing
If Not db Is Nothing Then db.Close: Set db = Nothing
%>
