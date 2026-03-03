<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="../includes/includes.asp" -->
<%
Response.CodePage = 65001
Response.Charset = "UTF-8"
Response.ContentType = "application/json"
Response.AddHeader "Access-Control-Allow-Origin", "*"

Dim db, bpSvc, stats, rsTimeline, timelineData(), i
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

Response.Write "{"
Response.Write """total"":" & stats(0) & ","
Response.Write """avg_spark"":" & Round(CDbl(stats(1)), 1) & ","
Response.Write """most_used_tag"":""" & H(stats(2)) & ""","
Response.Write """idea_count"":" & stats(3) & ","
Response.Write """shelved_count"":" & stats(4) & ","
Response.Write """shipped_count"":" & stats(5) & ","
Response.Write """timeline"":["

Dim j
For j = 0 To i - 1
    If j > 0 Then Response.Write ","
    Response.Write "{"
    Response.Write """month"":""" & timelineData(j)(0) & ""","
    Response.Write """count"":" & timelineData(j)(1) & ","
    Response.Write """avg_spark"":" & Round(timelineData(j)(2), 1)
    Response.Write "}"
Next

Response.Write "]"
Response.Write "}"

db.Close
Set db = Nothing
%>
