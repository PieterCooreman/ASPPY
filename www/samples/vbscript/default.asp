<%@ Language="VBScript" CodePage="65001" %>
<%
Option Explicit

' ============================================================
'  CLASSIC ASP / VBSCRIPT SHOWCASE  —  default.asp
' ============================================================

' ── 1. SERVER-SIDE VARIABLES & ARITHMETIC ──────────────────
Dim pageTitle, visitCount, today
pageTitle  = "Classic ASP Feature Showcase"
visitCount = 42
today      = Now()

' ── 2. STRING FUNCTIONS ────────────────────────────────────
Dim rawStr, upperStr, lowerStr, trimmedStr, replacedStr
rawStr      = "   Hello, Classic ASP World!   "
upperStr    = UCase(rawStr)
lowerStr    = LCase(rawStr)
trimmedStr  = Trim(rawStr)
replacedStr = Replace(trimmedStr, "Classic ASP", "VBScript")

' ── 3. DATE & TIME FUNCTIONS ───────────────────────────────
Dim yr, mo, dy, hr, mn, sc
yr = Year(today)   :  mo = Month(today)  :  dy = Day(today)
hr = Hour(today)   :  mn = Minute(today) :  sc = Second(today)
Dim formattedDate
formattedDate = dy & "/" & mo & "/" & yr & "  " & _
                Right("0"&hr,2) & ":" & Right("0"&mn,2) & ":" & Right("0"&sc,2)

' ── 4. ARRAYS ──────────────────────────────────────────────
Dim fruits(4)
fruits(0) = "Apple"
fruits(1) = "Banana"
fruits(2) = "Cherry"
fruits(3) = "Date"
fruits(4) = "Elderberry"

' Dynamic array
Dim vegs()
ReDim vegs(2)
vegs(0) = "Carrot"
vegs(1) = "Broccoli"
vegs(2) = "Spinach"
ReDim Preserve vegs(3)   ' extend without losing data
vegs(3) = "Kale"

' ── 5. FUNCTIONS & SUBS ────────────────────────────────────
Function Factorial(n)
    If n <= 1 Then
        Factorial = 1
    Else
        Factorial = n * Factorial(n - 1)
    End If
End Function

Function IsPrime(n)
    Dim i
    If n < 2 Then IsPrime = False : Exit Function
    For i = 2 To Int(Sqr(n))
        If n Mod i = 0 Then IsPrime = False : Exit Function
    Next
    IsPrime = True
End Function

Function PadLeft(s, totalLen, padChar)
    Do While Len(s) < totalLen
        s = padChar & s
    Loop
    PadLeft = s
End Function

Sub LogMessage(msg)
    ' In real code this might write to a log file
    ' Here we just store it for display
    logOutput = logOutput & "<li>" & Server.HTMLEncode(msg) & "</li>" & vbCrLf
End Sub

Dim logOutput
logOutput = ""
LogMessage "Page initialised at " & formattedDate
LogMessage "User agent: " & Request.ServerVariables("HTTP_USER_AGENT")
LogMessage "Remote IP:  " & Request.ServerVariables("REMOTE_ADDR")

' ── 6. LOOPS & CONDITIONALS ────────────────────────────────
Dim i, primeList, factList
primeList = ""
factList  = ""

For i = 2 To 20


    If IsPrime(i) Then
        primeList = primeList & i & " "
    End If
Next

For i = 1 To 8
    factList = factList & i & "! = " & Factorial(i) & "<br>"
Next

' ── 7. DO WHILE / DO UNTIL ─────────────────────────────────
Dim counter, fibResult
counter   = 0
fibResult = ""
Dim a, b, temp
a = 0 : b = 1
Do While a < 200
    fibResult = fibResult & a & " "
    temp = a + b
    a    = b
    b    = temp
    counter = counter + 1
Loop

' ── 8. SELECT CASE ─────────────────────────────────────────
Dim dayName
Select Case Weekday(today)
    Case 1: dayName = "Sunday"
    Case 2: dayName = "Monday"
    Case 3: dayName = "Tuesday"
    Case 4: dayName = "Wednesday"
    Case 5: dayName = "Thursday"
    Case 6: dayName = "Friday"
    Case 7: dayName = "Saturday"
End Select

' ── 9. FORM HANDLING (POST) ────────────────────────────────
Dim formName, formAge, formMsg
formName = ""
formAge  = 0
formMsg  = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    formName = Trim(Request.Form("txtName"))
	if Request.Form("txtAge")="" then 
		formAge=0
	else
		formAge  = CInt(Request.Form("txtAge"))
	end if

    If formName = "" Then
        formMsg = "⚠ Please enter your name."
    ElseIf formAge < 1 Or formAge > 120 Then
        formMsg = "⚠ Enter a valid age (1–120)."
    Else
        formMsg = "✅ Hello, <strong>" & Server.HTMLEncode(formName) & "</strong>! " & _
                  "You are <strong>" & formAge & "</strong> years old. " & _
                  "In 10 years you will be <strong>" & (formAge + 10) & "</strong>."
    End If
End If

' ── 10. COOKIES ────────────────────────────────────────────
Dim cookieVisits
cookieVisits = Request.Cookies("asp_visits")
If cookieVisits = "" Then
    cookieVisits = 1
Else
    cookieVisits = CInt(cookieVisits) + 1
End If
Response.Cookies("asp_visits")         = cookieVisits
Response.Cookies("asp_visits").Expires = DateAdd("d", 30, Now())

' ── 11. SESSION ────────────────────────────────────────────
If Session("first_hit") = "" Then
    Session("first_hit") = formattedDate
End If

' ── 12. QUERYSTRING ────────────────────────────────────────
Dim qsColor
qsColor = Request.QueryString("color")
If qsColor = "" Then qsColor = "steelblue"
' Basic sanitise – keep only alphanum/# so it's safe to drop into CSS
Dim safeColor, c
safeColor = ""
Dim asc_c
For i = 1 To Len(qsColor)
    c     = Mid(qsColor, i, 1)
    asc_c = Asc(UCase(c))
    If (asc_c >= 65 And asc_c <= 90) Or _
       (asc_c >= 48 And asc_c <= 57) Or c = "#" Then
        safeColor = safeColor & c
    End If
Next
If safeColor = "" Then safeColor = "steelblue"

' ── 13. ERROR HANDLING ─────────────────────────────────────
Dim errDemo
On Error Resume Next
errDemo = CInt("not-a-number")
Dim errNum, errDesc
errNum  = Err.Number
errDesc = Err.Description
On Error GoTo 0

' ── 14. FILE I/O via FSO ───────────────────────────────────
Dim fso, tmpFile, tmpPath, fileContent
Set fso     = Server.CreateObject("Scripting.FileSystemObject")
tmpPath     = Server.MapPath(".") & "\asp_demo_log.txt"
fileContent = "(FSO write/read not permitted in this sandbox)"

On Error Resume Next
Dim ts
Set ts = fso.OpenTextFile(tmpPath, 8, True)   ' 8 = ForAppending, True = create
If Err.Number = 0 Then
    ts.WriteLine "[" & formattedDate & "] Page visited. Visits so far: " & cookieVisits
    ts.Close
    Set ts = fso.OpenTextFile(tmpPath, 1)      ' 1 = ForReading
    fileContent = Server.HTMLEncode(ts.ReadAll())
    ts.Close
End If
On Error GoTo 0
Set fso = Nothing

' ── 15. DICTIONARY OBJECT ──────────────────────────────────
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")
dict.Add "language", "VBScript"
dict.Add "platform", "Classic ASP / IIS"
dict.Add "released", "1996"
dict.Add "fun_level", "Over 9000"

' ── 16. MATH & CONVERSION FUNCTIONS ───────────────────────
Dim mathResults
mathResults = "Sqr(144)=" & Sqr(144) & " | " & _
              "Abs(-77)=" & Abs(-77)  & " | " & _
              "Int(9.9)=" & Int(9.9)  & " | " & _
              "Rnd≈"      & FormatNumber(Rnd(),4)

' ── PAGE COMPLETE – NOW RENDER HTML ────────────────────────
Response.ContentType = "text/html"
Response.CharSet     = "utf-8"
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><%= pageTitle %></title>
<style>
  :root{--accent:<%=safeColor%>}
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;color:#222;line-height:1.6}
  header{background:var(--accent);color:#fff;padding:2rem;text-align:center}
  header h1{font-size:2rem;letter-spacing:.03em}
  header p{opacity:.85;margin-top:.4rem}
  main{max-width:1100px;margin:2rem auto;padding:0 1rem;display:grid;gap:1.5rem;
       grid-template-columns:repeat(auto-fill,minmax(320px,1fr))}
  section{background:#fff;border-radius:8px;padding:1.4rem;box-shadow:0 2px 8px rgba(0,0,0,.08)}
  section h2{font-size:1rem;text-transform:uppercase;letter-spacing:.08em;color:var(--accent);
             border-bottom:2px solid var(--accent);padding-bottom:.4rem;margin-bottom:.8rem}
  code,pre{font-family:Consolas,monospace;font-size:.85rem}
  pre{background:#f7f7f7;padding:.8rem;border-radius:4px;overflow-x:auto;white-space:pre-wrap}
  table{width:100%;border-collapse:collapse;font-size:.88rem}
  th{background:var(--accent);color:#fff;padding:.4rem .6rem;text-align:left}
  td{padding:.35rem .6rem;border-bottom:1px solid #eee}
  tr:last-child td{border-bottom:none}
  ul{padding-left:1.2rem;font-size:.88rem}
  li{margin:.2rem 0}
  .badge{display:inline-block;background:var(--accent);color:#fff;border-radius:12px;
         padding:.15rem .7rem;font-size:.78rem;margin:.15rem}
  form label{display:block;font-size:.88rem;margin-bottom:.25rem;font-weight:600}
  form input[type=text],form input[type=number]{width:100%;padding:.45rem .6rem;
    border:1px solid #ccc;border-radius:4px;font-size:.9rem;margin-bottom:.7rem}
  form button{background:var(--accent);color:#fff;border:none;padding:.5rem 1.4rem;
              border-radius:4px;cursor:pointer;font-size:.95rem}
  .msg{margin-top:.8rem;padding:.6rem;border-radius:4px;background:#e8f4fd;font-size:.9rem}
  .tip{font-size:.78rem;color:#888;margin-top:.5rem}
  footer{text-align:center;padding:1.5rem;color:#888;font-size:.82rem}
</style>
</head>
<body>

<header>
  <h1>⚙ Classic ASP Feature Showcase</h1>
  <p>All server-side — rendered by VBScript on IIS &nbsp;|&nbsp; Today: <strong><%=dayName%>, <%=formattedDate%></strong></p>
  <p class="tip">Tip: append <code>?color=crimson</code> (or any CSS colour) to change the accent.</p>
</header>

<main>

  <!-- 1. STRING FUNCTIONS -->
  <section>
    <h2>1 · String Functions</h2>
    <table>
      <tr><th>Function</th><th>Result</th></tr>
      <tr><td>UCase()</td><td><%=upperStr%></td></tr>
      <tr><td>LCase()</td><td><%=lowerStr%></td></tr>
      <tr><td>Trim()</td><td>"<%=trimmedStr%>"</td></tr>
      <tr><td>Replace()</td><td><%=replacedStr%></td></tr>
      <tr><td>Len(trimmed)</td><td><%=Len(trimmedStr)%></td></tr>
      <tr><td>Left(,5)</td><td><%=Left(trimmedStr,5)%></td></tr>
      <tr><td>Right(,6)</td><td><%=Right(trimmedStr,6)%></td></tr>
      <tr><td>Mid(,8,7)</td><td><%=Mid(trimmedStr,8,7)%></td></tr>
      <tr><td>InStr("World")</td><td>pos <%=InStr(trimmedStr,"World")%></td></tr>
      <tr><td>StrReverse()</td><td><%=StrReverse("VBScript")%></td></tr>
      <tr><td>String(5,"*")</td><td><%=String(5,"*")%></td></tr>
      <tr><td>Space(4) + "|"</td><td>[<%=Space(4)%>|]</td></tr>
    </table>
  </section>

  <!-- 2. DATE & TIME -->
  <section>
    <h2>2 · Date &amp; Time Functions</h2>
    <table>
      <tr><th>Function</th><th>Result</th></tr>
      <tr><td>Now()</td><td><%=Now()%></td></tr>
      <tr><td>Date()</td><td><%=Date()%></td></tr>
      <tr><td>Time()</td><td><%=Time()%></td></tr>
      <tr><td>Year / Month / Day</td><td><%=yr%> / <%=mo%> / <%=dy%></td></tr>
      <tr><td>Hour / Minute / Second</td><td><%=hr%> : <%=mn%> : <%=sc%></td></tr>
      <tr><td>Weekday()</td><td><%=Weekday(today)%> (<%=dayName%>)</td></tr>
      <tr><td>WeekdayName()</td><td><%=WeekdayName(Weekday(today))%></td></tr>
      <tr><td>MonthName()</td><td><%=MonthName(mo)%></td></tr>
      <tr><td>DateAdd("d",7)</td><td><%=DateAdd("d",7,Date())%></td></tr>
      <tr><td>DateDiff("d",#1/1/2000#)</td><td><%=DateDiff("d","1/1/2000",Date())%> days</td></tr>
      <tr><td>FormatDateTime</td><td><%=FormatDateTime(today,1)%></td></tr>
    </table>
  </section>

  <!-- 3. ARRAYS -->
  <section>
    <h2>3 · Arrays</h2>
    <p><strong>Static array (fruits):</strong></p>
    <p>
    <% Dim f : For f = 0 To UBound(fruits) %>
      <span class="badge"><%=fruits(f)%></span>
    <% Next %>
    </p>
    <br>
    <p><strong>Dynamic array after ReDim Preserve (vegs):</strong></p>
    <p>
    <% Dim v : For v = 0 To UBound(vegs) %>
      <span class="badge"><%=vegs(v)%></span>
    <% Next %>
    </p>
    <br>
    <p><strong>Split() &amp; Join():</strong></p>
    <%
      Dim csvLine, splitArr, joinedBack
      csvLine   = "red,green,blue,yellow"
      splitArr  = Split(csvLine, ",")
      joinedBack = Join(splitArr, " | ")
    %>
    <code>Split: [<%=splitArr(0)%>] [<%=splitArr(1)%>] [<%=splitArr(2)%>] [<%=splitArr(3)%>]</code><br>
    <code>Join:  <%=joinedBack%></code>
  </section>

  <!-- 4. FUNCTIONS & RECURSION -->
  <section>
    <h2>4 · Functions, Subs &amp; Recursion</h2>
    <p><strong>Factorial (recursive):</strong></p>
    <pre><%=factList%></pre>
    <p><strong>Custom PadLeft():</strong></p>
    <code><%=PadLeft("42",8,"0")%></code>
  </section>

  <!-- 5. LOOPS & PRIMES -->
  <section>
    <h2>5 · Loops &amp; Conditionals</h2>
    <p><strong>Primes 2–20 (For…Next + IsPrime):</strong></p>
    <%
      Dim p : For Each p In Split(Trim(primeList)," ")
        If p <> "" Then %><span class="badge"><%=p%></span><% End If
      Next
    %>
    <br><br>
    <p><strong>Fibonacci &lt;200 (Do While):</strong></p>
    <pre><%=Trim(fibResult)%></pre>
    <br>
    <p><strong>For Each over Dictionary keys:</strong></p>
    <table>
      <tr><th>Key</th><th>Value</th></tr>
      <% Dim k : For Each k In dict %>
      <tr><td><%=k%></td><td><%=dict(k)%></td></tr>
      <% Next %>
    </table>
  </section>

  <!-- 6. MATH & TYPE CONVERSION -->
  <section>
    <h2>6 · Math &amp; Type Conversion</h2>
    <pre><%=mathResults%></pre>
    <table>
      <tr><th>Expression</th><th>Result</th></tr>
      <tr><td>CStr(3.14)</td><td><%=CStr(3.14)%></td></tr>
      <tr><td>CInt("99")</td><td><%=CInt("99")%></td></tr>
      <tr><td>CDbl("1.23456")</td><td><%=CDbl("1.23456")%></td></tr>
      <tr><td>CBool(0)</td><td><%=CBool(0)%></td></tr>
      <tr><td>CBool(1)</td><td><%=CBool(1)%></td></tr>
      <tr><td>Hex(255)</td><td><%=Hex(255)%></td></tr>
      <tr><td>Oct(255)</td><td><%=Oct(255)%></td></tr>
      <tr><td>FormatNumber(Pi,4)</td><td><%=FormatNumber(4*Atn(1),4)%></td></tr>
      <tr><td>FormatCurrency(9.5)</td><td><%=FormatCurrency(9.5)%></td></tr>
      <tr><td>FormatPercent(0.765)</td><td><%=FormatPercent(0.765)%></td></tr>
      <tr><td>IsNumeric("abc")</td><td><%=IsNumeric("abc")%></td></tr>
      <tr><td>IsDate("Feb 22")</td><td><%=IsDate("Feb 22")%></td></tr>
    </table>
  </section>

  <!-- 7. ERR HANDLING -->
  <section>
    <h2>7 · Err Handling</h2>
    <p>Attempted <code>CInt("not-a-number")</code> inside <code>On Err Resume Next</code>:</p>
    <table>
      <tr><th>Property</th><th>Value</th></tr>
      <tr><td>Err.Number</td><td><%=errNum%></td></tr>
      <tr><td>Err.Description</td><td><%=errDesc%></td></tr>
    </table>
  </section>

  <!-- 8. REQUEST OBJECTS -->
  <section>
    <h2>8 · Request Object</h2>
    <table>
      <tr><th>Item</th><th>Value</th></tr>
      <tr><td>REQUEST_METHOD</td><td><%=Request.ServerVariables("REQUEST_METHOD")%></td></tr>
      <tr><td>SERVER_NAME</td><td><%=Request.ServerVariables("SERVER_NAME")%></td></tr>
      <tr><td>SERVER_PORT</td><td><%=Request.ServerVariables("SERVER_PORT")%></td></tr>
      <tr><td>SCRIPT_NAME</td><td><%=Request.ServerVariables("SCRIPT_NAME")%></td></tr>
      <tr><td>HTTP_ACCEPT_LANGUAGE</td><td><%=Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")%></td></tr>
      <tr><td>QueryString ?color</td><td><%=Server.HTMLEncode(Request.QueryString("color"))%></td></tr>
    </table>
    <p class="tip">User-Agent: <%=Server.HTMLEncode(Left(Request.ServerVariables("HTTP_USER_AGENT"),80))%>…</p>
  </section>

  <!-- 9. SESSION & COOKIES -->
  <section>
    <h2>9 · Session &amp; Cookies</h2>
    <table>
      <tr><th>Item</th><th>Value</th></tr>
      <tr><td>Session("first_hit")</td><td><%=Session("first_hit")%></td></tr>
      <tr><td>Session.SessionID</td><td><%=Session.SessionID%></td></tr>
      <tr><td>Session.Timeout</td><td><%=Session.Timeout%> min</td></tr>
      <tr><td>Cookie asp_visits</td><td><%=cookieVisits%> time(s)</td></tr>
    </table>
    <p class="tip">The cookie expires in 30 days. Reload to increment the counter.</p>
  </section>

  <!-- 10. RESPONSE OBJECT -->
  <section>
    <h2>10 · Response Object</h2>
    <table>
      <tr><th>Property / Method</th><th>Value / Action</th></tr>
      <tr><td>Response.ContentType</td><td>text/html</td></tr>
      <tr><td>Response.CharSet</td><td>utf-8</td></tr>
      <tr><td>Response.Buffer</td><td><%=Response.Buffer%></td></tr>
      <tr><td>Response.IsClientConnected</td><td><%=Response.IsClientConnected%></td></tr>
      <tr><td>Response.Write()</td><td><% Response.Write "<em>Written inline!</em>" %></td></tr>
    </table>
    <% Response.Flush() ' flush buffer mid-page %>
    <p class="tip"><code>Response.Flush()</code> was called after this table.</p>
  </section>

  <!-- 11. SERVER OBJECT -->
  <section>
    <h2>11 · Server Object</h2>
    <table>
      <tr><th>Method</th><th>Result</th></tr>
      <tr><td>Server.HTMLEncode</td><td><%=Server.HTMLEncode("<b>Bold & <i>Italic</i></b>")%></td></tr>
      <tr><td>Server.URLEncode</td><td><%=Server.URLEncode("hello world & more")%></td></tr>
      <tr><td>Server.MapPath(".")</td><td><%=Server.MapPath(".")%></td></tr>
      <tr><td>Server.ScriptTimeout</td><td><%=Server.ScriptTimeout%> sec</td></tr>
    </table>
  </section>

  <!-- 12. SCRIPTING.DICTIONARY -->
  <section>
    <h2>12 · Scripting.Dictionary</h2>
    <%
      Dim dict2
      Set dict2 = Server.CreateObject("Scripting.Dictionary")
      dict2.CompareMode = 1   ' vbTextCompare – case-insensitive
      dict2("Alpha")   = 100
      dict2("Beta")    = 200
      dict2("Gamma")   = 300
      dict2("Delta")   = 400
    %>
    <table>
      <tr><th>Key</th><th>Value</th><th>Exists?</th></tr>
      <% Dim dk : For Each dk In dict2 %>
      <tr><td><%=dk%></td><td><%=dict2(dk)%></td><td><%=dict2.Exists(dk)%></td></tr>
      <% Next %>
      <tr><td>Epsilon</td><td>—</td><td><%=dict2.Exists("Epsilon")%></td></tr>
    </table>
    <p class="tip">Count: <%=dict2.Count%> | Keys: <%=Join(dict2.Keys,", ")%></p>
    <% Set dict2 = Nothing %>
  </section>

  <!-- 13. FILE SYSTEM OBJECT -->
  <section>
    <h2>13 · FileSystemObject (FSO)</h2>
    <% If fileContent = "(FSO write/read not permitted in this sandbox)" Then %>
      <p><%=fileContent%></p>
      <p class="tip">On a real IIS server with write permission, this section shows the live log file written and read back each visit.</p>
    <% Else %>
      <p><strong>Log file contents (asp_demo_log.txt):</strong></p>
      <pre><%=fileContent%></pre>
    <% End If %>
  </section>

  <!-- 14. SERVER LOG SUB OUTPUT -->
  <section>
    <h2>14 · LogMessage Sub Output</h2>
    <ul>
      <%=logOutput%>
    </ul>
  </section>

  <!-- 15. FORM PROCESSING -->
  <section style="grid-column:1/-1">
    <h2>15 · Form Processing (POST)</h2>
    <%
      Dim formAction
      If Request.QueryString("color") <> "" Then
          formAction = "default.asp?color=" & Server.URLEncode(Request.QueryString("color"))
      Else
          formAction = "default.asp"
      End If
    %>
    <form method="post" action="<%=formAction%>">
      <label for="txtName">Your Name</label>
      <input type="text" id="txtName" name="txtName"
             value="<%=Server.HTMLEncode(formName)%>" placeholder="e.g. Ada Lovelace">
      <label for="txtAge">Your Age</label>
      <input type="number" id="txtAge" name="txtAge"
             value="<% If formAge > 0 Then Response.Write(formAge) End If %>" placeholder="e.g. 30" min="1" max="120">
      <button type="submit">Submit via POST ↵</button>
    </form>
    <% If formMsg <> "" Then %>
    <div class="msg"><%=formMsg%></div>
    <% End If %>
    <p class="tip">VBScript validates server-side: empty name → err, bad age → ERR, otherwise greeting + age + 10.</p>
  </section>

</main>

<footer>
  Generated entirely server-side by VBScript / Classic ASP &nbsp;·&nbsp;
  Page rendered at <%=formattedDate%> &nbsp;·&nbsp;
  Visit #<%=cookieVisits%> (cookie)
</footer>

</body>
</html>
