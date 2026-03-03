<%
Function H(v)
    H = Server.HTMLEncode("" & v)
End Function

Function Q(v)
    Q = Replace("" & v, "'", "''")
End Function

Function Nz(v, fallback)
    If IsNull(v) Or IsEmpty(v) Then
        Nz = fallback
    Else
        Nz = v
    End If
End Function

Function ToInt(v, fallback)
    Dim s
    s = Trim("" & v)
    If s = "" Then
        ToInt = fallback
    ElseIf IsNumeric(s) Then
        ToInt = CLng(s)
    Else
        ToInt = fallback
    End If
End Function

Function IsPost()
    IsPost = (UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST")
End Function

Function GetTimestamp()
    GetTimestamp = Now()
End Function

Function GenerateUUID()
    Randomize
    GenerateUUID = "" & _
        Hex(Int(Rnd * 65535)) & "-" & _
        Hex(Int(Rnd * 65535)) & "-" & _
        Hex(Int(Rnd * 65535)) & "-" & _
        Hex(Int(Rnd * 65535)) & "-" & _
        Hex(Int(Rnd * 65535) * 65535)
End Function

Function GetCategories()
    Dim cats(7)(1)
    cats(0)(0) = "tool": cats(0)(1) = "🛠️ Tool"
    cats(1)(0) = "game": cats(1)(1) = "🎮 Game"
    cats(2)(0) = "app": cats(2)(1) = "📱 App"
    cats(3)(0) = "web": cats(3)(1) = "🌐 Web"
    cats(4)(0) = "ai": cats(4)(1) = "🤖 AI"
    cats(5)(0) = "business": cats(5)(1) = "🏪 Business"
    cats(6)(0) = "creative": cats(6)(1) = "🎨 Creative"
    cats(7)(0) = "social": cats(7)(1) = "🌍 Social"
    GetCategories = cats
End Function

Function GetCategoryLabel(category)
    Select Case LCase(category)
        Case "tool": GetCategoryLabel = "🛠️ Tool"
        Case "game": GetCategoryLabel = "🎮 Game"
        Case "app": GetCategoryLabel = "📱 App"
        Case "web": GetCategoryLabel = "🌐 Web"
        Case "ai": GetCategoryLabel = "🤖 AI"
        Case "business": GetCategoryLabel = "🏪 Business"
        Case "creative": GetCategoryLabel = "🎨 Creative"
        Case "social": GetCategoryLabel = "🌍 Social"
        Case Else: GetCategoryLabel = "📦 Other"
    End Select
End Function

Function GetStatuses()
    Dim stats(4)(1)
    stats(0)(0) = "idea": stats(0)(1) = "💡 Idea"
    stats(1)(0) = "exploring": stats(1)(1) = "🔨 Exploring"
    stats(2)(0) = "building": stats(2)(1) = "🚀 Building"
    stats(3)(0) = "shipped": stats(3)(1) = "📦 Shipped"
    stats(4)(0) = "shelved": stats(4)(1) = "🗄️ Shelved"
    GetStatuses = stats
End Function

Function GetStatusLabel(status)
    Select Case LCase(status)
        Case "idea": GetStatusLabel = "💡 Idea"
        Case "exploring": GetStatusLabel = "🔨 Exploring"
        Case "building": GetStatusLabel = "🚀 Building"
        Case "shipped": GetStatusLabel = "📦 Shipped"
        Case "shelved": GetStatusLabel = "🗄️ Shelved"
        Case Else: GetStatusLabel = "💡 Idea"
    End Select
End Function

Function GetSparkStars(score)
    Dim stars, i
    stars = ""
    For i = 1 To 5
        If i <= score Then
            stars = stars & "★"
        Else
            stars = stars & "☆"
        End If
    Next
    GetSparkStars = stars
End Function

Function GetSparkColor(score)
    Select Case score
        Case 5: GetSparkColor = "#ff6b6b"
        Case 4: GetSparkColor = "#ffa502"
        Case 3: GetSparkColor = "#ffd43b"
        Case 2: GetSparkColor = "#69db7c"
        Case 1: GetSparkColor = "#74c0fc"
        Case Else: GetSparkColor = "#adb5bd"
    End Select
End Function

Function FormatDate(dateStr)
    On Error Resume Next
    If dateStr <> "" Then
        FormatDate = FormatDateTime(CDate(dateStr), 2)
    Else
        FormatDate = ""
    End If
    On Error GoTo 0
End Function

Function GetTagList(tagsStr)
    GetTagList = Split(tagsStr, ",")
End Function

Function JoinTags(tagsArr)
    Dim i, result
    result = ""
    For i = 0 To UBound(tagsArr)
        If Trim(tagsArr(i)) <> "" Then
            If result <> "" Then result = result & ","
            result = result & Trim(tagsArr(i))
        End If
    Next
    JoinTags = result
End Function
%>
