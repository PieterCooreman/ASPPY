<%
Class cls_blueprint
    Public Function GetAllBlueprints(db, sortBy, search, category, status)
        Dim sql, rs
        sql = "SELECT * FROM blueprints WHERE 1=1"
        
        If search <> "" Then
            sql = sql & " AND (name LIKE '%" & Q(search) & "%' OR one_liner LIKE '%" & Q(search) & "%' OR problem LIKE '%" & Q(search) & "%' OR tags LIKE '%" & Q(search) & "%')"
        End If
        
        If category <> "" And category <> "all" Then
            sql = sql & " AND category = '" & Q(category) & "'"
        End If
        
        If status <> "" And status <> "all" Then
            sql = sql & " AND status = '" & Q(status) & "'"
        End If
        
        Select Case sortBy
            Case "spark"
                sql = sql & " ORDER BY spark_score DESC, created_at DESC"
            Case "name"
                sql = sql & " ORDER BY name ASC"
            Case Else
                sql = sql & " ORDER BY created_at DESC"
        End Select
        
        Set rs = db.Query(sql)
        Set GetAllBlueprints = rs
    End Function
    
    Public Function GetBlueprintById(db, id)
        Dim sql, rs
        sql = "SELECT * FROM blueprints WHERE id = '" & Q(id) & "'"
        Set rs = db.Query(sql)
        Set GetBlueprintById = rs
    End Function
    
    Public Function CreateBlueprint(db, name, oneLiner, problem, userTarget, hook, risk, sparkScore, category, tags, status, notes)
        Dim id, createdAt, updatedAt, sql
        
        id = GenerateUUID()
        createdAt = GetTimestamp()
        updatedAt = createdAt
        
        sql = "INSERT INTO blueprints (id, name, one_liner, problem, user_target, hook, risk, spark_score, category, tags, status, notes, created_at, updated_at) VALUES ('" & _
            Q(id) & "', '" & Q(name) & "', '" & Q(oneLiner) & "', '" & Q(problem) & "', '" & Q(userTarget) & "', '" & _
            Q(hook) & "', '" & Q(risk) & "', " & sparkScore & ", '" & Q(category) & "', '" & Q(tags) & "', '" & _
            Q(status) & "', '" & Q(notes) & "', '" & Q(createdAt) & "', '" & Q(updatedAt) & "')"
        
        db.Execute sql
        
        db.Execute "INSERT INTO spark_history (id, blueprint_id, score, recorded_at) VALUES ('" & GenerateUUID() & "', '" & id & "', " & sparkScore & ", '" & Q(createdAt) & "')"
        
        CreateBlueprint = id
    End Function
    
    Public Function UpdateBlueprint(db, id, name, oneLiner, problem, userTarget, hook, risk, sparkScore, category, tags, status, notes)
        Dim updatedAt, sql
        updatedAt = GetTimestamp()
        
        sql = "UPDATE blueprints SET name = '" & Q(name) & "', one_liner = '" & Q(oneLiner) & "', problem = '" & Q(problem) & "', " & _
            "user_target = '" & Q(userTarget) & "', hook = '" & Q(hook) & "', risk = '" & Q(risk) & "', " & _
            "spark_score = " & sparkScore & ", category = '" & Q(category) & "', tags = '" & Q(tags) & "', " & _
            "status = '" & Q(status) & "', notes = '" & Q(notes) & "', updated_at = '" & Q(updatedAt) & "' " & _
            "WHERE id = '" & Q(id) & "'"
        
        db.Execute sql
    End Function
    
    Public Function DeleteBlueprint(db, id)
        db.Execute "DELETE FROM spark_history WHERE blueprint_id = '" & Q(id) & "'"
        db.Execute "DELETE FROM blueprints WHERE id = '" & Q(id) & "'"
    End Function
    
    Public Function AddSparkRevisit(db, blueprintId, score)
        Dim id, recordedAt, sql
        id = GenerateUUID()
        recordedAt = GetTimestamp()
        
        db.Execute "UPDATE blueprints SET spark_score = " & score & ", updated_at = '" & Q(recordedAt) & "' WHERE id = '" & Q(blueprintId) & "'"
        
        sql = "INSERT INTO spark_history (id, blueprint_id, score, recorded_at) VALUES ('" & _
            Q(id) & "', '" & Q(blueprintId) & "', " & score & ", '" & Q(recordedAt) & "')"
        db.Execute sql
    End Function
    
    Public Function GetSparkHistory(db, blueprintId)
        Dim sql, rs
        sql = "SELECT * FROM spark_history WHERE blueprint_id = '" & Q(blueprintId) & "' ORDER BY recorded_at DESC"
        Set rs = db.Query(sql)
        Set GetSparkHistory = rs
    End Function
    
    Public Function GetStats(db)
        Dim stats(6)
        
        stats(0) = db.Scalar("SELECT COUNT(*) FROM blueprints", 0)
        stats(1) = db.Scalar("SELECT AVG(spark_score) FROM blueprints", 0)
        
        Dim rs, tagCounts, mostUsedTag, maxCount
        mostUsedTag = ""
        maxCount = 0
        
        Set rs = db.Query("SELECT tags FROM blueprints WHERE tags <> ''")
        Do While Not rs.EOF
            Dim tagsArr, t
            tagsArr = Split("" & rs("tags"), ",")
            For t = 0 To UBound(tagsArr)
                If Trim(tagsArr(t)) <> "" Then
                    Dim currentCount
                    currentCount = db.Scalar("SELECT COUNT(*) FROM blueprints WHERE tags LIKE '%" & Q(Trim(tagsArr(t))) & "%'", 0)
                    If currentCount > maxCount Then
                        maxCount = currentCount
                        mostUsedTag = Trim(tagsArr(t))
                    End If
                End If
            Next
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        stats(2) = mostUsedTag
        stats(3) = db.Scalar("SELECT COUNT(*) FROM blueprints WHERE status = 'idea'", 0)
        stats(4) = db.Scalar("SELECT COUNT(*) FROM blueprints WHERE status = 'shelved'", 0)
        stats(5) = db.Scalar("SELECT COUNT(*) FROM blueprints WHERE status = 'shipped'", 0)
        
        GetStats = stats
    End Function
    
    Public Function GetColdIdeas(db)
        Dim sql, rs
        sql = "SELECT * FROM blueprints WHERE updated_at < datetime('now', '-30 days') AND status IN ('idea', 'exploring') ORDER BY updated_at ASC"
        Set rs = db.Query(sql)
        Set GetColdIdeas = rs
    End Function
    
    Public Function GetTimelineData(db)
        Dim sql, rs
        sql = "SELECT strftime('%Y-%m', created_at) as month, COUNT(*) as count, AVG(spark_score) as avg_spark FROM blueprints GROUP BY month ORDER BY month DESC LIMIT 12"
        Set rs = db.Query(sql)
        Set GetTimelineData = rs
    End Function
End Class
%>
