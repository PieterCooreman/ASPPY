<%
Class cls_group
    Public Function ListAll(db)
        Set ListAll = db.Query("SELECT id,name,sort_order FROM groups ORDER BY sort_order ASC, id ASC")
    End Function

    Public Sub CreateGroup(db, name)
        Dim nextOrder
        nextOrder = ToInt(db.Scalar("SELECT COALESCE(MAX(sort_order),0)+1 FROM groups", 1), 1)
        db.Execute "INSERT INTO groups(name,sort_order) VALUES('" & Q(name) & "'," & nextOrder & ")"
    End Sub

    Public Sub UpdateGroup(db, id, name)
        db.Execute "UPDATE groups SET name='" & Q(name) & "' WHERE id=" & ToInt(id, 0)
    End Sub

    Public Sub DeleteGroup(db, id)
        Dim gid
        gid = ToInt(id, 0)
        db.Execute "UPDATE contacts SET group_id=NULL WHERE group_id=" & gid
        db.Execute "DELETE FROM groups WHERE id=" & gid
        NormalizeOrder db
    End Sub

    Public Sub MoveUp(db, id)
        Dim curOrder, prevId
        NormalizeOrder db
        curOrder = ToInt(db.Scalar("SELECT sort_order FROM groups WHERE id=" & ToInt(id, 0), 0), 0)
        prevId = ToInt(db.Scalar("SELECT id FROM groups WHERE sort_order=" & (curOrder - 1), 0), 0)
        If prevId > 0 Then
            db.Execute "UPDATE groups SET sort_order=" & (curOrder - 1) & " WHERE id=" & ToInt(id, 0)
            db.Execute "UPDATE groups SET sort_order=" & curOrder & " WHERE id=" & prevId
        End If
    End Sub

    Public Sub MoveDown(db, id)
        Dim curOrder, nextId
        NormalizeOrder db
        curOrder = ToInt(db.Scalar("SELECT sort_order FROM groups WHERE id=" & ToInt(id, 0), 0), 0)
        nextId = ToInt(db.Scalar("SELECT id FROM groups WHERE sort_order=" & (curOrder + 1), 0), 0)
        If nextId > 0 Then
            db.Execute "UPDATE groups SET sort_order=" & (curOrder + 1) & " WHERE id=" & ToInt(id, 0)
            db.Execute "UPDATE groups SET sort_order=" & curOrder & " WHERE id=" & nextId
        End If
    End Sub

    Public Sub NormalizeOrder(db)
        Dim rs, i
        i = 1
        Set rs = db.Query("SELECT id FROM groups ORDER BY sort_order ASC, id ASC")
        Do Until rs.EOF
            db.Execute "UPDATE groups SET sort_order=" & i & " WHERE id=" & ToInt(rs("id"), 0)
            i = i + 1
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End Sub
End Class
%>
