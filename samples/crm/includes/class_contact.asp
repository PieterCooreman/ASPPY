<%
Class cls_contact
    Public Function ListAll(db, groupId, q)
        Dim sql, whereSql
        whereSql = " WHERE 1=1 "
        If ToInt(groupId, 0) > 0 Then
            whereSql = whereSql & " AND c.group_id=" & ToInt(groupId, 0)
        End If
        If Trim("" & q) <> "" Then
            whereSql = whereSql & " AND (c.first_name LIKE '%" & Q(q) & "%' OR c.last_name LIKE '%" & Q(q) & "%' OR c.email LIKE '%" & Q(q) & "%' OR c.company LIKE '%" & Q(q) & "%')"
        End If
        sql = "SELECT c.id,c.first_name,c.last_name,c.email,c.phone,c.company,c.notes,c.group_id,c.created_at,g.name AS group_name FROM contacts c LEFT JOIN groups g ON g.id=c.group_id" & whereSql & " ORDER BY c.last_name ASC, c.first_name ASC, c.id DESC"
        Set ListAll = db.Query(sql)
    End Function

    Public Function GetById(db, id)
        Set GetById = db.Query("SELECT * FROM contacts WHERE id=" & ToInt(id, 0) & " LIMIT 1")
    End Function

    Public Sub SaveContact(db, id, firstName, lastName, email, phone, company, notes, groupId)
        Dim sql, gid
        gid = ToInt(groupId, 0)
        If gid <= 0 Then
            gid = "NULL"
        End If
        If ToInt(id, 0) <= 0 Then
            sql = "INSERT INTO contacts(first_name,last_name,email,phone,company,notes,group_id,created_at) VALUES('" & Q(firstName) & "','" & Q(lastName) & "','" & Q(email) & "','" & Q(phone) & "','" & Q(company) & "','" & Q(notes) & "'," & gid & ",datetime('now'))"
        Else
            sql = "UPDATE contacts SET first_name='" & Q(firstName) & "',last_name='" & Q(lastName) & "',email='" & Q(email) & "',phone='" & Q(phone) & "',company='" & Q(company) & "',notes='" & Q(notes) & "',group_id=" & gid & ",updated_at=datetime('now') WHERE id=" & ToInt(id, 0)
        End If
        db.Execute sql
    End Sub

    Public Sub DeleteContact(db, id)
        db.Execute "DELETE FROM contacts WHERE id=" & ToInt(id, 0)
    End Sub

    Public Function ExportCsv(db)
        Dim rs, csv
        csv = "id,first_name,last_name,email,phone,company,group_name,created_at" & vbCrLf
        Set rs = db.Query("SELECT c.id,c.first_name,c.last_name,c.email,c.phone,c.company,COALESCE(g.name,'') AS group_name,c.created_at FROM contacts c LEFT JOIN groups g ON g.id=c.group_id ORDER BY c.last_name,c.first_name,c.id")
        Do Until rs.EOF
            csv = csv & CsvCell(rs("id")) & "," & CsvCell(rs("first_name")) & "," & CsvCell(rs("last_name")) & "," & CsvCell(rs("email")) & "," & CsvCell(rs("phone")) & "," & CsvCell(rs("company")) & "," & CsvCell(rs("group_name")) & "," & CsvCell(rs("created_at")) & vbCrLf
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        ExportCsv = csv
    End Function

    Private Function CsvCell(v)
        Dim s
        s = Replace("" & Nz(v, ""), Chr(34), Chr(34) & Chr(34))
        CsvCell = Chr(34) & s & Chr(34)
    End Function
End Class
%>
