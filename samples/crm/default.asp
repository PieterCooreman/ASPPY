<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file="includes/includes.asp" -->
<%
Dim db, contactsSvc, groupsSvc, m, action, flashMsg
Dim filterGroup, filterQ
Set db = New cls_db
db.Open
Set contactsSvc = New cls_contact
Set groupsSvc = New cls_group

m = LCase(Trim("" & Request.QueryString("m")))
If m = "" Then m = "contacts"
action = LCase(Trim("" & Request.Form("action")))

If IsPost() Then
    On Error Resume Next
    If m = "contacts" Then
        If action = "save" Then
            contactsSvc.SaveContact db, Request.Form("id"), Request.Form("first_name"), Request.Form("last_name"), Request.Form("email"), Request.Form("phone"), Request.Form("company"), Request.Form("notes"), Request.Form("group_id")
            If Err.Number <> 0 Then
                SetFlash "Could not save contact."
                Err.Clear
            Else
                SetFlash "Contact saved."
            End If
        ElseIf action = "delete" Then
            contactsSvc.DeleteContact db, Request.Form("id")
            SetFlash "Contact deleted."
        End If
        db.Close: Set db = Nothing
        Response.Redirect "default.asp?m=contacts"
        Response.End
    End If

    If m = "groups" Then
        If action = "create" Then
            groupsSvc.CreateGroup db, Request.Form("name")
            SetFlash "Group created."
        ElseIf action = "update" Then
            groupsSvc.UpdateGroup db, Request.Form("id"), Request.Form("name")
            SetFlash "Group updated."
        ElseIf action = "delete" Then
            groupsSvc.DeleteGroup db, Request.Form("id")
            SetFlash "Group deleted."
        ElseIf action = "up" Then
            groupsSvc.MoveUp db, Request.Form("id")
            SetFlash "Group order updated."
        ElseIf action = "down" Then
            groupsSvc.MoveDown db, Request.Form("id")
            SetFlash "Group order updated."
        End If
        db.Close: Set db = Nothing
        Response.Redirect "default.asp?m=groups"
        Response.End
    End If
    On Error GoTo 0
End If

If m = "export" Then
    Dim csvData
    csvData = contactsSvc.ExportCsv(db)
    Response.ContentType = "text/csv"
    Response.AddHeader "Content-Disposition", "attachment; filename=contacts.csv"
    Response.Write csvData
    db.Close: Set db = Nothing
    Response.End
End If

flashMsg = GetFlash()
filterGroup = ToInt(Request.QueryString("group"), 0)
filterQ = Trim("" & Request.QueryString("q"))
%>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>CRM Sample</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css" rel="stylesheet" crossorigin="anonymous">
  <style>
  body{background:#f5f7fb}
  .app-shell{max-width:1200px;margin:0 auto}
  </style>
</head>
<body>
<main class="app-shell py-4">
  <div class="d-flex justify-content-between align-items-center mb-3">
    <h1 class="h4 mb-0">CRM Contacts</h1>
    <div class="d-flex gap-2">
      <a class="btn btn-outline-secondary btn-sm" href="default.asp?m=groups">Manage Groups</a>
      <a class="btn btn-primary btn-sm" href="default.asp?m=export">Export CSV</a>
    </div>
  </div>

  <% If flashMsg <> "" Then %><div class="alert alert-info"><%=H(flashMsg)%></div><% End If %>

  <% If m = "groups" Then %>
    <div class="card border-0 shadow-sm mb-3"><div class="card-body">
      <h2 class="h6">Add Group</h2>
      <form class="row g-2" method="post" action="default.asp?m=groups">
        <input type="hidden" name="action" value="create">
        <div class="col-md-6"><input class="form-control" name="name" placeholder="Group name" required></div>
        <div class="col-auto"><button class="btn btn-primary" type="submit">Add</button></div>
      </form>
    </div></div>

    <div class="card border-0 shadow-sm"><div class="table-responsive"><table class="table mb-0 align-middle">
      <thead><tr><th>Order</th><th>Name</th><th>Actions</th></tr></thead><tbody>
      <%
      Dim rsGroups
      Set rsGroups = groupsSvc.ListAll(db)
      Do Until rsGroups.EOF
      %>
      <tr>
        <td><%=H(Nz(rsGroups("sort_order"), ""))%></td>
        <td>
          <form class="d-flex gap-2" method="post" action="default.asp?m=groups">
            <input type="hidden" name="action" value="update"><input type="hidden" name="id" value="<%=rsGroups("id")%>">
            <input class="form-control form-control-sm" name="name" value="<%=H(rsGroups("name"))%>">
            <button class="btn btn-sm btn-outline-primary" type="submit">Save</button>
          </form>
        </td>
        <td class="d-flex gap-2">
          <form method="post" action="default.asp?m=groups"><input type="hidden" name="action" value="up"><input type="hidden" name="id" value="<%=rsGroups("id")%>"><button class="btn btn-sm btn-outline-secondary" type="submit">Up</button></form>
          <form method="post" action="default.asp?m=groups"><input type="hidden" name="action" value="down"><input type="hidden" name="id" value="<%=rsGroups("id")%>"><button class="btn btn-sm btn-outline-secondary" type="submit">Down</button></form>
          <form method="post" action="default.asp?m=groups"><input type="hidden" name="action" value="delete"><input type="hidden" name="id" value="<%=rsGroups("id")%>"><button class="btn btn-sm btn-outline-danger" type="submit" onclick="return confirm('Delete this group?')">Delete</button></form>
        </td>
      </tr>
      <% rsGroups.MoveNext: Loop
      rsGroups.Close: Set rsGroups = Nothing %>
      </tbody></table></div></div>
    <div class="mt-3"><a href="default.asp" class="btn btn-sm btn-outline-dark">Back to contacts</a></div>

  <% Else %>
    <div class="card border-0 shadow-sm mb-3"><div class="card-body">
      <h2 class="h6">Filters</h2>
      <form class="row g-2" method="get" action="default.asp">
        <input type="hidden" name="m" value="contacts">
        <div class="col-md-4"><input class="form-control" name="q" value="<%=H(filterQ)%>" placeholder="Search name/email/company"></div>
        <div class="col-md-3">
          <select class="form-select" name="group">
            <option value="0">All groups</option>
            <%
            Set rsGroups = groupsSvc.ListAll(db)
            Do Until rsGroups.EOF
            %>
            <option value="<%=rsGroups("id")%>" <%If filterGroup=ToInt(rsGroups("id"),0) Then Response.Write("selected")%>><%=H(rsGroups("name"))%></option>
            <% rsGroups.MoveNext: Loop
            rsGroups.Close: Set rsGroups = Nothing %>
          </select>
        </div>
        <div class="col-auto"><button class="btn btn-outline-primary" type="submit">Apply</button></div>
      </form>
    </div></div>

    <%
    Dim editId, fFirst, fLast, fEmail, fPhone, fCompany, fNotes, fGroupId, rsEdit
    editId = ToInt(Request.QueryString("id"), 0)
    fFirst = "": fLast = "": fEmail = "": fPhone = "": fCompany = "": fNotes = "": fGroupId = 0
    If editId > 0 Then
      Set rsEdit = contactsSvc.GetById(db, editId)
      If Not rsEdit.EOF Then
        fFirst = "" & Nz(rsEdit("first_name"), "")
        fLast = "" & Nz(rsEdit("last_name"), "")
        fEmail = "" & Nz(rsEdit("email"), "")
        fPhone = "" & Nz(rsEdit("phone"), "")
        fCompany = "" & Nz(rsEdit("company"), "")
        fNotes = "" & Nz(rsEdit("notes"), "")
        fGroupId = ToInt(Nz(rsEdit("group_id"), 0), 0)
      End If
      rsEdit.Close: Set rsEdit = Nothing
    End If
    %>
    <div class="card border-0 shadow-sm mb-3"><div class="card-body">
      <h2 class="h6"><%If editId>0 Then Response.Write("Edit Contact") Else Response.Write("Add Contact") End If%></h2>
      <form method="post" action="default.asp?m=contacts" class="row g-2">
        <input type="hidden" name="action" value="save"><input type="hidden" name="id" value="<%=editId%>">
        <div class="col-md-3"><input class="form-control" name="first_name" placeholder="First name" value="<%=H(fFirst)%>" required></div>
        <div class="col-md-3"><input class="form-control" name="last_name" placeholder="Last name" value="<%=H(fLast)%>" required></div>
        <div class="col-md-3"><input class="form-control" name="email" placeholder="Email" type="email" value="<%=H(fEmail)%>"></div>
        <div class="col-md-3"><input class="form-control" name="phone" placeholder="Phone" value="<%=H(fPhone)%>"></div>
        <div class="col-md-4"><input class="form-control" name="company" placeholder="Company" value="<%=H(fCompany)%>"></div>
        <div class="col-md-4">
          <select class="form-select" name="group_id">
            <option value="0">No group</option>
            <%
            Set rsGroups = groupsSvc.ListAll(db)
            Do Until rsGroups.EOF
            %>
            <option value="<%=rsGroups("id")%>" <%If fGroupId=ToInt(rsGroups("id"),0) Then Response.Write("selected")%>><%=H(rsGroups("name"))%></option>
            <% rsGroups.MoveNext: Loop
            rsGroups.Close: Set rsGroups = Nothing %>
          </select>
        </div>
        <div class="col-md-4 d-grid"><button class="btn btn-primary" type="submit">Save Contact</button></div>
        <div class="col-12"><textarea class="form-control" rows="3" name="notes" placeholder="Notes"><%=H(fNotes)%></textarea></div>
      </form>
    </div></div>

    <div class="card border-0 shadow-sm"><div class="table-responsive"><table class="table align-middle mb-0">
      <thead><tr><th>Name</th><th>Email</th><th>Phone</th><th>Company</th><th>Group</th><th>Actions</th></tr></thead><tbody>
      <%
      Dim rsContacts
      Set rsContacts = contactsSvc.ListAll(db, filterGroup, filterQ)
      Do Until rsContacts.EOF
      %>
      <tr>
        <td><%=H(rsContacts("first_name") & " " & rsContacts("last_name"))%></td>
        <td><%=H(rsContacts("email"))%></td>
        <td><%=H(rsContacts("phone"))%></td>
        <td><%=H(rsContacts("company"))%></td>
        <td><%=H(Nz(rsContacts("group_name"), ""))%></td>
        <td class="d-flex gap-2">
          <a class="btn btn-sm btn-outline-primary" href="default.asp?m=contacts&id=<%=rsContacts("id")%>">Edit</a>
          <form method="post" action="default.asp?m=contacts"><input type="hidden" name="action" value="delete"><input type="hidden" name="id" value="<%=rsContacts("id")%>"><button class="btn btn-sm btn-outline-danger" type="submit" onclick="return confirm('Delete this contact?')">Delete</button></form>
        </td>
      </tr>
      <% rsContacts.MoveNext: Loop
      rsContacts.Close: Set rsContacts = Nothing %>
      </tbody></table></div></div>
  <% End If %>
</main>
</body>
</html>
<%
db.Close
Set db = Nothing
%>
