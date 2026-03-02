<%@ Language="VBScript" %>
<%
' ============================================================================
' Classic ASP / VBScript Quirks & Pitfalls Demo
' Save as quirks.asp and request in a browser via IIS
' ============================================================================

'--- Simple SafeCStr helper to avoid Invalid use of Null ----------------------
Function SafeCStr(v)
    If IsNull(v) Or IsEmpty(v) Then ' Handle Null and Empty explicitly. [web:11][web:22][web:25][web:27]
        SafeCStr = ""
    Else
        SafeCStr = CStr(v)
    End If
End Function

'--- CodePage / Charset mismatch (form vs page) --------------------------------
Response.CodePage = 65001
Response.CharSet  = "UTF-8"

Response.Write "<h1>Classic ASP / VBScript Quirks Demo</h1>"

'--- Variant, Empty, Null, and type coercion -----------------------------------
Dim vEmpty, vNull, vZero, vStrZero
vEmpty   = Empty       ' Variant/Empty. [web:25]
vNull    = Null        ' Variant/Null. [web:22][web:27]
vZero    = 0
vStrZero = "0"

Response.Write "<h2>Empty / Null / Type Coercion</h2>"

Response.Write "Empty + 5 = " & (vEmpty + 5) & "<br>"
Response.Write "Null + 5 = " & (vNull + 5) & " (looks empty!)<br>"

If vEmpty = "" Then
    Response.Write "vEmpty = """" is True (Empty coerces to empty string).<br>"
Else
    Response.Write "vEmpty = """" is False.<br>"
End If

If vNull = "" Then
    Response.Write "vNull = """" is False (but beware Null in expressions).<br>"
Else
    Response.Write "vNull = """" is False (explicitly showing typical expectation).<br>"
End If

If IsEmpty(vEmpty) Then
    Response.Write "IsEmpty(vEmpty) is True.<br>"
End If
If IsNull(vNull) Then
    Response.Write "IsNull(vNull) is True.<br>"
End If

'--- Non‑short‑circuit boolean evaluation -------------------------------------
Response.Write "<h2>No Short‑Circuit &amp; Null in Conditions</h2>"

Dim obj
Set obj = Nothing

On Error Resume Next
If (obj Is Nothing) And (obj.SomeMethod() = 1) Then
    Response.Write "You will never see this.<br>"
End If

If Err.Number <> 0 Then
    Response.Write "Err in compound If: " & Err.Number & " - " & Err.Description & "<br>"
    Err.Clear
End If
On Error GoTo 0

If (obj Is Nothing) Then
    Response.Write "obj Is Nothing, so we DON'T call methods on it.<br>"
Else
    If obj.SomeMethod() = 1 Then
        Response.Write "Safe nested condition.<br>"
    End If
End If

Dim b
b = (vNull = True) ' Comparison with Null yields Null, not True/False. [web:4][web:27]
Response.Write "Result of (Null = True) assigned to b: " & SafeCStr(b) & "<br>"

'--- On Error Resume Next abuse & clearing Err --------------------------------
Response.Write "<h2>Err Handling Pitfalls</h2>"

Dim x, y, z
x = "abc"
y = 2

On Error Resume Next
z = x / y
Response.Write "After invalid division x / y, z = """ & SafeCStr(z) & """ (likely empty).<br>"

If Err.Number <> 0 Then
    Response.Write "Err after division: " & Err.Number & " - " & Err.Description & "<br>"
    Err.Clear
    On Error GoTo 0
End If

'--- String concatenation vs numeric addition ---------------------------------
Response.Write "<h2>&amp; vs + for Concatenation</h2>"

Dim a, bNum
a    = "5"
bNum = 10

Response.Write """5"" + 10 = " & (a + bNum) & " (numeric addition, becomes 15).<br>"
Response.Write """5"" &amp; 10 = " & (a & bNum) & " (string concatenation, becomes 510).<br>"

'--- Option Explicit and implicit Variants ------------------------------------
Response.Write "<h2>Option Explicit &amp; Implicit Variants</h2>"

Dim countItems
countItems = 5
' Typo: coumtItems instead of countItems
coumtItems = 10

Response.Write "countItems = " & countItems & " (5). coumtItems = " & coumtItems & " (typo created new variable).<br>"

'--- Returning Nothing vs Null from functions (objects vs values) -------------
Response.Write "<h2>Function Return: Nothing vs Null</h2>"

Function GetFakeObject(shouldReturn)
    If shouldReturn Then
        Dim fake
        Set fake = Server.CreateObject("Scripting.Dictionary")
        Set GetFakeObject = fake
    Else
        Set GetFakeObject = Nothing
    End If
End Function

Dim o1, o2
Set o1 = GetFakeObject(True)
Set o2 = GetFakeObject(False)

If o1 Is Nothing Then
    Response.Write "o1 Is Nothing (unexpected).<br>"
Else
    Response.Write "o1 is a real object: " & TypeName(o1) & "<br>"
End If

If o2 Is Nothing Then
    Response.Write "o2 Is Nothing (correct for 'no object').<br>"
Else
    Response.Write "o2 is NOT Nothing (unexpected).<br>"
End If

'--- Arrays, UBound and Redim Preserve quirks ---------------------------------
Response.Write "<h2>Arrays, UBound and Redim Preserve</h2>"

Dim arr()
ReDim arr(2)
arr(0) = "a"
arr(1) = "b"
arr(2) = "c"

Response.Write "UBound(arr) before Redim Preserve: " & UBound(arr) & "<br>"

ReDim Preserve arr(1)
Response.Write "After Redim Preserve arr(1), UBound(arr) = " & UBound(arr) & " and elements: "
Response.Write arr(0) & ", " & arr(1) & "<br>"

Dim i
For i = 0 To UBound(arr)
    Response.Write "Index " & i & " = " & arr(i) & "<br>"
Next

'--- Date and string parsing ---------------------------------------------------
Response.Write "<h2>Date and String Parsing Surprises</h2>"

Dim d
d = "01/02/2026"
Response.Write "Literal ""01/02/2026"" interpreted as: " & SafeCStr(CDate(d)) & " (depends on server locale).<br>"

'--- Response.Write vs inline performance quirk -------------------------
Response.Write "<h2>Response.Write vs Inline Expressions</h2>"

Dim s, ch
s = ""
For i = Asc("A") To Asc("Z")
    ch = Chr(i)
    s = s & ch
Next

Response.Write "Naively concatenated alphabet: " & s & "<br>"

'--- Form handling & missing keys ---------------------------------------------
Response.Write "<h2>Form Collection Quirks</h2>"

Dim formMissing
formMissing = Request("thisFieldDoesNotExist")
If formMissing = "" Then
    Response.Write "Request(""thisFieldDoesNotExist"") = """" (silent empty string).<br>"
End If

'--- Classic ASP Session / Application locking (conceptual mention) -----------
Response.Write "<h2>Session / Application Locking (Mention)</h2>"
Response.Write "Remember Session and Application are single‑threaded per scope; long operations while locked can stall other requests.<br>"

Response.Write "<hr><p>End of quirks demo.</p>"

%>
