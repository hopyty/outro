Public Session, conn

Dim rs

Sub ConectarBanco()


    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BD DadosSistema.mdb;"
    Set rs = CreateObject("ADODB.Recordset")
End Sub