Attribute VB_Name = "ModuloApiCnpj"
Public Type T_DadosCnpjResponse
   Nome As String
   CodigoIbge As String
   Logradouro As String
   Numero As String
   Bairro As String
   Cidade As String
   Uf As String
   Complemento As String
   Cep As String
   Ddd As String
   Telefone As String
   Email As String
   InscricaoEstadual As String
End Type

Const CONST_URL_API As String = "https://open.cnpja.com/office/"

Private Function GetCnpjData(ByVal P_Cnpj As String) As String

   On Error GoTo TrataErro
   Dim HTTP As Object
   Dim url As String
   Dim Response As String
   
   url = CONST_URL_API & P_Cnpj
   
   Set HTTP = CreateObject("Microsoft.XmlHttp")
   HTTP.open "GET", url, False
   HTTP.send
   
   If HTTP.Status = 200 Then
      Response = HTTP.responseText
   Else
      Response = ""
   End If
   
   GetCnpjData = Response
   
   Exit Function
   
TrataErro:
   MsgBox "Erro na requisição: " & Err.Description, vbCritical, "Erro"
   GetCnpjData = ""
   
End Function

Private Function Nz(Valor As Variant, Optional Padrao As String = "") As Variant
    If IsNull(Valor) Or IsEmpty(Valor) Then
        Nz = Padrao
    Else
        Nz = Valor
    End If
End Function

Public Function PreencherDadosCnpj(ByRef P_Cnpj As String) As T_DadosCnpjResponse
   On Error GoTo TrataErro
   
   Dim JsonObj As Object
   Dim JsonText As String
   Dim DadosCnpjResponse As T_DadosCnpjResponse
   
   JsonText = GetCnpjData(P_Cnpj)
   
   If Len(JsonText) = 0 Then
      MsgBox "Falha ao obter dados do CNPJ", vbExclamation, "Erro"
      Exit Function
   End If
   
   Set JsonObj = JSON.parse(JsonText)
   
   With DadosCnpjResponse
      .Nome = UCase(Nz(JsonObj.Item("company").Item("name"), ""))
        .CodigoIbge = Nz(JsonObj.Item("address").Item("municipality"), "")
        .Logradouro = UCase(Nz(JsonObj.Item("address").Item("street"), ""))
        .Numero = Nz(JsonObj.Item("address").Item("number"), "")
        .Bairro = UCase(Nz(JsonObj.Item("address").Item("district"), ""))
        .Cidade = UCase(Nz(JsonObj.Item("address").Item("city"), ""))
        .Uf = UCase(Nz(JsonObj.Item("address").Item("state"), ""))
        .Complemento = UCase(Nz(JsonObj.Item("address").Item("details"), ""))
        .Cep = Nz(JsonObj.Item("address").Item("zip"), "")
        
        If Not JsonObj.Item("phones") Is Nothing And JsonObj.Item("phones").Count > 0 Then
            .Ddd = Nz(JsonObj.Item("phones")(1).Item("area"), "")
            .Telefone = Nz(JsonObj.Item("phones")(1).Item("number"), "")
        Else
            .Ddd = ""
            .Telefone = ""
        End If

        If Not JsonObj.Item("emails") Is Nothing And JsonObj.Item("emails").Count > 0 Then
            .Email = Nz(JsonObj.Item("emails")(1).Item("address"), "")
        Else
            .Email = ""
        End If

        If Not JsonObj.Item("registrations") Is Nothing And JsonObj.Item("registrations").Count > 0 Then
            .InscricaoEstadual = Nz(JsonObj.Item("registrations")(1).Item("number"), "")
        Else
            .InscricaoEstadual = ""
        End If
   End With
   
   PreencherDadosCnpj = DadosCnpjResponse
   
   Exit Function
   
TrataErro:
   MsgBox "Erro: " & Err.Description, vbCritical, "Erro"
   
End Function







