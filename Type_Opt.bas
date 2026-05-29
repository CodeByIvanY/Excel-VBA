Option Explicit
Type ClientOpt
    ClientName As String
    Data_1 As Long
    Data_2 As String
End Type

Function GetClient(ByVal clientKey As String) As ClientOpt
    Dim result As ClientOpt
    
    Select Case LCase(clientKey)
        Case "ivan"
            result.ClientName = "Ivan Yang"
            result.Data_1 = 2
            result.Data_2 = "Male"
            
        Case "kelly"
            result.ClientName = "Kelly Yang"
            result.Data_1 = 2
            result.Data_2 = "Female"
            
        Case Else
            result.ClientName = "Unknown"
            result.Data_1 = 0
            result.Data_2 = "N/A"
    End Select
    
    GetClient = result
End Function

Sub Test()
    Dim client As ClientOpt
    
    client = GetClient("Ivan")
    MsgBox client.ClientName, vbInformation, "Client Name"
End Sub

