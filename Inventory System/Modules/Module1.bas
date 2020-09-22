Attribute VB_Name = "Module1"
Public db As New ADODB.Connection
Public comp_db As New ADODB.Connection



Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As meminfo_status)
Public COM_RS As New ADODB.Recordset '(SRARCH RESULT)




Public Type meminfo_status
    dwlength As Long
    dwmemoryload As Long
    dwtotalphy As Long
    dwavaiphy As Long
    dwtotalpagefile As Long
    dwavaipagefile As Long
    dwtotalvirtual As Long
    dwavailabelvirtual As Long
End Type

Public meminfo As meminfo_status


Public Function SALES_INVOICE_NUMBER() As String
Dim ino As New ADODB.Recordset
Dim INO_COUNT As New ADODB.Recordset

ino.Open "SELECT * FROM SORTEDINVOICE_NO", db, adOpenDynamic, adLockOptimistic
INO_COUNT.Open "SELECT COUNT(*) FROM SORTEDINVOICE_NO", db, adOpenDynamic, adLockOptimistic
If INO_COUNT.Fields(0).Value = 0 Then
    SALES_INVOICE_NUMBER = "COMP0001"
Else
    
    
    ino.MoveLast
    
    While ino.BOF <> True
    
    
    If Mid(ino.Fields(0).Value, 1, 4) <> "WITH" Then
        Dim ST As String
        ST = ino.Fields(0).Value
    
        Dim N As Integer
        N = Mid(ST, 5, Len(ino.Fields(0).Value))
        N = N + 1
    
        Dim NO As String
        NO = N
        ino.Close
    
        If Len(NO) = 1 Then
            SALES_INVOICE_NUMBER = "COMP000" & NO
        ElseIf Len(NO) = 2 Then
            SALES_INVOICE_NUMBER = "COMP00" & NO
        ElseIf Len(NO) = 3 Then
            SALES_INVOICE_NUMBER = "COMP0" & NO
        ElseIf Len(NO) = 4 Then
            SALES_INVOICE_NUMBER = "COMP0" & NO
        End If
        
        Exit Function
    End If
    
    ino.MovePrevious
    If ino.BOF = True Then
        SALES_INVOICE_NUMBER = "COMP0001"
        Exit Function
    End If
    
    Wend
    
End If

End Function

Public Function SYSTEM_NO() As String
Dim SNO As New ADODB.Recordset
Dim SNO_COUNT As New ADODB.Recordset


SNO.Open "SELECT * FROM SORTED_SYSTEM_ID", db, adOpenDynamic, adLockOptimistic
SNO_COUNT.Open "SELECT COUNT(*) FROM SORTED_SYSTEM_ID", db, adOpenDynamic, adLockOptimistic
If SNO_COUNT.Fields(0).Value = 0 Then
    SYSTEM_NO = "S00001"
Else
    SNO.MoveLast
    Dim ST As String
    ST = SNO.Fields(0).Value
    
    Dim N As Integer
    N = Mid(ST, 2, Len(SNO.Fields(0).Value))
    
    N = N + 1
    Dim NO As String
    NO = N
    SNO.Close
    
    If Len(NO) = 1 Then
        SYSTEM_NO = "S0000" & NO
    ElseIf Len(NO) = 2 Then
        SYSTEM_NO = "S000" & NO
    ElseIf Len(NO) = 3 Then
        SYSTEM_NO = "S00" & NO
    ElseIf Len(NO) = 4 Then
        SYSTEM_NO = "S0" & NO
    ElseIf Len(NO) = 5 Then
        SYSTEM_NO = "S" & NO
    End If
    
End If
End Function



Public Function TOTAL_AMT(TRAN_TYPE As String) As Double
    Dim TOTAL_RS As New ADODB.Recordset
    Dim t As Double
    t = 0
    If TRAN_TYPE = "SALES" Then
            TOTAL_RS.Open "SELECT * FROM SYS_CURRENT_SALES_ITEMS", db, adOpenDynamic, adLockOptimistic
            While TOTAL_RS.EOF <> True
                t = t + VAL(TOTAL_RS.Fields(3).Value)
                TOTAL_RS.MoveNext
            Wend
            TOTAL_AMT = t
    ElseIf TRAN_TYPE = "PURCHASE" Then
            TOTAL_RS.Open "SELECT * FROM SYS_CURRENT_INVOICE", db, adOpenDynamic, adLockOptimistic
            While TOTAL_RS.EOF <> True
                t = t + VAL(TOTAL_RS.Fields(4).Value)
                TOTAL_RS.MoveNext
            Wend
            TOTAL_AMT = t
    End If
End Function
