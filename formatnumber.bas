Attribute VB_Name = "formatnumber"
Option Explicit

Sub Format()
Attribute Format.VB_ProcData.VB_Invoke_Func = " \n14"
Dim column1, column2, column3, column4, column5, column6, column7, column8, column9, column10, column11, column12, column13, column14 As String

column1 = Worksheets("frmsettings").Range("D7").Value
column2 = Worksheets("frmsettings").Range("D8").Value
column3 = Worksheets("frmsettings").Range("D9").Value
column4 = Worksheets("frmsettings").Range("D10").Value
column5 = Worksheets("frmsettings").Range("D11").Value
column6 = Worksheets("frmsettings").Range("D12").Value
column7 = Worksheets("frmsettings").Range("D13").Value
column8 = Worksheets("frmsettings").Range("D14").Value
column9 = Worksheets("frmsettings").Range("D15").Value
column10 = Worksheets("frmsettings").Range("D16").Value
column11 = Worksheets("frmsettings").Range("D17").Value
column12 = Worksheets("frmsettings").Range("D18").Value
column13 = Worksheets("frmsettings").Range("D19").Value
column14 = Worksheets("frmsettings").Range("D20").Value

If Len(column1) > 0 Then
    
    If Worksheets("frmsettings").Range("E7").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column1 & ":" & column1).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E7").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column1 & ":" & column1).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E7").Value = "DATE" Then

        Worksheets("MAIN").Columns(column1 & ":" & column1).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E7").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column1 & ":" & column1).NumberFormat = "General"
     Else
     End If
       Else
    End If
    
If Len(column2) > 0 Then
    
    If Worksheets("frmsettings").Range("E8").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column2 & ":" & column2).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E8").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column2 & ":" & column2).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E8").Value = "DATE" Then

        Worksheets("MAIN").Columns(column2 & ":" & column2).NumberFormat = "yyyy-mm-dd"
        
     ElseIf Worksheets("frmsettings").Range("E8").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column2 & ":" & column2).NumberFormat = "General"
     Else
        
    End If
     Else
    End If

If Len(column3) > 0 Then
    
    If Worksheets("frmsettings").Range("E9").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column3 & ":" & column3).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E9").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column3 & ":" & column3).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E9").Value = "DATE" Then

        Worksheets("MAIN").Columns(column3 & ":" & column3).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E9").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column3 & ":" & column3).NumberFormat = "General"
     Else
        
    End If
      Else
    End If

If Len(column4) > 0 Then
    
    If Worksheets("frmsettings").Range("E10").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column4 & ":" & column4).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E10").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column4 & ":" & column4).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E10").Value = "DATE" Then

        Worksheets("MAIN").Columns(column4 & ":" & column4).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E10").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column4 & ":" & column4).NumberFormat = "General"
     Else
        
    End If
    
      Else
    End If

If Len(column5) > 0 Then
    
    If Worksheets("frmsettings").Range("E11").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column5 & ":" & column5).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E11").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column5 & ":" & column5).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E11").Value = "DATE" Then

        Worksheets("MAIN").Columns(column5 & ":" & column5).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E11").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column5 & ":" & column5).NumberFormat = "General"
     Else
        
    End If
    Else
    End If
    
    
If Len(column6) > 0 Then
    
    If Worksheets("frmsettings").Range("E12").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column6 & ":" & column6).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E12").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column6 & ":" & column6).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E12").Value = "DATE" Then

        Worksheets("MAIN").Columns(column6 & ":" & column6).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E12").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column6 & ":" & column6).NumberFormat = "General"
     Else
        
    End If
    
      Else
    End If
    
    
If Len(column7) > 0 Then
    
    If Worksheets("frmsettings").Range("E13").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column7 & ":" & column7).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E13").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column7 & ":" & column7).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E13").Value = "DATE" Then

        Worksheets("MAIN").Columns(column7 & ":" & column7).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E13").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column7 & ":" & column7).NumberFormat = "General"
     Else
        
    End If
    
     Else
    End If
    
    
If Len(column8) > 0 Then
    
    If Worksheets("frmsettings").Range("E14").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column8 & ":" & column8).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E14").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column8 & ":" & column8).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E14").Value = "DATE" Then

        Worksheets("MAIN").Columns(column8 & ":" & column8).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E14").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column8 & ":" & column8).NumberFormat = "General"
     Else
        
    End If
    
  Else
    End If
    
    
If Len(column9) > 0 Then
    
    If Worksheets("frmsettings").Range("E15").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column9 & ":" & column9).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E15").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column9 & ":" & column9).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E15").Value = "DATE" Then

        Worksheets("MAIN").Columns(column9 & ":" & column9).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E15").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column9 & ":" & column9).NumberFormat = "General"
     Else
        
    End If
    
      Else
    End If
    
If Len(column10) > 0 Then
    
    If Worksheets("frmsettings").Range("E16").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column10 & ":" & column10).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E16").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column10 & ":" & column10).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E16").Value = "DATE" Then

        Worksheets("MAIN").Columns(column10 & ":" & column10).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E16").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column10 & ":" & column10).NumberFormat = "General"
     Else
        
    End If
    
     Else
    End If
    
If Len(column11) > 0 Then
    
    If Worksheets("frmsettings").Range("E17").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column11 & ":" & column11).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E17").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column11 & ":" & column11).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E17").Value = "DATE" Then

        Worksheets("MAIN").Columns(column11 & ":" & column11).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E17").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column11 & ":" & column11).NumberFormat = "General"
     Else
        
    End If
      Else
    End If
    
If Len(column12) > 0 Then
    
    If Worksheets("frmsettings").Range("E18").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column12 & ":" & column12).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E18").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column12 & ":" & column12).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E18").Value = "DATE" Then

        Worksheets("MAIN").Columns(column12 & ":" & column12).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E18").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column12 & ":" & column12).NumberFormat = "General"
     Else
        
    End If
    
      Else
    End If
    

If Len(column13) > 0 Then
    
    If Worksheets("frmsettings").Range("E19").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column13 & ":" & column13).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E19").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column13 & ":" & column13).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E19").Value = "DATE" Then

        Worksheets("MAIN").Columns(column13 & ":" & column13).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E19").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column13 & ":" & column13).NumberFormat = "General"
     Else
        
    End If
    Else
    End If
    
 If Len(column14) > 0 Then
    
    If Worksheets("frmsettings").Range("E20").Value = "NUMBER" Then

        Worksheets("MAIN").Columns(column14 & ":" & column14).NumberFormat = "#,##0.00"
        
    ElseIf Worksheets("frmsettings").Range("E20").Value = "TEXT" Then

        Worksheets("MAIN").Columns(column14 & ":" & column14).NumberFormat = "@"
        
     ElseIf Worksheets("frmsettings").Range("E20").Value = "DATE" Then

        Worksheets("MAIN").Columns(column14 & ":" & column14).NumberFormat = "yyyy-mm-dd;@"
        
     ElseIf Worksheets("frmsettings").Range("E20").Value = "GENERAL" Then

        Worksheets("MAIN").Columns(column14 & ":" & column14).NumberFormat = "General"
     Else
        
    End If
    
    Else
    End If
End Sub
