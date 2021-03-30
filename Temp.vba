Option Compare Database
Option Explicit

Dim CurrentUserid As String

Sub setCurrentUserid(id As Integer)
    CurrentUserid = id
End Sub

Function getCurrentUserid() As Integer
    getCurrentUserid = CurrentUserid
End Function
    
Function PasswordCompare(CurrentUserid_password As String) As Boolean

    If CurrentUserid_password = Me.User_password Then
        PasswordCompare = True
        DoCmd.Close
    End If
        
End Function
    
Function ErrorMessage(ErrorMessage_Type As Integer)

    If ErrorMessage_Type = 1 Then
        MsgBox ("Incorrect userid")
    
    ElseIf ErrorMessage_Type = 2 Then
        MsgBox ("Incorrect password")
    
    ElseIf ErrorMessage_Type = 3 Then
        MsgBox ("Unknow User")
        
    Else
        MsgBox ("Unknown Error")
        
    End If
    
End Function




Private Sub btn_loginOverlay_Click()


    Dim CurrentUserid_Code As String
    Dim CurrentUserid_password As String
    
    If IsNull(Me.User_id) And IsNull(Me.User_password) Then
        MsgBox "Enter user id and password", vbInformation, "Empty field"
        Me.User_id.SetFocus
    
    ElseIf IsNull(Me.User_id) Then
        MsgBox "Enter a user id", vbInformation, "Empty field"
        Me.User_id.SetFocus
    
    ElseIf IsNull(Me.User_password) Then
        MsgBox "Enter a password", vbInformation, "Empty field"
        Me.User_password.SetFocus
    
    Else
        setCurrentUserid (Nz(Me.User_id, 0))
        CurrentUserid_Code = Left(getCurrentUserid, 1)
        
'=======================================================================
        If (CurrentUserid_Code = "T") Then
            CurrentUserid_password = Nz(DLookup("Teacher_password", "Teacher_information", "Teacher_id ='" & Me.User_id & "'"), "0")
            
            If CurrentUserid_password = 0 Then
                ErrorMessage (1)
                
            Else
                If PasswordCompare(CurrentUserid_password) Then
                    DoCmd.OpenForm "Student_welcome"
            
                Else
                    ErrorMessage (2)
                
                End If
            
            End If
            
'=======================================================================
        ElseIf (CurrentUserid_Code = "A") Then
            CurrentUserid_password = Nz(DLookup("Admin_password", "Admin_information", "Admin_id ='" & Me.User_id & "'"), "0")
        
            If CurrentUserid_password = 0 Then
                ErrorMessage (1)
                
            Else
                If PasswordCompare(CurrentUserid_password) Then
                    DoCmd.OpenForm "Student_welcome"
            
                Else
                    ErrorMessage (2)
                
                End If
            
            End If
        
'=======================================================================
        ElseIf (CurrentUserid_Code = "S") Then
            CurrentUserid_password = Nz(DLookup("Student_password", "Student_information", "Student_id ='" & Me.User_id & "'"), "0")
            
            If CurrentUserid_password = 0 Then
                ErrorMessage (1)
                
            Else
                If PasswordCompare(CurrentUserid_password) Then
                    DoCmd.OpenForm "Student_Personal_general"
            
                Else
                    ErrorMessage (2)
                
                End If
            
            End If
            
        Else
            ErrorMessage (3)
        
        End If
    End If
End Sub


Private Sub Form_Load()

    Dim CurrentTmpGender As String
    Dim TmpTeacherid As String
    Dim TmpTeacherName As String
    Dim TmpStudentid As String

    Me.club_id.Caption = getCurrentTMPid
    Me.club_name.Caption = Nz(DLookup("[Club_name]", "Club_information", "[Club_id] = '" & [getCurrentTMPid] & "'"), 0)
    Me.club_des.Caption = Nz(DLookup("[Club_description]", "Club_information", "[Club_id] = '" & [getCurrentTMPid] & "'"), 0)

    TmpTeacherid = Nz(DLookup("[Teacher_id]", "Club_schoolyear", "[Club_id] = '" & [getCurrentTMPid] & "'"), 0)
    CurrentTmpGender = Nz(DLookup("[Teacher_gender]", "Teacher_information", "[Teacher_id] = '" & [TmpTeacherid] & "'"), 0)
    
    If CurrentTmpGender = "F" Then
        TmpTeacherName = "Ms."
        
    Else
        TmpTeacherName = "Mr."
        
    End If
    
    TmpTeacherName = TmpTeacherName & " " & Nz(DLookup("[Teacher_lastName]", "Teacher_information", "[Teacher_id] = '" & TmpTeacherid & "'"), 0) &" "& Nz(DLookup("[Teacher_firstName]", "Teacher_information", "[Teacher_id] = '" & TmpTeacherid & "'"), 0)
    Me.club_teach.Caption = TmpTeacherName

    TmpStudentid = Nz(DLookup("[Student_id]", "Club_schoolyear", "[Club_id] = '" & [getCurrentTMPid] & "'"), 0)
    Me.club_chairper.Caption = Nz(DLookup("[Student_lastName]", "Student_information", "[Student_id] = '" & TmpStudentid & "'"), 0) &" "& Nz(DLookup("[Student_firstName]", "Student_information", "[Student_id] = '" & TmpStudentid & "'"), 0)

End Sub


