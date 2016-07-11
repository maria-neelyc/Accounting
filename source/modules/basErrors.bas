Option Compare Database

Function FnIsErr(DataErr, Optional AddMsg As String) As String
Dim strMsg As String
Select Case DataErr
    Case 2105
        FnIsErr = "You can't live record!" & _
        "Some informations missing or First/Last Record "
    Case 2046
        FnIsErr = "Operation is not available!" 'Undo and Delete
    Case 2107
        FnIsErr = "Incorrect input!"
    Case 2113
        FnIsErr = "Incorrect input!"
    Case 2279
        FnIsErr = "Incorrect input for this field!"
    Case 2237 'Pograsan unos u Combo
        FnIsErr = "Incorrect Value Selected from List!"
    Case 2757 'Ostavljeno prazno polje sa obaveznim unosom
        FnIsErr = "Incorrect input!" & vbLf & "(Duplicated Record or Required Field is Empty)"
    Case 3201
        FnIsErr = "You can't live record! Some informations missing."
    Case 3022
        FnIsErr = "Duplicate Record!"
    Case 3075
        FnIsErr = "No source found"
    Case 3101
        FnIsErr = "Incorect input!"
    Case 2169
        FnIsErr = "Incorrect input!"
    Case 3021
'        Resume Next
        Exit Function
    Case 3162
        FnIsErr = "Incorect input for this fild!"
    Case 3200
        FnIsErr = "Reltion exsist! You can't delete this record!"
    Case 3314
        FnIsErr = "Must fill this field!"
    Case 8519 'kada se brise sa record selector-om
    Case 7001
        FnIsErr = "There is at least one transaction present for account under this or following levels." & Chr(10) _
                & "Cannot remove this reference"
    Case 7002
        FnIsErr = "There is at least one sub-level reference under the referenece you are trying to remove." & Chr(10) _
                & "Cannot remove this reference"
    Case 7003
        FnIsErr = "There is an account joint with this reference." & Chr(10) _
                & "Are you sure you want to remove this reference?"
    Case 7004
        FnIsErr = "You are about to delete this reference entry." & Chr(10) _
                & "Are you sure?"
    Case 7005
        FnIsErr = "Balaces for the last year and opening balances for the new year do not match." & Chr(10) _
                & "There is inconsistency in the database"
    Case 7006
        FnIsErr = "There was an error connecting connecting to database."
    Case 7007
        FnIsErr = "There are transactions present for this account. Cannot delete."
    Case 7008
        FnIsErr = "Next year's account or chart has been already modified and cannot be automatically updated." & Chr(10) _
                & "Any changes have to be inputed manually."
    Case 7009
        FnIsErr = "Next year's account or chart has been already modified. This may cause problems during opening balances update." & Chr(10) _
                & "Please check errorlog for bad entries."
    Case 7010
        FnIsErr = "Document date is not from the period!"
    Case 7011
        FnIsErr = "Account already taken by " & AddMsg
    Case Else
        FnIsErr = DataErr & " " & Err.Description
        Exit Function
End Select
End Function