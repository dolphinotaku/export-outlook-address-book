Sub Main()

Dim addressBookName As String
Dim cvsLinesList() As String
Dim totalMatchedSearch As Integer
    
addressBookName = "Contacts" 'please change to your address book name, by default is Contacts

cvsLinesList = GetAddressBook(addressBookName)
totalMatchedSearch = (UBound(cvsLinesList) - LBound(cvsLinesList) + 1) - 1
    
If Not totalMatchedSearch = 0 Then
    ExportAddressBook cvsLinesList
    MsgBox "File is successfully generated" & vbCrLf & vbCrLf & "Total records: " & totalMatchedSearch, vbOKOnly, "Success"
End If

End Sub

Public Function ExportAddressBook(cvsLinesList() As String)

    Dim curDateTime As String
    curDateTime = Format(Now(), "yyyyMMdd_hhmmss")

    ' the path cannot without &
    outfile = "d:" & "\contact." & curDateTime & ".csv" 'please to your path
    Open outfile For Output As #2
    
    Dim cvsLine As Variant
    For Each cvsLine In cvsLinesList
        Print #2, cvsLine
    Next cvsLine
    
    Close #2

End Function

Public Function GetAddressBook(addressListName As String) As String()

' enable error handler
On Error GoTo ErrorHandler

If Len(addressListName) <= 0 Then
    addressListName = "Global Address List"
End If
    
    Dim colAL As Outlook.AddressLists
    Dim oAL As Outlook.AddressList
    Dim colAE As Outlook.AddressEntries
    Dim oAE As Outlook.addressEntry
    Dim oExUser As Outlook.exchangeUser
    
    Set colAL = Application.Session.AddressLists
    
    Dim cvsHeading As String
    Dim cvsLine As String
    Dim textLinesList() As String
    ReDim Preserve textLinesList(0)
    
    ' prepare the cvs header line
    cvsHeading = AppendCvsValue(cvsHeading, "Address")
    cvsHeading = AppendCvsValue(cvsHeading, "AddressEntryUserType")
    cvsHeading = AppendCvsValue(cvsHeading, "alias")
    cvsHeading = AppendCvsValue(cvsHeading, "AssistantName")
    cvsHeading = AppendCvsValue(cvsHeading, "businessTelephoneNumber")
    cvsHeading = AppendCvsValue(cvsHeading, "city")
    cvsHeading = AppendCvsValue(cvsHeading, "Comments")
    cvsHeading = AppendCvsValue(cvsHeading, "companyName")
    cvsHeading = AppendCvsValue(cvsHeading, "department")
    cvsHeading = AppendCvsValue(cvsHeading, "DisplayType")
    cvsHeading = AppendCvsValue(cvsHeading, "firstName")
    cvsHeading = AppendCvsValue(cvsHeading, "ID")
    cvsHeading = AppendCvsValue(cvsHeading, "jobTitle")
    cvsHeading = AppendCvsValue(cvsHeading, "lastName")
    cvsHeading = AppendCvsValue(cvsHeading, "MobileTelephoneNumber")
    cvsHeading = AppendCvsValue(cvsHeading, "name")
    cvsHeading = AppendCvsValue(cvsHeading, "OfficeLocation")
    cvsHeading = AppendCvsValue(cvsHeading, "postalCode")
    cvsHeading = AppendCvsValue(cvsHeading, "primarySmtpAddress")
    cvsHeading = AppendCvsValue(cvsHeading, "streetaddress")
    cvsHeading = AppendCvsValue(cvsHeading, "Type")
    
    cvsHeading = AppendCvsValue(cvsHeading, "oAE.Address")
    cvsHeading = AppendCvsValue(cvsHeading, "oAE.AddressEntryUserType")
    cvsHeading = AppendCvsValue(cvsHeading, "oAE.DisplayType")
    cvsHeading = AppendCvsValue(cvsHeading, "oAE.ID")
    cvsHeading = AppendCvsValue(cvsHeading, "oAE.name")
    cvsHeading = AppendCvsValue(cvsHeading, "oAE.Type")
                            
                            
    textLinesList(0) = cvsHeading
    
    Dim counter As Integer
    Dim validAddressEntrycounter As Integer
    
    For Each oAL In colAL
        If oAL.name = addressListName Then
            
            'Address list is an Exchange Global Address List
            'If oAL.AddressListType = olExchangeGlobalAddressList Then
                Set colAE = oAL.AddressEntries
                For Each oAE In colAE
                    counter = counter + 1
                    cvsLine = ""
                    
                    ' exclude the group, only the user type is 0 or 5 allowed to call GetExchangeUser
'olExchangeAgentAddressEntry 3 An address entry that is an Exchange agent.
'olExchangeDistributionListAddressEntry 1 An address entry that is an Exchange distribution list.
'olExchangeOrganizationAddressEntry 4 An address entry that is an Exchange organization.
'olExchangePublicFolderAddressEntry 2 An address entry that is an Exchange public folder.
'olExchangeRemoteUserAddressEntry 5 An Exchange user that belongs to a different Exchange forest.
'olExchangeUserAddressEntry 0 An Exchange user that belongs to the same Exchange forest.
'olLdapAddressEntry 20 An address entry that uses the Lightweight Directory Access Protocol (LDAP).
'olOtherAddressEntry 40 A custom or some other type of address entry such as FAX.
'olOutlookContactAddressEntry 10 An address entry in an Outlook Contacts folder.
'olOutlookDistributionListAddressEntry 11 An address entry that is an Outlook distribution list.
'olSmtpAddressEntry 30 An address entry that uses the Simple Mail Transfer Protocol (SMTP).
                    If oAE.AddressEntryUserType = _
                        olExchangeUserAddressEntry _
                        Or oAE.AddressEntryUserType = _
                        olExchangeRemoteUserAddressEntry Then
                        
                        validAddressEntrycounter = validAddressEntrycounter + 1
                        Set oExUser = oAE.GetExchangeUser
                            
                            cvsLine = AppendCvsValue(cvsLine, oExUser.Address)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.AddressEntryUserType)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.alias)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.AssistantName)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.businessTelephoneNumber)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.city)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.Comments)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.companyName)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.department)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.DisplayType)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.firstName)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.ID)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.jobTitle)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.lastName)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.MobileTelephoneNumber)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.name)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.OfficeLocation)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.postalCode)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.primarySmtpAddress)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.streetaddress)
                            cvsLine = AppendCvsValue(cvsLine, oExUser.Type)
                            
                            cvsLine = AppendCvsValue(cvsLine, oAE.Address)
                            cvsLine = AppendCvsValue(cvsLine, oAE.AddressEntryUserType)
                            cvsLine = AppendCvsValue(cvsLine, oAE.DisplayType)
                            cvsLine = AppendCvsValue(cvsLine, oAE.ID)
                            cvsLine = AppendCvsValue(cvsLine, oAE.name)
                            cvsLine = AppendCvsValue(cvsLine, oAE.Type)
                            
                            ReDim Preserve textLinesList(UBound(textLinesList) + 1)
                            textLinesList(UBound(textLinesList)) = cvsLine
                            
                    End If
                Next
            'End If
        End If
    Next
    
    GetAddressBook = textLinesList
    
ErrorHandler:
    'MsgBox Err.Number & ":" & Err.Description
    Debug.Print Err.Number & ":" & Err.description
    
    Resume Next
End Function

Public Function AppendCvsValue(line As String, value As String) As String

'If Not IsEmpty(line) Then
If Len(line) > 0 Then
    line = line & "," & Chr(34) & value & Chr(34)
Else
    line = Chr(34) & value & Chr(34)
End If

AppendCvsValue = line

End Function

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
