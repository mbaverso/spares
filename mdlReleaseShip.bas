Attribute VB_Name = "mdlReleaseShip"
Option Explicit

Public Sub SendEmail(oItems As CLineItems)
    Dim OutApp As New Outlook.Application
    Dim OutMail As Outlook.MailItem

    If g_bHandleErrors Then On Error GoTo ErrorHandler
    PushCallStack "SendEmail"
    
    Set OutMail = OutApp.CreateItem(olMailItem)

    GetRecipients OutMail, oItems
    
    With OutMail
        .BodyFormat = olFormatHTML
        .Subject = oItems.Item(1).Contract
        .HTMLBody = CreateEmailBody(oItems)
        .Display
    End With
    
ProcExit:
    Set OutApp = Nothing
    PopCallStack
    Exit Sub

ErrorHandler:
    GlobalErrHandler
    Resume ProcExit
    
End Sub

Private Function CreateEmailBody(oItems As CLineItems) As String
    Dim sMessage As String
    Dim oItem As CLineItem
    Dim sPending As String
    Dim oSheet As Worksheet
    Dim sDRN As String
    Dim sSignature As String
    
    'TODO Refactor this to improve formatting
    sSignature = "<br><p><font size=2 face=" & Chr(34) & "Arial" & Chr(34) & ">Regards,<br><br>Michael A.Baverso<br>Project Engineer, IBLE - A<br>AREVA Inc. (External)<br>" & _
        "100 East Kensinger Drive, Suite 100<br>Cranberry Township, PA 16066<br>Phone: 724.591.7046<br>" & _
        "Fax: 434.382.4106</p></font>"
        
    Set oSheet = Sheets("Variables")
    
    sMessage = "<HTML><body><p><font size=2 face=" & Chr(34) & "Arial" & Chr(34) & ">Engineering release to ship" & _
     IIf(oItems.DRNReqd, " pending final approval in Documentum of attached QADP", "") & "<br><br></p></font>"

    sMessage = sMessage & "<table border=1 cellpadding=3 cellspacing=0><font size=1 face=" & Chr(34) & "Arial" & Chr(34) & _
             "><b><tr><td>Item Number</td><td>Part Number</td><td>Description</td><td>Quantity</td>" & _
             "<td>C of C</td></tr></font>"
  
    For Each oItem In oItems
        If oItem.OKtoShip Then
            sMessage = sMessage & "<font size=1 face=" & Chr(34) & "Arial" & Chr(34) & "><tr><td>" & oItem.ItemNumber & "</td>" & "<td>" & oItem.PartNumber & "</td><td>" & oItem.description & "</td><td>" & _
            oItem.Quantity & "</td><td>" & IIf(oItem.CoCReqd, "Y", "N") & " </td></tr></font>"
        End If
    Next
    
    sMessage = sMessage & "</table>"
    
    sMessage = sMessage & IIf(oItems.DRNReqd, "<p><font size=2 face=" & _
        Chr(34) & "Arial" & Chr(34) & ">" & oSheet.Range("C12") & ",<br>" & _
        DRN & "<br><br>" & oSheet.Range("C10") & ",<br>" & SPEC_INS_TAMPER & "</p>", "") & _
        sSignature & "</body></HTML>"
    
    CreateEmailBody = sMessage

End Function

Private Sub GetRecipients(oEmail As Outlook.MailItem, oItems As CLineItems)
    Dim oSheet As Worksheet
    Dim oItem As CLineItem
    
    Set oSheet = Sheets("Variables")
    
    With oSheet
        oEmail.To = .Range("B10").Value & IIf(oItems.DRNReqd, ";" & .Range("B12").Value, "")
        oEmail.Cc = .Range("B11").Value
    End With

End Sub
