<%


sch = "http://schemas.microsoft.com/cdo/configuration/" 
 
    Set cdoConfig = CreateObject("CDO.Configuration") 
 
    With cdoConfig.Fields 
        .Item(sch & "sendusing") = 2 ' cdoSendUsingPort 
        .Item(sch & "smtpserver") = "smtp8.net4india.com" 
        .update 
    End With 
 
    Set cdoMessage = CreateObject("CDO.Message") 
 
    With cdoMessage 
        Set .Configuration = cdoConfig 
        .From = "Sender Email ID" 
        .To = email 
        .Subject = "Sample CDO Message" 
        .TextBody = "This is a test for CDO.message" 
        .Send 
    End With 
 
    Set cdoMessage = Nothing 
    Set cdoConfig = Nothing 

%>