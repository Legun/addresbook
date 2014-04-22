'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ������� ������ ��������� �� �������� ������������� Active directory '
'          ��������� ������� ������ �� ������� �� ��������:           '
'              � ��������� ������ ��� ����������� �����!              '
'   ������ ������� ������� ���� ������������ � 2013 ����. ������ 3.   '
'              ������ ���� � �����: http://ithelp.moy.su              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
On error resume next ' ��������� ������

' ��������� ����������
Dim objRootDSE, strDNSDomain, strBase
Dim adoCommand, adoConnection, objRS, strFilter, strAttributes, strQuery
Dim objExcel, strName, strPhone, strMail, strOtherphone, arrOtherPhone, strMobile, strCountry, strCity, strCompany, strDepartment, strTitle
Dim strGivenName, strSN, strDisplayName, strItem
Dim objRoot, objOU, objDomain, objContact
Dim strDNS, strContainer, strContactName, strEmail
Dim UserN, UserD, UNK, UDK
UNK = 0
UDK = 0
UserN = ""
UserD = ""

' ����, �������
' ���������� ��� ������
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection
strBase = "<LDAP://" & strDNSDomain & ">"

' ����� ��� �������� ������� ������.
' ���� UseraccountControl ���: http://support.microsoft.com/kb/305144
' ��� ���� ���������� ���: http://www.computerperformance.co.uk/Logon/LDAP_attributes_active_directory.htm
strFilter = "(&(objectCategory=person)(objectClass=user)(|(useraccountControl=66048)(useraccountcontrol=512)))" 'useraccountcontrol=512 - ���������� ������������!
' � ���������� strAttributes ����������� ����������� ������ ���������, ������� �����������, ����� �������� �� �����!
strAttributes = "name,mail,telephoneNumber,otherTelephone,mobile,c,L,company,department,title,givenname,sn,displayname"

' ��������� ������ �������.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

' �������� ������.
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 500 ' ������������ ���������� �������������, ������� ������� � ������ ������ � AD!
adoCommand.Properties("Timeout") = 307
adoCommand.Properties("Cache Results") = False
Set objRS = adoCommand.Execute

' ���������� ���������� ������ �� �������
While not objRS.EOF
    strName = objRS.Fields("name").Value
    strMail = objRS.Fields("mail").value
    strPhone = objRS.Fields("telephoneNumber").Value
        strPhone = Replace(strPhone,"-","") '���������� ���� �� ��������
        strPhone = Replace(strPhone," ","") '���������� ������� �� ��������   
    arrOtherPhone = objRS.Fields("otherTelephone").Value
    strMobile = objRS.Fields("mobile").Value
        strMobile = Replace(strMobile,"-","") '���������� ���� �� ��������
        strMobile = Replace(strMobile," ","") '���������� ������� �� ��������
    strCountry = objRS.Fields("c").Value
    strCity = objRS.Fields("L").Value
    strCompany = objRS.Fields("company").Value
    strDepartment = objRS.Fields("department").Value
    strTitle = objRS.Fields("title").Value
    strGivenName=objRS.Fields("givenname").Value
    strSN=objRS.Fields("sn").Value
    strDisplayName=objRS.Fields("displayname").Value
    If IsNull(arrOtherPhone) Then
        strOtherPhone = ""
    Else
        strOtherPhone = ""
        For Each strItem In arrOtherPhone
            If (strOtherPhone = "") Then
                strOtherPhone = strItem
            Else
                strItem = Replace(strItem,"-","") '���������� ���� �� ��������
                strItem = Replace(strItem," ","") '���������� ������� �� ��������   
                strOtherPhone = strOtherPhone & " " & strItem
            End If
        Next
    End If
   
    '�������� ����, ���� ���� ����������� �����
    if strMail<>"" then
        Err.Clear
        ' ������ ������� �� ������������. ����������.
        strContainer = "cn=users" '�������� ����������, � ������� ����� ������ ��������, ��������� �� �������������.
        strContactName = "cn=" & strName '�������� ��������
        strEmail = strMail '����� ����������� �����

        ' ������ ������ � Active Directory
        Set objRoot = GetObject("LDAP://rootDSE")
        strDNS = objRoot.Get("defaultNamingContext")
        Set objDomain = GetObject("LDAP://" & strDNS)

        ' ������ ������� �� ������������.
        Set objOU = GetObject("LDAP://"& strContainer & "," & strDNS)
        Set objContact = objOU.Create("contact", strContactName)
        objContact.Put "Mail", strEmail
        if strGivenName <> "" then objContact.Put "givenname", strGivenName
        if strSN <> "" then objContact.Put "sn", strSN
        if strDisplayName <> "" then objContact.Put "displayname", strDisplayName
        if strPhone <> "" then objContact.Put "telephoneNumber", strPhone
        if strOtherPhone <> "" then objContact.Put "otherTelephone", strOtherPhone
        if strMobile <> "" then objContact.Put "mobile", strMobile
        if strCountry <> "" then objContact.Put "c", strCountry
        if strCity <> "" then objContact.Put "L", strCity
        if strCompany <> "" then objContact.Put "company", strCompany
        if strDepartment <> "" then objContact.Put "department", strDepartment
        if strTitle <> "" then objContact.Put "title", strTitle
        objContact.SetInfo '������� ������� ������! (� ������, ���� ����� ��� ��������, �� ��������� ������, ������� ��������������� ��������!)
        If Err.Number = 0 then
            UNK = UNK + 1
            UserN = UserN + strName + chr(13) + chr(10)
        end if
    end if

    objRS.MoveNext
Wend

' � ������ ������� �� AD �������� ��������������� �������������
strFilter = "(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))"' - ����������� ������������!
' � ���������� strAttributes ���������� ����������� ������ ���������, ������� �����������, ����� �������� �� �����!
strAttributes = "name,mail"' � ���� ��� ��� �� ���������� ��������� ���������

' ��������� ������ �������.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

' �������� ������.
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 500
adoCommand.Properties("Timeout") = 307
adoCommand.Properties("Cache Results") = False
Set objRS = adoCommand.Execute
While not objRS.EOF
    strName = objRS.Fields("name").Value
    strMail = objRS.Fields("mail").value
    if strMail<>"" then
        Err.Clear
        '������� ������� ���������������� ������������
        strContainer = "cn=users" '�������� ����������, � ������� ����� ������� �������� ��������������� �������������.
        strContactName = "cn=" & chr(34) & strName & chr(34) '�������� ��������
        Set objOU = GetObject("LDAP://"& strContainer & "," & strDNS)
        a = objOU.Delete ("Contact", strContactName)
        If Err.Number = 500 then
            UDK = UDK + 1
            UserD = UserD + strName + chr(13) + chr(10)
        end if
    end if
    objRS.MoveNext
Wend

'���� ���� ����������� �������������� ����������� � ��������, �� ���� �� ���� �������� ������ �� ����������� �����

If (UNK + UDK) > 0 then
    Dim objEmail, MSG
    MSG = "�������� ������ �������, ������������� ���������� �������� �� ������������� � AD:" + chr(13) + chr(10) + chr(13) + chr(10)
    MSG = MSG + "�������� �������� ���������:" + chr(13) + chr(10)
    MSG = MSG + UserN
    MSG = MSG + "����� ������� ���������: " + CStr(UNK) + chr(13) + chr(10) + chr(13) + chr(10)
    MSG = MSG + "�������� �������� ���������:" + chr(13) + chr(10)
    MSG = MSG + UserD
    MSG = MSG + "����� ������� ���������: " + CStr(UDK) + chr(13) + chr(10)+ chr(13) + chr(10)
    MSG = MSG + "�����/���� ��������� �������: " + CStr(Time) + "/" + CStr(Date) + chr(13) + chr(10)
   
    Const EmailFrom = "bot@firma.ru"         ' �� ���� ����� ������������ e-mail
    Const EmailPassword = "SuperPassword"       ' ������ �� e-mail
    Const strSmtpServer = "smtp.firma.ru" ' smtp ������
    Const EmailTo = "admin@firma.ru"    ' ���� ����� ������������ e-mail
    Set objEmail = CreateObject("CDO.Message")
   
    objEmail.From = EmailFrom
    objEmail.To = EmailTo
    objEmail.Subject = "����� �� ������ ������� MailFromUsers" ' ���� ������
    objEmail.Textbody = MSG
    objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = EmailFrom
    objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = EmailPassword
    objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSmtpServer
    objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    objEmail.Configuration.Fields.Update
    objEmail.Send
    Set objEmail = Nothing ' ������ ������.
End If

' �� ��� � ��! ������ ������ � ������� �� �������!
Set objRS = Nothing
Set adoCommand = Nothing
Set adoConnection = Nothing
Set objOU = Nothing
WScript.Quit
