Option Explicit
on error resume next
Dim HTML, HHead, HCss, HJava, HFuter, strI, SE
HTML = ""
HHead = ""
HScc = ""
HJava = ""
HFuter = ""
strI = 1
SE = chr(13) + chr(10)
HHead = "<html>" + SE
HHead = HHead + "<head>" + SE
HHead = HHead + "<TITLE>���������� ���������� �����������.</TITLE>" + SE
HHead = HHead + "<style> * { font-family:tahoma; font-size:11px;}</style>" + SE
HHead = HHead + "<p align=" + chr(34) + "center" + chr(34) +">" + SE
HHead = HHead + "<IMG src=" + chr(34)+"i/logo.gif" + chr(34) + "/IMG>" + SE
HHead = HHead + "</p>" + SE
HHead = HHead + "<p align=" + chr(34) + "center" + chr(34) +">" + SE
HHead = HHead + "��������� ����������, ������ ���������� ���������� ��������������. ����������� 1 ��� � �����. �������, ���� �� ������� �����-���� ������ � ���������� ������, �� ������ �������� ��� � ���." + SE
HHead = HHead + "<BR>" + SE
HHead = HHead + "��������! ���� ������ �������� �������� �������, ��������� � ���, ��� � ����� ������������ InternetExplorer ������������ ��������� �������� ����������, �� ������ ������� �� ��� ������� � ��������� �������� ����������, ����� ����� �������� ������������������� ���������� � �������." + SE
HHead = HHead + "<BR>" + SE
HHead = HHead + "�������� �����: � ������� ������ ���� ����� ������������ ���������� ��������, �, ���� ��� ���� ���������� ������� Ctrl, �� ����� ���������� ������� ��������!" + SE
HHead = HHead + "</p>" + SE
HCss = HCss + "<style type=" + chr(34) + "text/css" + chr(34) + ">" + SE
HCss = HCss + "table.sort{" + SE
HCss = HCss + "border-spacing:0.1em;" + SE
HCss = HCss + "margin-bottom:1em;" + SE
HCss = HCss + "margin-top:1em" + SE
HCss = HCss + "}" + SE
HCss = HCss + "table.sort td{" + SE
HCss = HCss + "border:0 solid #CCCCCC;" + SE
HCss = HCss + "padding:0.3em 1em" + SE
HCss = HCss + "}" + SE
HCss = HCss + "table.sort thead td{" + SE
HCss = HCss + "cursor:pointer;" + SE
HCss = HCss + "cursor:hand;" + SE
HCss = HCss + "font-weight:bold;" + SE
HCss = HCss + "text-align:center;" + SE
HCss = HCss + "vertical-align:middle" + SE
HCss = HCss + "}" + SE
HCss = HCss + "table.sort thead td.curcol{" + SE
HCss = HCss + "background-color:#999999;" + SE
HCss = HCss + "color:#FFFFFF" + SE
HCss = HCss + "}" + SE
HCss = HCss + "</style>" + SE
HJava = HJava + "<script type=" + chr(34) + "text/javascript" + chr(34) + " src=" + chr(34) + "Sort.js" + chr(34) + "></script>" + SE
HJava = HJava + "<script src=" + chr(34) + "jquery.min.js" + chr(34) + " type=" + chr(34) + "text/javascript" + chr(34) + "></script>" + SE
HJava = HJava + "<script src=" + chr(34) + "jquery.liveFilter.js" + chr(34) + " type=" + chr(34) + "text/javascript" + chr(34) + "></script>" + SE
HJava = HJava + "<script type=" + chr(34) + "text/javascript" + chr(34) + ">" + SE
HJava = HJava + "    $(document).ready(function() {" + SE
HJava = HJava + "    $('table.live_filter').liveFilter('fade');" + SE
HJava = HJava + "    });" + SE
HJava = HJava + "    $(document).ready(function() {" + SE
HJava = HJava + "    $('ul.list_filter').liveFilter('slide');" + SE
HJava = HJava + "    });" + SE
HJava = HJava + "</script>" + SE
Hhml = Html + "</Head>" + SE
html = html + "<body>" + SE
html = html + "<p>������ ������:" + SE
html = html + "<input class=" + chr(34) + "filter" + chr(34) + " type=" + chr(34) + "text" + chr(34) + " value=" + chr(34) + chr(34) + " name=" + chr(34) + "livefilter" + chr(34) + "></input></p>" + SE
html = html + "<p>������ �������� �� ���� �������� �������! �������� - � ������ ������ �� ����������� �������!</p>" + SE
html = html + "<table border=" + chr(34) + "0" + chr(34) + "width=" + chr(34) + "100%" + chr(34) + " cellpadding=" + chr(34) + "11" + chr(34) + " class=" + chr(34) + "sort live_filter" + chr(34) + " align=" + chr(34) + "center" + chr(34)+">" + SE
html = html + "<CAPTION><H1>���������� ���������� �����������.</H1></CAPTION>" + SE
html = html + "<thead>" + SE
html = html + "<tr BGCOLOR=#999900>" + SE
html = html + "<td>��� ����������:</td>" + SE
html = html + "<td>����������� �����:</td>" + SE
html = html + "<td>�������:</td>" + SE
' = html + "<td>������ �����:</td>" + SE
html = html + "<td>���������:</td>" + SE
html = html + "<td>������:</td>" + SE
html = html + "<td>�����:</td>" + SE
html = html + "<td>�����������:</td>" + SE
html = html + "<td>�����:</td>" + SE
html = html + "<td>���������:</td>" + SE
html = html + "</tr>" + SE
html = html + "</thead>" + SE
html = html + "<tbody>" + SE
Dim objRootDSE, strDNSDomain, strBase
Dim adoCommand, adoConnection, objRS, strFilter, strAttributes, strQuery
Dim objExcel, strName, strPhone, strMail, strOtherphone, arrOtherPhone, strMobile, strCountry, strCity, strCompany, strDepartment, strTitle
Dim strGivenName, strSN, strDisplayName, strItem
Dim objRoot, objOU, objDomain, objContact, strYourDescription
Dim strDNS, strContainer, strContactName, strEmail
Dim strCon
Dim ext
ext = chr(160) & "ext"  & chr(160)
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection
strBase = "<LDAP://" & strDNSDomain & ">"
'strFilter = "(&(objectCategory=person)(objectClass=user)(|(useraccountControl=66048)(useraccountcontrol=512)))" '�������� �� �������������
strFilter = "(&(objectCategory=person)(objectClass=contact))" '�������� �� ���������
strAttributes = "name,mail,telephoneNumber,otherTelephone,mobile,c,L,company,department,title,givenname,sn,displayname"
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 500
adoCommand.Properties("Timeout") = 307
adoCommand.Properties("Cache Results") = False
Set objRS = adoCommand.Execute
While not objRS.EOF
    strName = ext
    strMail = ext
    strPhone = ext
    arrOtherPhone = ext
    strMobile = ext
    strCountry = ext
    strCity = ext
    strCompany = ext
    strDepartment = ext
    strTitle = ext
    strGivenName = ext
    strSN = ext
    strDisplayName = ext
    strName = objRS.Fields("name").Value
    strName = Replace(strName," ",chr(160)) '�������� ������� ������ �� �����������
    strMail = objRS.Fields("mail").value
    strPhone = objRS.Fields("telephoneNumber").Value
        strPhone = Replace(strPhone,"-","") '���������� ���� �� ��������
        strPhone = Replace(strPhone," ","") '���������� ������� �� ��������
        a = Left(strPhone,1)
        if a = "8" then strPhone = "+7" + Mid(strPhone,2)
        if a = "(" then strPhone = "+7" + strPhone
    arrOtherPhone = objRS.Fields("otherTelephone").Value
    strMobile = objRS.Fields("mobile").Value
        strMobile = Replace(strMobile,"-","") '���������� ���� �� ��������
        strMobile = Replace(strMobile," ","") '���������� ������� �� ��������
        a = Left(strMobile,1)
        if a = "8" then strMobile = "+7" + Mid(strMobile,2)
        if a = "(" then strMobile = "+7" + strMobile
    strCountry = objRS.Fields("c").Value
    strCity = objRS.Fields("L").Value
    strCompany = objRS.Fields("company").Value
    strDepartment = objRS.Fields("department").Value
        strDepartment = Replace(strDepartment,"�����","") '���������� �� �������� ������ ����� �����,
        strDepartment = Replace(strDepartment,"�����","") '��� ��� ��� ������������ � ����� �������!
        strDepartment = Trim(strDepartment) '������� ������� ������� � ������ �������� ������.
        strDepartment = uCase(Left(strDepartment,1)) + Mid(strDepartment,2) '������ ������ ����� � �������� ���������, ��������� ��������� ��� ���������.
    strTitle = objRS.Fields("title").Value
    strGivenName = objRS.Fields("givenname").Value
    strGivenName = Replace(strName," ",chr(160)) '�������� ������� ������ �� �����������
    strSN = objRS.Fields("sn").Value
    strDisplayName = objRS.Fields("displayname").Value
    a = Instr(strDisplayName,"-")
    if a=0 then    strDisplayName = Replace(strDisplayName," ",chr(160)) '�������� ������� ������ �� �����������, ���� � ��� ��� ������� �������
    If IsNull(arrOtherPhone) Then
        strOtherPhone = ""
    Else
        strOtherPhone = ""
        For Each strItem In arrOtherPhone
            If (strOtherPhone = "") Then
                strItem = Replace(strItem,"-","") '���������� ���� �� ��������
                strItem = Replace(strItem," ","") '���������� ������� �� ��������
                a = Left(strItem,1)
                if a = "8" then strItem = "+7" + Mid(strItem,2)
                if a = "(" then strItem = "+7" + strItem
                strOtherPhone = strItem
            Else
                strItem = Replace(strItem,"-","") '���������� ���� �� ��������
                strItem = Replace(strItem," ","") '���������� ������� �� ��������
                a = Left(strItem,1)
                if a = "8" then strItem = "+7" + Mid(strItem,2)
                if a = "(" then strItem = "+7" + strItem
                strOtherPhone = strOtherPhone & "<br>" & strItem
            End If
        Next
    End If
    If strOtherPhone <> "" then strPhone = strPhone & "<br>" & strOtherPhone
    'if strMail<>"" then '���� �������� �� �������������, �� ������� �����, ����� �� ������������� ��������� ������
        html = html + "<tr BGCOLOR=#CCCCCC>" + SE
        html = html + "<td>" & strDisplayName & "</td>" + SE
        html = html + "<td><a class=" & chr(34) & "link" & chr(34) & " href=" & chr(34) & "mailto:" & strMail & chr(34) & ">" & strMail & "</a></td>" + SE
        html = html + "<td>" & strPhone & "</td>" + SE
        'html = html + "<td>" & strOtherPhone & "</td>" + SE
        html = html + "<td>" & strMobile & "</td>" + SE
        if strCountry = "RU" then strCountry = "������"
        if strCountry = "UA" then strCountry = "�������"
        if strCountry = "BY" then strCountry = "��������"
        if strCountry = "KZ" then strCountry = "���������"
        if strCountry = "GB" then strCountry = "��������������"
        if strCountry = "DE" then strCountry = "��������"
        if strCountry = "IT" then strCountry = "������"
        if strCountry = "FR" then strCountry = "�������"
        if strCountry = "BE" then strCountry = "�������"
        if strCountry = "US" then strCountry = "���"
        html = html + "<td>" & strCountry & "</td>" + SE
        html = html + "<td>" & strCity & "</td>" + SE
        html = html + "<td>" & strCompany & "</td>" + SE
        html = html + "<td>" & strDepartment & "</td>" + SE
        html = html + "<td>" & strTitle & "</td>" + SE
        html = html + "</tr>" + SE
        strI = strI + 1
    'end if '���� �������� �� �������������, �� ������� �����, ����� �� ������������� ��������� ������
    objRS.MoveNext
Wend
Set objRS = Nothing
Set adoCommand = Nothing
Set adoConnection = Nothing
Dim a, b
html = html + "</tbody>" + SE
html = html + "</table>" + SE
html = html + "<p align=" + chr(34) + "center" + chr(34) +">" + SE
strI = strI - 1
a = cstr(strI)
b = right(a,1)
html = html + "����� � ���������� ����������� ���������� " + a + " �����"
select case b
    case "0", "5", "6", "7", "8", "9"
        html = html + "��." + SE
    case "1"
        html = html + "�." + SE
    case "2", "3", "4"
        html = html + "�." + SE
end select
html = html + "<BR>" + SE
html = html + "</p>" + SE
html = html + "</BODY>" + SE
html = html + "</HTML>" + SE
HFuter = HFuter + "<div class=" + chr(34) + "footer" + chr(34) + ">" + SE
HFuter = HFuter + "<p align=" + chr(34) + "center" + chr(34) +">" + SE
HFuter = HFuter + "��������� ����������� ����������� ���� ����������� ������������ ������ ��. � ���������, ��-�����." + SE
HFuter = HFuter + "</p>" + SE
HFuter = HFuter + "</div>" + SE
Dim fso, tf
Set fso = CreateObject("Scripting.FileSystemObject")
Set tf = fso.CreateTextFile("\\FileServer\Contacts\���������� ���������� �����������.htm", True)
tf.WriteLine(HHead)
tf.WriteLine(HCss)
tf.WriteLine(HJava)
tf.WriteLine(HTML)
tf.WriteLine(HFuter)
tf.Close
Set fso = Nothing