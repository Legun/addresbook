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
HHead = HHead + "<TITLE>Телефонный справочник сотрудников.</TITLE>" + SE
HHead = HHead + "<style> * { font-family:tahoma; font-size:11px;}</style>" + SE
HHead = HHead + "<p align=" + chr(34) + "center" + chr(34) +">" + SE
HHead = HHead + "<IMG src=" + chr(34)+"i/logo.gif" + chr(34) + "/IMG>" + SE
HHead = HHead + "</p>" + SE
HHead = HHead + "<p align=" + chr(34) + "center" + chr(34) +">" + SE
HHead = HHead + "Уважаемые сотрудники, данный телефонный справочник автоматический. Обновляется 1 раз в сутки. Поэтому, если Вы увидите какие-либо ошибки в контактных данных, то просто сообщите нам о них." + SE
HHead = HHead + "<BR>" + SE
HHead = HHead + "Внимание! Если вверху страницы повилась надпись, говорящая о том, что в целях безопасности InternetExplorer заблокировал некоторое активное содержимое, то просто нажмите на эту надпись и разрешите активное содержимое, тогда будет работать полнофункциональная сортировка в таблице." + SE
HHead = HHead + "<BR>" + SE
HHead = HHead + "Полезный совет: с помощью ролика мыши можно прокручивать содержимое страницы, а, если при этом удерживать клавишу Ctrl, то будет изменяться масштаб страницы!" + SE
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
html = html + "<p>Фильтр поиска:" + SE
html = html + "<input class=" + chr(34) + "filter" + chr(34) + " type=" + chr(34) + "text" + chr(34) + " value=" + chr(34) + chr(34) + " name=" + chr(34) + "livefilter" + chr(34) + "></input></p>" + SE
html = html + "<p>Фильтр работает по всем столбцам таблицы! Внимание - в строке поиска не используйте пробелы!</p>" + SE
html = html + "<table border=" + chr(34) + "0" + chr(34) + "width=" + chr(34) + "100%" + chr(34) + " cellpadding=" + chr(34) + "11" + chr(34) + " class=" + chr(34) + "sort live_filter" + chr(34) + " align=" + chr(34) + "center" + chr(34)+">" + SE
html = html + "<CAPTION><H1>Телефонный справочник сотрудников.</H1></CAPTION>" + SE
html = html + "<thead>" + SE
html = html + "<tr BGCOLOR=#999900>" + SE
html = html + "<td>Имя сотрудника:</td>" + SE
html = html + "<td>Электронная почта:</td>" + SE
html = html + "<td>Телефон:</td>" + SE
' = html + "<td>Другой номер:</td>" + SE
html = html + "<td>Мобильный:</td>" + SE
html = html + "<td>Страна:</td>" + SE
html = html + "<td>Город:</td>" + SE
html = html + "<td>Организация:</td>" + SE
html = html + "<td>Отдел:</td>" + SE
html = html + "<td>Должность:</td>" + SE
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
'strFilter = "(&(objectCategory=person)(objectClass=user)(|(useraccountControl=66048)(useraccountcontrol=512)))" 'Просмотр по пользователям
strFilter = "(&(objectCategory=person)(objectClass=contact))" 'Просмотр по контактам
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
    strName = Replace(strName," ",chr(160)) 'Заменяем обычный пробел на неразрывный
    strMail = objRS.Fields("mail").value
    strPhone = objRS.Fields("telephoneNumber").Value
        strPhone = Replace(strPhone,"-","") 'Откидываем тире из телефона
        strPhone = Replace(strPhone," ","") 'Откидываем пробелы из телефона
        a = Left(strPhone,1)
        if a = "8" then strPhone = "+7" + Mid(strPhone,2)
        if a = "(" then strPhone = "+7" + strPhone
    arrOtherPhone = objRS.Fields("otherTelephone").Value
    strMobile = objRS.Fields("mobile").Value
        strMobile = Replace(strMobile,"-","") 'Откидываем тире из телефона
        strMobile = Replace(strMobile," ","") 'Откидываем пробелы из телефона
        a = Left(strMobile,1)
        if a = "8" then strMobile = "+7" + Mid(strMobile,2)
        if a = "(" then strMobile = "+7" + strMobile
    strCountry = objRS.Fields("c").Value
    strCity = objRS.Fields("L").Value
    strCompany = objRS.Fields("company").Value
    strDepartment = objRS.Fields("department").Value
        strDepartment = Replace(strDepartment,"Отдел","") 'Откидываем из названия отдела слово отдел,
        strDepartment = Replace(strDepartment,"отдел","") 'так как оно присутствует в шапке таблицы!
        strDepartment = Trim(strDepartment) 'Убираем пробелы вначале и вконце названия отдела.
        strDepartment = uCase(Left(strDepartment,1)) + Mid(strDepartment,2) 'Делаем первую букву в названии заглавной, остальные оставляем без изменения.
    strTitle = objRS.Fields("title").Value
    strGivenName = objRS.Fields("givenname").Value
    strGivenName = Replace(strName," ",chr(160)) 'Заменяем обычный пробел на неразрывный
    strSN = objRS.Fields("sn").Value
    strDisplayName = objRS.Fields("displayname").Value
    a = Instr(strDisplayName,"-")
    if a=0 then    strDisplayName = Replace(strDisplayName," ",chr(160)) 'Заменяем обычный пробел на неразрывный, если в ФИО нет двойной фамилии
    If IsNull(arrOtherPhone) Then
        strOtherPhone = ""
    Else
        strOtherPhone = ""
        For Each strItem In arrOtherPhone
            If (strOtherPhone = "") Then
                strItem = Replace(strItem,"-","") 'Откидываем тире из телефона
                strItem = Replace(strItem," ","") 'Откидываем пробелы из телефона
                a = Left(strItem,1)
                if a = "8" then strItem = "+7" + Mid(strItem,2)
                if a = "(" then strItem = "+7" + strItem
                strOtherPhone = strItem
            Else
                strItem = Replace(strItem,"-","") 'Откидываем тире из телефона
                strItem = Replace(strItem," ","") 'Откидываем пробелы из телефона
                a = Left(strItem,1)
                if a = "8" then strItem = "+7" + Mid(strItem,2)
                if a = "(" then strItem = "+7" + strItem
                strOtherPhone = strOtherPhone & "<br>" & strItem
            End If
        Next
    End If
    If strOtherPhone <> "" then strPhone = strPhone & "<br>" & strOtherPhone
    'if strMail<>"" then 'Если смотреть по пользователям, то условие нужно, чтобы не прописывались системные учётки
        html = html + "<tr BGCOLOR=#CCCCCC>" + SE
        html = html + "<td>" & strDisplayName & "</td>" + SE
        html = html + "<td><a class=" & chr(34) & "link" & chr(34) & " href=" & chr(34) & "mailto:" & strMail & chr(34) & ">" & strMail & "</a></td>" + SE
        html = html + "<td>" & strPhone & "</td>" + SE
        'html = html + "<td>" & strOtherPhone & "</td>" + SE
        html = html + "<td>" & strMobile & "</td>" + SE
        if strCountry = "RU" then strCountry = "Россия"
        if strCountry = "UA" then strCountry = "Украина"
        if strCountry = "BY" then strCountry = "Беларусь"
        if strCountry = "KZ" then strCountry = "Казахстан"
        if strCountry = "GB" then strCountry = "Великобритания"
        if strCountry = "DE" then strCountry = "Германия"
        if strCountry = "IT" then strCountry = "Италия"
        if strCountry = "FR" then strCountry = "Франция"
        if strCountry = "BE" then strCountry = "Бельгия"
        if strCountry = "US" then strCountry = "США"
        html = html + "<td>" & strCountry & "</td>" + SE
        html = html + "<td>" & strCity & "</td>" + SE
        html = html + "<td>" & strCompany & "</td>" + SE
        html = html + "<td>" & strDepartment & "</td>" + SE
        html = html + "<td>" & strTitle & "</td>" + SE
        html = html + "</tr>" + SE
        strI = strI + 1
    'end if 'Если смотреть по пользователям, то условие нужно, чтобы не прописывались системные учётки
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
html = html + "Всего в телефонном справочнике обнаружено " + a + " запис"
select case b
    case "0", "5", "6", "7", "8", "9"
        html = html + "ей." + SE
    case "1"
        html = html + "ь." + SE
    case "2", "3", "4"
        html = html + "и." + SE
end select
html = html + "<BR>" + SE
html = html + "</p>" + SE
html = html + "</BODY>" + SE
html = html + "</HTML>" + SE
HFuter = HFuter + "<div class=" + chr(34) + "footer" + chr(34) + ">" + SE
HFuter = HFuter + "<p align=" + chr(34) + "center" + chr(34) +">" + SE
HFuter = HFuter + "Программа телефонного справочника была разработана сотрудниками отдела ИТ. С уважением, ИТ-отдел." + SE
HFuter = HFuter + "</p>" + SE
HFuter = HFuter + "</div>" + SE
Dim fso, tf
Set fso = CreateObject("Scripting.FileSystemObject")
Set tf = fso.CreateTextFile("\\FileServer\Contacts\Телефонный справочник сотрудников.htm", True)
tf.WriteLine(HHead)
tf.WriteLine(HCss)
tf.WriteLine(HJava)
tf.WriteLine(HTML)
tf.WriteLine(HFuter)
tf.Close
Set fso = Nothing