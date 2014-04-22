'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Создаем список контактов из активных пользователей Active directory '
'          Системные учётные записи не трогаем по критерию:           '
'              у системной учётки нет электронной почты!              '
'   Скрипт написал Анчуров Олег Владимирович в 2013 году. Версия 3.   '
'              Скрипт взят с сайта: http://ithelp.moy.su              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
On error resume next ' отключаем ошибки

' Объявляем переменные
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

' Итак, поехали
' Определяем имя домена
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection
strBase = "<LDAP://" & strDNSDomain & ">"

' Найти все активные учетные записи.
' Коды UseraccountControl тут: http://support.microsoft.com/kb/305144
' Все коды аттрибутов тут: http://www.computerperformance.co.uk/Logon/LDAP_attributes_active_directory.htm
strFilter = "(&(objectCategory=person)(objectClass=user)(|(useraccountControl=66048)(useraccountcontrol=512)))" 'useraccountcontrol=512 - включенные пользователи!
' В переменной strAttributes обязательно перечисляем список атрибутов, которые переносятся, иначе работать не будет!
strAttributes = "name,mail,telephoneNumber,otherTelephone,mobile,c,L,company,department,title,givenname,sn,displayname"

' Формируем строку запроса.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

' Выполним запрос.
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 500 ' Максимальное количество пользователей, которое получит в ответе запрос к AD!
adoCommand.Properties("Timeout") = 307
adoCommand.Properties("Cache Results") = False
Set objRS = adoCommand.Execute

' Обработаем полученные данные из запроса
While not objRS.EOF
    strName = objRS.Fields("name").Value
    strMail = objRS.Fields("mail").value
    strPhone = objRS.Fields("telephoneNumber").Value
        strPhone = Replace(strPhone,"-","") 'Откидываем тире из телефона
        strPhone = Replace(strPhone," ","") 'Откидываем пробелы из телефона   
    arrOtherPhone = objRS.Fields("otherTelephone").Value
    strMobile = objRS.Fields("mobile").Value
        strMobile = Replace(strMobile,"-","") 'Откидываем тире из телефона
        strMobile = Replace(strMobile," ","") 'Откидываем пробелы из телефона
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
                strItem = Replace(strItem,"-","") 'Откидываем тире из телефона
                strItem = Replace(strItem," ","") 'Откидываем пробелы из телефона   
                strOtherPhone = strOtherPhone & " " & strItem
            End If
        Next
    End If
   
    'Заполним поля, если есть электронная почта
    if strMail<>"" then
        Err.Clear
        ' Создаём контакт из пользователя. Подготовка.
        strContainer = "cn=users" 'Название контейнера, в который будем ложить контакты, созданные из пользователей.
        strContactName = "cn=" & strName 'Название контакта
        strEmail = strMail 'Адрес электронной почты

        ' Создаём запрос к Active Directory
        Set objRoot = GetObject("LDAP://rootDSE")
        strDNS = objRoot.Get("defaultNamingContext")
        Set objDomain = GetObject("LDAP://" & strDNS)

        ' Создаём контакт из пользователя.
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
        objContact.SetInfo 'контакт успешно создан! (в случае, если такой уже окажется, то вызовется ошибка, которая проигнорируется скриптом!)
        If Err.Number = 0 then
            UNK = UNK + 1
            UserN = UserN + strName + chr(13) + chr(10)
        end if
    end if

    objRS.MoveNext
Wend

' А теперь удаляем из AD контакты заблокированных пользователей
strFilter = "(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))"' - выключенные пользователи!
' В переменной strAttributes оязательно перечисляем список атрибутов, которые переносятся, иначе работать не будет!
strAttributes = "name,mail"' в этот раз нас не интересуют остальные параметры

' Формируем строку запроса.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

' Выполним запрос.
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
        'Удаляем контакт заблокированного пользователя
        strContainer = "cn=users" 'Название контейнера, в котором будем удалять контакты заблокированных пользователей.
        strContactName = "cn=" & chr(34) & strName & chr(34) 'Название контакта
        Set objOU = GetObject("LDAP://"& strContainer & "," & strDNS)
        a = objOU.Delete ("Contact", strContactName)
        If Err.Number = 500 then
            UDK = UDK + 1
            UserD = UserD + strName + chr(13) + chr(10)
        end if
    end if
    objRS.MoveNext
Wend

'Если были произведены автоматические манипуляции с учётками, то надо об этом сообщить админу по электронной почте

If (UNK + UDK) > 0 then
    Dim objEmail, MSG
    MSG = "ПРОТОКОЛ РАБОТЫ СКРИПТА, АВТОМАТИЧЕСКИ СОЗДАЮЩЕГО КОНТАКТЫ ИЗ ПОЛЬЗОВАТЕЛЕЙ В AD:" + chr(13) + chr(10) + chr(13) + chr(10)
    MSG = MSG + "Протокол создания контактов:" + chr(13) + chr(10)
    MSG = MSG + UserN
    MSG = MSG + "Всего создано контактов: " + CStr(UNK) + chr(13) + chr(10) + chr(13) + chr(10)
    MSG = MSG + "Протокол удаления контактов:" + chr(13) + chr(10)
    MSG = MSG + UserD
    MSG = MSG + "Всего удалено контактов: " + CStr(UDK) + chr(13) + chr(10)+ chr(13) + chr(10)
    MSG = MSG + "Время/дата отработки скрипта: " + CStr(Time) + "/" + CStr(Date) + chr(13) + chr(10)
   
    Const EmailFrom = "bot@firma.ru"         ' от кого будет отправляться e-mail
    Const EmailPassword = "SuperPassword"       ' пароль от e-mail
    Const strSmtpServer = "smtp.firma.ru" ' smtp сервер
    Const EmailTo = "admin@firma.ru"    ' Кому будет отправляться e-mail
    Set objEmail = CreateObject("CDO.Message")
   
    objEmail.From = EmailFrom
    objEmail.To = EmailTo
    objEmail.Subject = "Отчёт по работе скрипта MailFromUsers" ' Тема письма
    objEmail.Textbody = MSG
    objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = EmailFrom
    objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = EmailPassword
    objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSmtpServer
    objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    objEmail.Configuration.Fields.Update
    objEmail.Send
    Set objEmail = Nothing ' Чистим память.
End If

' Ну вот и всё! Чистим память и выходим из скрипта!
Set objRS = Nothing
Set adoCommand = Nothing
Set adoConnection = Nothing
Set objOU = Nothing
WScript.Quit
