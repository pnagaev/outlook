' При создании нового письма макрос берёт из поля to ФИО в формате "Фамилия Имя Отчество" и добавляет первую строку в письме "Уважаемый Имя Отчество"
' Написано ChatGPT
Sub InsertGreeting()
    Dim objMail As Outlook.MailItem
    Dim objInspector As Outlook.inspector
    Dim objDoc As Object ' Word.Document
    Dim objSel As Object ' Word.Selection
    
    Dim recipientFullName As String
    Dim nameParts() As String
    Dim cleanName As String
    Dim greeting As String

    Set objInspector = Application.ActiveInspector
    If objInspector Is Nothing Then Exit Sub
    
    ' Проверяем, что открыто именно письмо
    If TypeOf objInspector.CurrentItem Is Outlook.MailItem Then
        Set objMail = objInspector.CurrentItem
        
        ' 1. Получаем имя и формируем приветствие
        If objMail.Recipients.Count > 0 Then
            recipientFullName = Trim(objMail.Recipients(1).Name)
            nameParts = Split(recipientFullName, " ")
            
            ' Логика Фамилия(0) Имя(1) Отчество(2)
            If UBound(nameParts) >= 2 Then
                cleanName = nameParts(1) & " " & nameParts(2)
                ' Пол по окончанию отчества (второе слово в cleanName)
                If Right(LCase(nameParts(2)), 1) = "а" Then
                    greeting = "Уважаемая " & cleanName & ","
                Else
                    greeting = "Уважаемый " & cleanName & ","
                End If
            ElseIf UBound(nameParts) = 1 Then
                cleanName = nameParts(1)
                If Right(LCase(cleanName), 1) = "а" Or Right(LCase(cleanName), 1) = "я" Then
                    greeting = "Уважаемая " & cleanName & ","
                Else
                    greeting = "Уважаемый " & cleanName & ","
                End If
            Else
                greeting = "Добрый день,"
            End If

            ' 2. Работаем с текстом через WordEditor (для управления курсором)
            Set objDoc = objInspector.WordEditor
            Set objSel = objDoc.Windows(1).Selection
            
            ' Переходим в самое начало письма
            objSel.HomeKey Unit:=6 ' 6 = wdStory (начало документа)
            
            ' Вставляем приветствие, два переноса строки и возвращаемся
            ' Текст будет выглядеть так:
            ' Уважаемый Иван Иванович,
            ' [Курсор здесь]
            ' (старый текст письма)
            
            objSel.TypeText Text:=greeting & vbCrLf & vbCrLf
            
            ' Опционально: если нужно оставить курсор на пустой строке под приветствием:
            ' Мы уже там после TypeText, так как добавили два vbCrLf.
            
        Else
            MsgBox "Поле 'Кому' пустое!", vbExclamation
        End If
    End If
End Sub
