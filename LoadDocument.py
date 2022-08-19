import pandas as pd
import smtplib

class LoadDocument:

    """Загрузка документа"""
    __table = str
    __emailTeacher = str
    __path = str
    __emailStudent = str
    def __init__(self, table = str, emailTeacher = str, emailStudent = str, path = str):
        self.__table = table
        self.__emailTeacher = emailTeacher
        self.__path = path
        self.__emailStudent = emailStudent

    def __LoadTable (self, table = str, path = str):
        table = pd.read_excel(io=path, engine="openpyxl", sheet_name=table, skiprows=[0, 2])
        return table

    def __LoadEmail (self, email = str, path = str):
        email = pd.read_excel(io=path, engine="openpyxl", sheet_name=email)
        return email

    # Функция создания таблицы студентов c предметами которые они провалили и ср.балл
    def __AddStudent(self, table=pd.DataFrame, _atributes=pd.Index):
        blackList = pd.DataFrame([])
        for k in range(len(table[_atributes[0]])):
            listSubject = []
            for i in range(2, len(_atributes) - 1):
                if table[_atributes[i]][k] < 1:
                    listSubject.append(_atributes[i])
            for i in range((len(_atributes) - 3) - len(listSubject)):
                listSubject.append("0")
            blackList[table[_atributes[1]][k]] = listSubject
        return blackList


    # Создание таблицы с предметами
    def __AddSubject(self, table=pd.DataFrame, atributes=pd.Index):
        listOfSubject = pd.DataFrame([])
        for i in range(3, len(atributes) - 1):
            listOfStudent = []
            for j in range(len(table[atributes[0]])):
                if table[atributes[i]][j] < 1:
                    listOfStudent.append(table[atributes[1]][j])
            for j in range(len(table[atributes[1]]) - len(listOfStudent)):
                listOfStudent.append("0")
            listOfSubject[atributes[i]] = listOfStudent
        return listOfSubject

    # Функция удаления студентов у которых бал по всем предметам больше 1
    def __Delete(self, table=pd.DataFrame):
        listStudent = table.columns
        bool = False
        for i in range(len(listStudent)):
            for j in range(len(table[listStudent[i]])):
                if table[listStudent[i]][j] != "0":
                    bool = False
                    break
                else:
                    bool = True
            if bool:
                table.pop(listStudent[i])
        return table


    # Получение таблицы неуспивающик стуентов
    def GetTableStudent(self):
        table = self.__LoadTable(self.__table, self.__path)
        table = self.__AddStudent(table, table.columns)
        table = self.__Delete(table)
        return table


    def GetTableSubject(self):
        table = self.__LoadTable(self.__table, self.__path)
        table = self.__AddSubject(table, table.columns)
        table = self.__Delete(table)
        return table



    def __MessegeStudent(self, tableOfStudent=pd.DataFrame, Student=str):
        message = "Уважаемый " + Student + "\n"
        message = message + "Вы являетесь должником по 6 кн, "
        listOfSubject = []
        for i in range(len(tableOfStudent[Student])):
            if tableOfStudent[Student][i] != "0":
                if tableOfStudent[Student][i] != "Ср. балл":
                    listOfSubject.append(tableOfStudent[Student][i])
            else:
                break
        if len(listOfSubject) == 1:
            message += "по дисциплине " + listOfSubject[
                0] + ". Необходимо в срок до 8 недели сдать всё долги. Информацию о сдаче доводить до сведения деканата раз в 2-3 дня"
        else:
            message += "по дисциплинам "
            for i in range(len(listOfSubject)):
                if i == 0:
                    message += listOfSubject[i]
                else:
                    message += ", " + listOfSubject[i]
            message += ". Необходимо в срок до 8 недели сдать всё долги. Информацию о сдаче доводить до сведения деканата раз в 2-3 дня"
        return message

    # Фуекция отправки сообщения
    def __SendMessage(self, adress_email_sender=str, password=str, andess_email_recipient=str, message=str):
        smtpObj = smtplib.SMTP("smtp.gmail.com", 587)
        smtpObj.ehlo()
        smtpObj.starttls()
        smtpObj.login(adress_email_sender, password)
        smtpObj.sendmail(adress_email_sender, andess_email_recipient, message.encode("utf-8"))
        smtpObj.quit()

    # Поиск и отправка сообщений студентам
    def __SelectStudent(self, table=pd.DataFrame, table_email=pd.DataFrame):
        listStudent = table.columns
        listAtributes = table_email.columns
        # поиск студентов со средним баллов меньше 1
        for i in range(len(listStudent)):
            for j in range(len(table[listStudent[i]])):
                if table[listStudent[i]][j] == "Ср. балл":
                    # поиск в таблице с почтой
                    for k in range(len(table_email[listAtributes[1]])):
                        if table_email[listAtributes[1]][k] == listStudent[i]:
                            self.__SendMessage("aleks24129958@gmail.com", "rqvzpuxoigdkqnmy", table_email[listAtributes[2]][k],
                                        self.__MessegeStudent(table, listStudent[i]))

    def __MessegeOfTeacher(self, tableOfSubject=pd.DataFrame, subject=str):
        messege = "У вас "
        arraySubject = []
        for i in range(len(tableOfSubject[subject])):
            if tableOfSubject[subject][i] != "0":
                arraySubject.append(tableOfSubject[subject][i])
        if len(arraySubject) == 1:
            messege += "1 должник: " + arraySubject[0]
        else:
            if len(arraySubject) >= 2 and len(arraySubject) <= 4:
                messege += str(len(arraySubject)) + " должника: "
            else:
                messege += str(len(arraySubject)) + " должников: "
        for i in range(len(arraySubject)):
            if i == 0:
                messege += arraySubject[i]
            else:
                messege += ", " + arraySubject[i]
        messege += ". Необходимо обеспечить сдачу долгов до 8 неделе. Доводить информацию о сдаче долгов до деканата."
        return messege

    # Поиск и отправка сообщений преподователям
    def __SelectTeacher(self, table=pd.DataFrame, table_email=pd.DataFrame):
        listOfSubject = table.columns
        listOfEmail = table_email.columns
        for i in range(len(listOfSubject)):
            for j in range(len(table_email[listOfEmail[1]])):
                if listOfSubject[i] == table_email[listOfEmail[1]][j]:
                    self.__SendMessage("aleks24129958@gmail.com", "", table_email[listOfEmail[3]][j],
                                     self.__MessegeOfTeacher(table, listOfSubject[i]))


    def SendMessege(self):
        tableOfStudent = self.GetTableStudent()
        emailOfStudent = self.__LoadEmail(self.__emailStudent, self.__path)
        self.__SelectStudent(tableOfStudent, emailOfStudent)
        tableOfSubject = self.GetTableSubject()
        emailOfTeacher = self.__LoadEmail(self.__emailTeacher, self.__path)
        self.__SelectTeacher(tableOfSubject, emailOfTeacher)