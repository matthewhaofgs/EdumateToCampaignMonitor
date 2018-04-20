Imports createsend_dotnet
Imports System.IO
Imports IBM.Data.DB2

Class emailSubscriber
    Public emailaddress As String
    Public name As String
    Public customFields As New List(Of SubscriberCustomField)
End Class



Module EdumateToCampaignMonitor


    Sub Main()
        addSubscriber(getEdumateSubscribers())
    End Sub


    Function readTextFile(filename As String)
        Dim directory As String = My.Application.Info.DirectoryPath

        Dim strKey

        Using sr As New StreamReader(directory & "\" & filename)
            Dim line As String
            While Not sr.EndOfStream
                line = sr.ReadLine
                If Not IsNothing(line) Then
                    strKey = line
                End If
            End While
        End Using

        readTextFile = strKey

    End Function


    Sub addSubscriber(users As List(Of emailSubscriber))
        Dim objAuth = New ApiKeyAuthenticationDetails(readTextFile("api.txt"))
        Dim strListID As String = readTextFile("listAPI.txt")
        Dim objSubscriberList As New Subscriber(objAuth, strListID)



        Dim i As Integer = 0
        For Each objEmailSubscriber In users

            Try
                objSubscriberList.Add(objEmailSubscriber.emailaddress, objEmailSubscriber.name, objEmailSubscriber.customFields, False)
            Catch ex As Exception
                Console.WriteLine(ex)
            End Try

            i = i + 1
            Console.WriteLine(i & " of " & users.Count)
        Next

    End Sub

    Function getEdumateSubscribers()

        Dim commandString As String = "SELECT DISTINCT  
firstname,
surname as lastname,
email_address,
listagg(short_name, ',') WITHIN GROUP (ORDER BY short_name ASC) as year_group,
listagg(pa_class, ',') WITHIN GROUP (ORDER BY pa_class ASC) as pa_class

	FROM(
		SELECT        
		parentcontact.firstname,
		parentcontact.surname,
		parentcontact.email_address,
		form.SHORT_NAME,
		course.code || '_' || class.identifier || ';' as pa_class 

		FROM            relationship

		INNER JOIN contact as ParentContact
		ON relationship.contact_id2 = Parentcontact.contact_id

		INNER JOIN contact as StudentContact 
		ON relationship.contact_id1 = studentContact.contact_id

		INNER JOIN student
		ON studentContact.contact_id = student.contact_id

		INNER JOIN carer 
		ON parentcontact.contact_id = carer.contact_id

		INNER JOIN student_form_run 
		ON student.student_id = student_form_run.STUDENT_ID

		INNER JOIN form_run 
		ON student_form_run.FORM_RUN_ID = form_run.FORM_RUN_ID

		INNER JOIN form 
		ON form_run.FORM_ID = form.form_id
		
		LEFT JOIN edumate.view_student_class_enrolment VSCE on (student.student_id = VSCE.student_id   AND
														VSCE.course like 'PA %' AND current_date between VSCE.start_date and VSCE.end_date)
		LEFT JOIN course on VSCE.course_id = course.course_id
		LEFT JOIN class on VSCE.class_id = class.class_id

		
		WHERE        (relationship.relationship_type_id IN (1, 4, 8, 15, 28, 33)) 
		AND current_date between student_form_run.start_date and student_form_run.end_date
		


		UNION

		SELECT        
		parentcontact.firstname,
		parentcontact.surname,
		parentcontact.email_address,
		form.SHORT_NAME,
		course.code || '_' || class.identifier || ';' as pa_class 

		FROM            relationship

		INNER JOIN contact as ParentContact
		ON relationship.contact_id1 = Parentcontact.contact_id

		INNER JOIN contact as StudentContact 
		ON relationship.contact_id2 = studentContact.contact_id

		INNER JOIN student
		ON studentContact.contact_id = student.contact_id

		INNER JOIN carer 
		ON parentcontact.contact_id = carer.contact_id

		INNER JOIN student_form_run 
		ON student.student_id = student_form_run.STUDENT_ID

		INNER JOIN form_run 
		ON student_form_run.FORM_RUN_ID = form_run.FORM_RUN_ID

		INNER JOIN form 
		ON form_run.FORM_ID = form.form_id

		LEFT JOIN edumate.view_student_class_enrolment VSCE on student.student_id = VSCE.student_id 
		LEFT JOIN course on VSCE.course_id = course.course_id
		LEFT JOIN class on VSCE.class_id = class.class_id

		WHERE        (relationship.relationship_type_id IN (2, 5, 9, 16, 29, 34)) 
		AND current_date between student_form_run.start_date and student_form_run.end_date
		AND VSCE.course like 'PA %' AND current_date between VSCE.start_date and VSCE.end_date
		
UNION

SELECT        
contact.firstname,
contact.surname,
contact.email_address,
'Staff' AS staff,
' ' as pa_class

FROM            STAFF

INNER JOIN Contact 
  ON staff.contact_id = contact.contact_id 
INNER JOIN staff_employment
  ON staff.staff_id = staff_employment.staff_id
LEFT JOIN sys_user 
  ON contact.contact_id = sys_user.contact_id
WHERE  (staff_employment.end_date is null or staff_employment.end_date >= current date)
and staff_employment.start_date <= (current date +90 DAYS)

)

-- ORDER BY email_address, surname, firstname	
GROUP BY email_address, surname, firstname"


        Dim users As New List(Of emailSubscriber)

        Using conn As New IBM.Data.DB2.DB2Connection(readTextFile("edumate.txt"))
            conn.Open()

            'define the command object to execute
            Dim command As New IBM.Data.DB2.DB2Command(commandString, conn)
            command.Connection = conn
            command.CommandText = commandString

            Dim dr As IBM.Data.DB2.DB2DataReader
            dr = command.ExecuteReader


            While dr.Read()
                If Not dr.IsDBNull(2) Then
                    users.Add(New emailSubscriber)
                    users.Last.emailaddress = dr.GetValue(2)
                    If Not dr.IsDBNull(0) Then users.Last.name = dr.GetValue(0) & " "
                    If Not dr.IsDBNull(1) Then users.Last.name = users.Last.name & dr.GetValue(1)

                    If Not dr.IsDBNull(3) Then
                        Dim objCustomField As New SubscriberCustomField
                        objCustomField.Key = "YEAR GROUP"
                        objCustomField.Value = dr.GetValue(3)
                        users.Last.customFields.Add(objCustomField)
                        objCustomField = Nothing
                    End If

                    If Not dr.IsDBNull(4) Then
                        Dim objCustomField As New SubscriberCustomField
                        objCustomField.Key = "PA CLASS"
                        objCustomField.Value = dr.GetValue(4)
                        users.Last.customFields.Add(objCustomField)
                        objCustomField = Nothing
                    End If

                End If


            End While
            conn.Close()
        End Using

        getEdumateSubscribers = users


    End Function

End Module


