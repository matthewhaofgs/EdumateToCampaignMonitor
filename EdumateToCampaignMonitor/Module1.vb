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

        Dim commandString As String = "
SELECT DISTINCT  
firstname,
surname as lastname,
email_address,
listagg(DISTINCT form, ',') WITHIN GROUP (ORDER BY form ASC) as year_group,
listagg(DISTINCT pa_class, ',') WITHIN GROUP (ORDER BY pa_class ASC) as pa_class,
listagg(DISTINCT groups, ',') WITHIN GROUP (ORDER BY groups ASC) as staff_group,
listagg(DISTINCT roll_class, ',') WITHIN GROUP (ORDER BY roll_class ASC) as roll_class

FROM
	
	(
		-- Get emails for carers of current students
		SELECT DISTINCT
		parent.firstname,
		parent.surname,
		parent.email_address,
		form.short_name	as form,
		course.code || '_' || class.identifier || ';' as pa_class, 
		NULL AS groups,
		VSCE.class as roll_class
		
		FROM
			edumate.view_student_class_enrolment VSCE
			
		INNER JOIN edumate.view_STUDENT_MAIL_CARERS VSMC on VSCE.student_id = VSMC.student_id
		INNER JOIN CONTACT parent on 
		(	VSMC.CARER1_CONTACT_ID = parent.contact_id OR
			VSMC.CARER2_CONTACT_ID = parent.contact_id OR
			VSMC.CARER3_CONTACT_ID = parent.contact_id OR
			VSMC.CARER4_CONTACT_ID = parent.contact_id )

		INNER JOIN student_form_run 
		ON VSCE.student_id = student_form_run.STUDENT_ID

		INNER JOIN form_run 
		ON student_form_run.FORM_RUN_ID = form_run.FORM_RUN_ID

		INNER JOIN form 
		ON form_run.FORM_ID = form.form_id

		LEFT JOIN edumate.view_student_class_enrolment VSPA on (VSCE.student_id = VSPA.student_id   AND
									(VSPA.course like 'PA %' OR VSPA.course like '%Keyboard Club%') 
									AND current_date between VSPA.start_date and VSPA.end_date)
		LEFT JOIN course on VSPA.course_id = course.course_id
		LEFT JOIN class on VSPA.class_id = class.class_id


		WHERE 
			VSCE.class_type_id = 2 AND current_date between VSCE.start_date and VSCE.end_date AND
			current_date between student_form_run.start_date and student_form_run.end_date	

		UNION
		
		-- Get emails for all current students
		SELECT        
			contact.firstname,
			contact.surname,
			contact.email_address,
			form.short_name	as form,
			course.code || '_' || class.identifier || ';' as pa_class, 
			NULL AS groups,
			VSCE.class as roll_class	
				
		FROM
			edumate.view_student_class_enrolment VSCE

		INNER JOIN student on VSCE.student_id = student.student_id
		
		INNER JOIN CONTACT on student.contact_id = contact.contact_id

		INNER JOIN student_form_run 
		ON VSCE.student_id = student_form_run.STUDENT_ID

		INNER JOIN form_run 
		ON student_form_run.FORM_RUN_ID = form_run.FORM_RUN_ID

		INNER JOIN form 
		ON form_run.FORM_ID = form.form_id
														
		LEFT JOIN edumate.view_student_class_enrolment VSPA on (VSCE.student_id = VSPA.student_id   AND
														
									(VSPA.course like 'PA %' OR VSPA.course like '%Keyboard Club%')  AND current_date between VSPA.start_date and VSPA.end_date)
		LEFT JOIN course on VSPA.course_id = course.course_id
		LEFT JOIN class on VSPA.class_id = class.class_id
		
		WHERE 
			VSCE.class_type_id = 2 AND current_date between VSCE.start_date and VSCE.end_date AND
			current_date between student_form_run.start_date and student_form_run.end_date	
		
		UNION
		
		-- Get details of current staff
		SELECT        
			contact.firstname,
			contact.surname,
			contact.email_address,
			'Staff' as form,
			null as pa_class,
			groups,
			class.class as roll_class	
			
		FROM 
			STAFF

			INNER JOIN Contact ON staff.contact_id = contact.contact_id 
			INNER JOIN staff_employment ON staff.staff_id = staff_employment.staff_id
			
			LEFT JOIN teacher on contact.contact_id = teacher.contact_id
			LEFT JOIN class_teacher on (teacher.teacher_id = class_teacher.teacher_id
			                            AND class_teacher.class_id IN (SELECT class_id
																		FROM  edumate.view_student_class_enrolment 
																		WHERE class_type_id = 2 AND
																			current_date between start_date and end_date))
			LEFT JOIN class on class_teacher.class_id = class.class_id
			
			LEFT JOIN group_membership ON group_membership.contact_id = contact.contact_id
			LEFT JOIN groups ON group_membership.groups_id = groups.groups_id
			LEFT JOIN work_detail on contact.contact_id = work_detail.contact_id
			LEFT JOIN work_type on work_detail.work_type_id = work_type.work_type_id
			
		WHERE  
			staff_employment.employment_type_id IN (1,2) AND
			(staff_employment.end_date is null or staff_employment.end_date >= current date)
			AND staff_employment.start_date <= (current date +2 DAYS)	AND
			(work_type.work_type <> 'COMPUTER' or work_type.work_type is null)		
	)
	
					
GROUP BY 
surname,
firstname,
email_address

ORDER BY
surname,
firstname,
email_address




"


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


                    Dim objCustomField As New SubscriberCustomField
                    objCustomField.Key = "YEAR GROUP"
                    If Not dr.IsDBNull(3) Then
                        objCustomField.Value = dr.GetValue(3)
                    Else
                        objCustomField.Value = "N/A"
                    End If
                    users.Last.customFields.Add(objCustomField)
                    objCustomField = Nothing

                    objCustomField = New SubscriberCustomField
                    objCustomField.Key = "PA CLASS"
                    If Not dr.IsDBNull(4) Then
                        objCustomField.Value = dr.GetValue(4)
                    Else
                        objCustomField.Value = "N/A"
                    End If
                    users.Last.customFields.Add(objCustomField)
                    objCustomField = Nothing

                    objCustomField = New SubscriberCustomField
                    objCustomField.Key = "STAFF GROUPS"
                    If Not dr.IsDBNull(5) Then
                        objCustomField.Value = dr.GetValue(5)
                    Else
                        objCustomField.Value = "N/A"
                    End If
                    users.Last.customFields.Add(objCustomField)
                    objCustomField = Nothing

                    objCustomField = New SubscriberCustomField
                    objCustomField.Key = "ROLL CLASS"
                    If Not dr.IsDBNull(6) Then
                        objCustomField.Value = dr.GetValue(6)
                    Else
                        objCustomField.Value = "N/A"
                    End If
                    users.Last.customFields.Add(objCustomField)
                    objCustomField = Nothing

                End If


            End While
            conn.Close()
        End Using

        getEdumateSubscribers = users


    End Function

End Module


