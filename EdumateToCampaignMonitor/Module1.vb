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
		Console.WriteLine("v2019-10-29-2")
		addSubscriber(getEdumateSubscribers())
		'test()
	End Sub

	Sub test()
		Dim testlist As New List(Of emailSubscriber)
		Dim testuser As New emailSubscriber

		testuser.emailaddress = "mharding@ofgs.nsw.edu.au"
		testuser.name = "Matthew Harding"

		Dim objCustomField As New SubscriberCustomField
		objCustomField.Key = "YEAR GROUP"
		objCustomField.Value = "N/A"

		testuser.customFields.Add(objCustomField)

		testlist.Add(testuser)

		addSubscriber(testlist)




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
				objSubscriberList.Add(objEmailSubscriber.emailaddress, objEmailSubscriber.name, objEmailSubscriber.customFields, False, 1)
			Catch ex As Exception
                Console.WriteLine(ex)
            End Try

            i = i + 1
			Console.WriteLine(i & " of " & users.Count & " - " & objEmailSubscriber.emailaddress)
		Next

    End Sub

    Function getEdumateSubscribers()

        Dim commandString As String = "
SELECT DISTINCT  
firstname,
surname as lastname,
email_address,
listagg(DISTINCT CAST((form) AS varchar(10000)),'') WITHIN GROUP (ORDER BY form ASC) as year_group,
listagg(DISTINCT CAST((pa_class) AS varchar(10000)),'') WITHIN GROUP (ORDER BY pa_class ASC) as pa_class,
listagg(DISTINCT CAST((groups) AS varchar(10000)),'') WITHIN GROUP (ORDER BY groups ASC) as staff_group,
listagg(DISTINCT CAST((roll_class) AS varchar(10000)),'') WITHIN GROUP (ORDER BY roll_class ASC) as roll_class,
listagg(DISTINCT CAST((pa_group) AS varchar(10000)),'') WITHIN GROUP (ORDER BY pa_group ASC) as pa_group,
listagg(DISTINCT CAST((all_classes) AS varchar(10000)),'') WITHIN GROUP (ORDER BY all_classes ASC) as all_classes,
salutation,
debtor_title

FROM
	
	(
		-- Get emails for all current students
--		SELECT DISTINCT       
--			contact.firstname,
--			contact.surname,
--			contact.email_address,
--			salutation.salutation,
--			'' as debtor_title,
--			form.short_name || ';' as form,
--			class.class || ';' as pa_group, 
--			course.code || '_' || class.identifier || ';' as pa_class, 
--			'' AS groups,
--			VSCE.class || ';' as roll_class,	
--			allCourse.code || '_' || allclass.identifier || ';' as all_classes
--				
--		FROM
--			edumate.VIEW_STUDENT_START_EXIT_DATES VSSED
--			
--		INNER JOIN student on VSSED.student_id = student.student_id 
--		
--		INNER JOIN contact on student.contact_id = contact.contact_id	
--		
--		INNER JOIN salutation on contact.salutation_id = salutation.salutation_id
--/*
--		INNER JOIN edumate.view_enroled_students_form_run sfr
--		ON VSSED.student_id = sfr.STUDENT_ID 
--		AND Date(CURRENT DATE)  between sfr.computed_v_start_date and sfr.computed_end_date
--*/
--
--		INNER JOIN student_form_run sfr
--		ON VSSED.student_id = sfr.STUDENT_ID 
--		AND sfr.form_run_id IN (SELECT form_run_id
--								FROM edumate.VIEW_FORM_RUN_DATES
--								WHERE
--								Date(CURRENT DATE) between v_start_date and end_date)
--		
--		INNER JOIN form_run 
--		ON sfr.form_run_id = form_run.form_run_id
--
--		INNER JOIN form 
--		ON form_run.form_id = form.form_id
--
--		-- Get roll call class
--		LEFT JOIN edumate.view_student_class_enrolment VSCE on (VSSED.student_id = VSCE.student_id   
--		AND	VSCE.class_type_id = 2
--		AND Date(CURRENT DATE)  between VSCE.start_date and VSCE.end_date)
--		
--		
--		
--		-- Get All classes
--		LEFT JOIN edumate.view_student_class_enrolment VSFCE on (VSSED.student_id = VSFCE.student_id   
--		AND Date(CURRENT DATE)  between (VSFCE.start_date - 10 DAYS) and VSFCE.end_date)
--		
--		LEFT JOIN edumate.course allCourse ON
--		vsfce.course_id = allCourse.course_id 
--		
--		LEFT JOIN edumate.class allClass ON
--		vsfce.class_id = allClass.class_id
--		
--		
--		-- Get performing arts classes
--		LEFT JOIN edumate.view_student_class_enrolment VSPA on (VSSED.student_id = VSPA.student_id
--		AND(VSPA.course like 'PA %' OR VSPA.course like '%Keyboard Club%') 
--		AND Date(CURRENT DATE)  between (VSPA.start_date - 10 days)  and VSPA.end_date)
--		
--		LEFT JOIN course on VSPA.course_id = course.course_id
--		LEFT JOIN class on VSPA.class_id = class.class_id
--
--		WHERE 
--			Date(CURRENT DATE)  between (VSSED.start_date - 60 days) and VSSED.exit_date 
--		
--		UNION
		
		
		
		-- Get emails for carers of current students
		SELECT DISTINCT
		parent.firstname,
		parent.surname,
		parent.email_address,
		salutation.salutation,
		debtor.title as debtor_title,
		form.short_name || ';' as form,
		class.class || ';' as pa_group, 
		course.code || '_' || class.identifier || ';' as pa_class, 
		'' AS groups,
		VSCE.class || ';' as roll_class,
		allCourse.code || '_' || allclass.identifier || ';' as all_classes
		
		FROM
			edumate.VIEW_STUDENT_START_EXIT_DATES VSSED
			
		INNER JOIN edumate.VIEW_STUDENT_MAIL_CARERS VSMC on VSSED.student_id = VSMC.student_id
		
		INNER JOIN CONTACT parent on 
		(	VSMC.CARER1_CONTACT_ID = parent.contact_id OR
			VSMC.CARER2_CONTACT_ID = parent.contact_id OR
			VSMC.CARER3_CONTACT_ID = parent.contact_id OR
			VSMC.CARER4_CONTACT_ID = parent.contact_id )
	
	
		INNER JOIN salutation on parent.salutation_id = salutation.salutation_id
	
		LEFT JOIN debtor_contact dc on parent.contact_id = dc.contact_id
		LEFT JOIN debtor on dc.debtor_id = debtor.debtor_id
		
/*
		INNER JOIN edumate.view_enroled_students_form_run sfr
		ON VSSED.student_id = sfr.STUDENT_ID 
		AND Date(CURRENT DATE)  between sfr.computed_v_start_date and sfr.computed_end_date
*/
		INNER JOIN student_form_run sfr
		ON VSSED.student_id = sfr.STUDENT_ID 
		AND sfr.form_run_id IN (SELECT form_run_id
								FROM edumate.VIEW_FORM_RUN_DATES
								WHERE
								Date(CURRENT DATE) between v_start_date and end_date)

		INNER JOIN form_run 
		ON sfr.form_run_id = form_run.form_run_id

		INNER JOIN form 
		ON form_run.form_id = form.form_id

		-- Get roll call class
		LEFT JOIN edumate.view_student_class_enrolment VSCE on (VSSED.student_id = VSCE.student_id   
		AND	VSCE.class_type_id = 2
		AND Date(CURRENT DATE)  between VSCE.start_date and VSCE.end_date)
									
		-- Get all classes 
		LEFT JOIN edumate.view_student_class_enrolment VSFCE on (VSSED.student_id = VSFCE.student_id   
		AND Date(CURRENT DATE)  between (VSFCE.start_date - 10 DAYS) and VSFCE.end_date)
		
		LEFT JOIN edumate.course allCourse ON
		vsfce.course_id = allCourse.course_id 
		
		LEFT JOIN edumate.class allClass ON
		vsfce.class_id = allClass.class_id
				
		-- Get performing arts classes
		LEFT JOIN edumate.view_student_class_enrolment VSPA on (VSSED.student_id = VSPA.student_id
		AND(VSPA.course like 'PA %' OR VSPA.course like '%Keyboard Club%') 
		AND Date(CURRENT DATE)  between (VSPA.start_date - 10 days) and VSPA.end_date)
		
		LEFT JOIN course on VSPA.course_id = course.course_id
		LEFT JOIN class on VSPA.class_id = class.class_id

		WHERE 
			Date(CURRENT DATE)  between (VSSED.start_date - 60 days) and VSSED.exit_date 

		UNION
		

		-- Get details of current staff
		SELECT DISTINCT     
			contact.firstname,
			contact.surname,
			contact.email_address,
			salutation.salutation,
			debtor.title as debtor_title,
			'Staff;' as form,
			'' as pa_group, 
			'' as pa_class,
			groups || ';' as groups,
			class.class || ';' as roll_class,	
			'' AS all_classes
			
		FROM 
			STAFF

			INNER JOIN contact ON staff.contact_id = contact.contact_id 
			INNER JOIN salutation on contact.salutation_id = salutation.salutation_id
	
			LEFT JOIN debtor_contact dc on contact.contact_id = dc.contact_id
			LEFT JOIN debtor on dc.debtor_id = debtor.debtor_id
			
			INNER JOIN staff_employment ON staff.staff_id = staff_employment.staff_id
			
			LEFT JOIN teacher on contact.contact_id = teacher.contact_id
			LEFT JOIN class_teacher on (teacher.teacher_id = class_teacher.teacher_id
			                            AND class_teacher.class_id IN (SELECT class_id
																		FROM  edumate.view_student_class_enrolment 
																		WHERE class_type_id = 2 AND
																		Date(CURRENT DATE)  between start_date and end_date))
			LEFT JOIN class on class_teacher.class_id = class.class_id
			
			LEFT JOIN group_membership ON group_membership.contact_id = contact.contact_id
			LEFT JOIN groups ON group_membership.groups_id = groups.groups_id
			LEFT JOIN work_detail on contact.contact_id = work_detail.contact_id
			LEFT JOIN work_type on work_detail.work_type_id = work_type.work_type_id
			
		WHERE  
						(staff_employment.end_date is null or staff_employment.end_date >= Date(CURRENT DATE) )
			AND staff_employment.start_date <= (Date(CURRENT DATE) +2 DAYS)	AND
			(work_type.work_type <> 'COMPUTER' or work_type.work_type is null)		
	
			
			
			UNION
			
			SELECT        
contact.firstname, 
contact.surname AS lastname, 
contact.EMAIL_ADDRESS,
salutation.salutation,
' ' AS debtor_title,
'HSC;' AS year_group,
' ' AS PA_GROUP,
' ' AS PA_CLASS,
' ' AS STAFF_GROUP,
' ' AS ROLL_CLASS,
' ' AS ALL_CLASSES



FROM            
STUDENT
INNER JOIN contact ON student.contact_id = contact.contact_id
INNER JOIN edumate.view_student_start_exit_dates ON student.student_id = edumate.view_student_start_exit_dates.student_id
INNER JOIN student_form_run ON student_form_run.student_id = student.student_id
INNER JOIN form_run ON student_form_run.form_run_id = form_run.form_run_id
INNER JOIN form ON form_run.form_id = form.form_id
INNER JOIN stu_school ON student.student_id = stu_school.student_id
LEFT JOIN class_enrollment ON student.STUDENT_ID = class_enrollment.STUDENT_ID
LEFT JOIN class ON class_enrollment.class_id = class.class_id 
INNER JOIN TIMETABLE ON form_run.TIMETABLE_ID = timetable.TIMETABLE_ID
LEFT JOIN salutation on contact.salutation_id = salutation.salutation_id
LEFT JOIN 
	(
	SELECT        
	student.student_id, 
	edumate.view_student_class_enrolment.class

	FROM            
	STUDENT

	INNER JOIN edumate.view_student_class_enrolment ON student.student_id = edumate.view_student_class_enrolment.student_id

	WHERE 
 	(edumate.view_student_class_enrolment.class_type_id = 2)
	AND (edumate.view_student_class_enrolment.academic_year = char(year(current timestamp)))
	) RollClass ON rollclass.student_id = student.student_id
	
	
WHERE 

(YEAR(edumate.view_student_start_exit_dates.exit_date) = YEAR(student_form_run.end_date)) 

AND YEAR(edumate.view_student_start_exit_dates.exit_date) = year(current_date)

AND edumate.view_student_start_exit_dates.exit_date = timetable.COMPUTED_END_DATE

AND form.SHORT_NAME = '12'

AND class.CLASS LIKE '12%'


AND student.student_id NOT IN
(
	SELECT distinct       

	student.student_id

	FROM            
	STUDENT
	INNER JOIN edumate.view_student_start_exit_dates ON student.student_id = edumate.view_student_start_exit_dates.student_id
	INNER JOIN student_form_run ON student_form_run.student_id = student.student_id
	INNER JOIN form_run ON student_form_run.form_run_id = form_run.form_run_id
	INNER JOIN form ON form_run.form_id = form.form_id
	INNER JOIN stu_school ON student.student_id = stu_school.student_id
	
	WHERE 

	(YEAR(edumate.view_student_start_exit_dates.exit_date) = YEAR(student_form_run.end_date)) 

	AND (SELECT (current date) FROM sysibm.sysdummy1) BETWEEN (edumate.view_student_start_exit_dates.start_date - 90 days) AND edumate.view_student_start_exit_dates.exit_date
)

GROUP BY 

contact.firstname, 
contact.surname, 
edumate.view_student_start_exit_dates.start_date, 
edumate.view_student_start_exit_dates.exit_date, 
student.student_id, 
form.short_name,
YEAR(student_form_run.end_date),
student.student_number,
contact.birthdate,
stu_school.library_card,
Rollclass.class,
contact.EMAIL_ADDRESS,
salutation.salutation


			
			
			
			
		)
	
WHERE
email_address IS NOT NULL
					
GROUP BY 
surname,
firstname,
email_address,
salutation,
debtor_title

ORDER BY
surname,
firstname,
email_address,
salutation,
debtor_title



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

					objCustomField = New SubscriberCustomField
					objCustomField.Key = "ALL CLASSES"
					If Not dr.IsDBNull(8) Then
						objCustomField.Value = dr.GetValue(8)
					Else
						objCustomField.Value = "N/A"
					End If
					users.Last.customFields.Add(objCustomField)
					objCustomField = Nothing

					objCustomField = New SubscriberCustomField
					objCustomField.Key = "SALUTATION"
					If Not dr.IsDBNull(9) Then
						objCustomField.Value = dr.GetValue(9)
					Else
						objCustomField.Value = ""
					End If
					users.Last.customFields.Add(objCustomField)
					objCustomField = Nothing

					objCustomField = New SubscriberCustomField
					objCustomField.Key = "DEBTOR_TITLE"
					If Not dr.IsDBNull(10) Then
						objCustomField.Value = Replace(dr.GetValue(10), "&amp;", "&")
					Else
						objCustomField.Value = ""
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


