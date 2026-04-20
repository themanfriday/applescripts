-- Creates a calendar appointment in Microsoft Outlook
-- Customize the properties below before running

set appointmentSubject to "Meeting"
set appointmentLocation to "Conference Room A"
set appointmentStart to current date
set appointmentEnd to appointmentStart + (1 * hours)
set appointmentNotes to "Agenda to be shared."
set isAllDay to false
set reminderMinutes to 15

tell application "Microsoft Outlook"
	set newAppointment to make new calendar event with properties {¬
		subject: appointmentSubject, ¬
		location: appointmentLocation, ¬
		start time: appointmentStart, ¬
		end time: appointmentEnd, ¬
		content: appointmentNotes, ¬
		all day flag: isAllDay, ¬
		reminder time: reminderMinutes}

	save newAppointment

	display dialog "Appointment \"" & appointmentSubject & "\" created successfully." buttons {"OK"} default button "OK"
end tell
