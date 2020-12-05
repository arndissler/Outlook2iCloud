set _sourceCalendarName to ""
set _iCloudCalendarName to ""

tell application "System Events"
	set propertyListFileName to "~/Library/Preferences/com.arndissler.outlook-to-icloud-sync.plist"
	if not (exists file propertyListFileName) then
		log "ERROR: settings cannot be read from " & propertyListFileName
		return
	end if
	tell property list file propertyListFileName
		set {sourceCalendarName, iCloudCalendarName} to property list items
		set {value:_sourceCalendarName} to sourceCalendarName
		set {value:_iCloudCalendarName} to iCloudCalendarName
	end tell
end tell

tell application "Microsoft Outlook"
	set sourceCalendars to every calendar whose name is _sourceCalendarName

	set sourceCalendar to missing value

	repeat with someCalendar in sourceCalendars
		set numberOfEventsInCalendar to count every calendar event of someCalendar
		if numberOfEventsInCalendar is greater than 0 then
			set sourceCalendar to someCalendar
			exit repeat
		end if
	end repeat

	if sourceCalendar is missing value then
		log "ERROR: source calendar '" & sourceCalendarName & "'cannot be identified"
		display dialog message "Synchroization failed, cannot find Outlook source calendar '" & sourceCalendarName & "'." with title "Outlook to iCloud sync" with icon stop

		return
	end if

	set {name:_calendarName} to sourceCalendar
	set allEvents to every calendar event of sourceCalendar
	set allEventCount to count allEvents

	display notification "Starting synchronization." with title "Outlook to iCloud sync"


	log "INFO: Cleaning up iCloud calendar '" & _iCloudCalendarName & "'..."
	tell application "Calendar"
		tell calendar _iCloudCalendarName
			set allDestinationCalendarEvents to events -- of calendar whose name is eq to _sourceCalendarName
			set countToBeRemoved to count events

			log "INFO: Removing all " & countToBeRemoved & " events from iCloud calendar"

			repeat with singleEvent in allDestinationCalendarEvents
				delete singleEvent
			end repeat
		end tell
	end tell

	log "INFO: Cleanup done"
	log "INFO: Syncing " & allEventCount & " event(s) from Outlook calendar '" & _calendarName & "' to iCloud"

	set limit to 65535
	set i to 0

	repeat with currentEvent in allEvents
		set i to i + 1
		# log "...sync event no. " & i
		set {subject:_subject, plain text content:_description, start time:_startTime, end time:_endTime, all day flag:_hasAllDayFlag, is recurring:_isRecurring, recurrence:_recurrence, account:_account, organizer:_organizer} to currentEvent
		set {name:_accountName} to _account
		set _eventRecurrence to missing value

		if _subject starts with "Canceled:" then
			log "INFO: Event canceled, skipping '" & _subject & "'"
		else
			if _isRecurring then
				if _recurrence is not equal to missing value then
					set {occurrence interval:_recurrenceInterval, recurrence type:_recurrenceType, end date:_endData} to _recurrence

					set _eventRecurringEnd to ""

					set {end type:_endType} to _endData

					if _endType is equal to no end type then
						# probably nothing to do?
					else
						set {data:_data} to _endData
						if _endType is equal to end date type then
							set {day:_day, year:_year} to _data
							set _month to text -2 thru -1 of ("0" & ((month of _data) * 1))
							set _eventRecurringEnd to ";UNTIL=" & _year & " " & _month & " " & _day
						else if _endType is equal to end numbered type then
							set _eventRecurringEnd to ";COUNT=" & _data
						end if
					end if

					if _recurrenceType is equal to daily then
						set _eventRecurrence to "FREQ=DAILY;INTERVAL=" & _recurrenceInterval & _eventRecurringEnd
					else if _recurrenceType is equal to weekly then
						set _eventRecurrence to "FREQ=WEEKLY;INTERVAL=" & _recurrenceInterval & _eventRecurringEnd
					else if _eventRecurrence is equal to relative monthly then
						_eventRecurrence
					else
						log "WARNING: unknown recurrence type found: " & _recurrenceType & ", event subject: " & _subject & ", start date: " & _startTime
					end if
				end if
			end if

			tell application "Calendar"
				tell calendar _iCloudCalendarName
					if _description is equal to missing value then
						set _description to "- no description given -"
					end if

					log _subject & " -- " & _eventRecurrence

					if _eventRecurrence is equal to missing value then
						make new event with properties {summary:_subject, description:_description, start date:_startTime, end date:_endTime, allday event:_hasAllDayFlag}
					else
						make new event with properties {summary:_subject, description:_description, start date:_startTime, end date:_endTime, allday event:_hasAllDayFlag, recurrence:_eventRecurrence}
					end if
				end tell
			end tell
		end if
		if i is equal to limit then exit repeat
	end repeat

	display notification "Sync to iCloud complete." with title "Outlook to iCloud sync" subtitle "Syncronized " & allEventCount & " event(s)"

end tell