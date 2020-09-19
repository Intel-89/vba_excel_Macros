Sub CreateNewAppointment()

Dim aOutlookVar As Outlook.Application      ' Declaring variables with type
Dim aOutlookApptVar As Outlook.AppointmentItem' Declaring variables with type

Set aOutlookVar = New Outlook.Application           ' Assigning actual instance to variables previously created
Set aOutlookApptVar = ol.CreateItem(olAppointmentItem)' Assigning actual instance to variables previously created

With aOutlookApptVar  ' Population properties from our Appointment variable
    .Subject = "My Appointment Subject"
    .Location = "My Appointment Location, like clasroom #"
    .Start = "Exact Date/Time, like: 09/17/2020 08:00 PM"
    .End = "Exact Date/Time, like: 09/17/2020 08:15 PM"
    .RequiredAttendees = "emails, divided by comas"
    .OptionalAttendees = "emails, divided by comas"
    .Body = "On body of invite will be this text"
    .Send 'Action of sending included in code, can be removed if action not needed
End With
    
End Sub
