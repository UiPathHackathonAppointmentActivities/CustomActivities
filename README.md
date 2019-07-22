# Appointment management custom activities.

The .NET custom activities for management of calendar appointments via [EWS API](https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications).

## Business needs.

While performing automation of synchronization internal Event Management System with appointments in Outlook this custom activities where developed.
They are using EWS API, that allows you to do not install Outlook or any other tool or programms to manage appointments.

## Custom activities description.

### Save Appointment.
Creates or updates appointment.
Has two overload groups for input: 
- _Appointment_ itself. If you want to update existing one.
- _AppointmentDetails_. If you want to create new one.

Both of them returns Id of appointment.

### Load Appointment.
Loads appointment by Id or specific data.
Has two overload groups for input: 
- _FindById_. If you know id.
- _FindByInfo_. If you need to find appointment by specific data. E.g. subject and date.

Returns appointment instance. 

### Delete Appointment.
Deletes appointment by Id. As simple as that.

### Resolve Names.
It is the same functionality as Check names in outlook.
As input - collection of strings with attendee names.
As output - collection of Attendee.

### Get Notification History.
By Given subject and start date returns top **N** notifications sent.

### Forward To Remind.
Forwards invitations to optional or required attendees who has not reacted on request (responce type NoResponseReceived or Unknown).

### Determine Weekly Pattern
A little bit specific activity with logic of determining pattern of recurrence by datatable with occurrences.
Required columns: Start (of DateTime), End (of DateTime), Location (of string).

### Synchronize Appointment
Another specific activity to synchronization of appointment data with server.

## Example
Example of usage could be found [here](https://github.com/UiPathHackathonAppointmentActivities/ShowCase)

## Licenses.

Copyright (c) 2019 EPAM LLC.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.