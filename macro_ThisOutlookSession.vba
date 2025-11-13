Sub SyncCalendars()
    Call SyncCalendarsCustomerToPersonal
    Call SyncCalendarsPersonalToCustomer
End Sub

Sub SyncCalendarsCustomerToPersonal()
    Call SyncCalendarsParametric( _
        "CHANGEME@Customer", "Calendario", _
        "CHANGEME@Private", "Calendario", _
        -7, 7, _
        "CHANGEME PREFIX: ", _
        "CHANGEME CategoryToSetInPrivate", _
        False, _
        False)
End Sub

Sub SyncCalendarsPersonalToCustomer()
    Call SyncCalendarsParametric( _
        "CHANGEME@Private", "Calendario", _
        "CHANGEME@Customer", "Calendario", _
        -7, 7, _
        "Placeholder", _
        "CHANGEME CategoryToSetInCustomer", _
        True, _
        False)
End Sub

Function appointmentToString(ByRef appt As Outlook.AppointmentItem) As String
    If appt Is Nothing Then
        appointmentToString = "Nothing"
    Else
        appointmentToString = appt.Start & "-" & Format(appt.End, "hh:nn:ss") & " " & appt.Subject
    End If
End Function

Sub copyAppointmentA2B(ByRef apptA As Outlook.AppointmentItem, ByRef apptB As Outlook.AppointmentItem, entryIdA As String, anonymize As Boolean, categoryToSet As String, prefix As String, dryRun As Boolean)
    ' Popola i campi di B
    apptB.Start = apptA.Start
    apptB.End = apptA.End
    apptB.MeetingStatus = apptA.MeetingStatus
    apptB.BusyStatus = apptA.BusyStatus
    apptB.Categories = categoryToSet
    If anonymize Then
        apptB.Subject = prefix
    Else
        apptB.Subject = prefix & apptA.Subject
        apptB.Body = apptA.Body
        apptB.RequiredAttendees = apptA.RequiredAttendees
        apptB.Location = apptA.Location
        Dim rA, rB As Outlook.Recipient
        For Each rA In apptA.Recipients
            Set rB = apptB.Recipients.Add(rA.Address)
            rB.Type = rA.Type
        Next rA
        apptB.Recipients.ResolveAll
    End If
    
    ' Popola i campi di ricorrenza
    ' DISABILITATO perché per misteriosi motivi si creano eventi duplicati e vecchi di mesi
    ' Serve ulteriore indagine su come sono rappresentati gli appuntamenti ricorrenti in Outlook e come gestirli
    If apptA.IsRecurring And False Then
        Dim patternA, patternB As Outlook.RecurrencePattern
        Set patternA = apptA.GetRecurrencePattern
        Set patternB = apptB.GetRecurrencePattern
        patternB.RecurrenceType = patternA.RecurrenceType
        If patternA.RecurrenceType = olRecursYearly Or patternA.RecurrenceType = olRecursYearNth Then
            patternB.MonthOfYear = patternA.MonthOfYear
        End If
        If patternA.RecurrenceType = olRecursWeekly Or patternA.RecurrenceType = olRecursMonthNth Or patternA.RecurrenceType = olRecursYearNth Then
            patternB.DayOfWeekMask = patternA.DayOfWeekMask
        End If
        If patternA.RecurrenceType = olRecursMonthNth Or patternA.RecurrenceType = olRecursYearNth Then
            patternB.Instance = patternA.Instance
        End If
        If patternA.NoEndDate Then
            patternB.Occurrences = patternA.Occurrences
        Else
            patternB.PatternEndDate = patternA.PatternEndDate
        End If
        
        patternB.Interval = patternA.Interval
        patternB.PatternStartDate = patternA.PatternStartDate
        'patternB.NoEndDate = patternA.NoEndDate
    End If
    
    ' Setta il riferimento da B verso A
    If entryIdA <> "" Then
        Call setPropertySourceId(apptB, entryIdA)
    End If
    
    ' Salva
    If Not dryRun Then
        apptB.Save
    End If
End Sub

Function getPropertySourceId(ByRef appt As Outlook.AppointmentItem) As String
    If Not appt Is Nothing Then
        Dim prop As Outlook.UserProperty
        Set prop = appt.UserProperties.Find("SourceID")
        If Not prop Is Nothing Then
            getPropertySourceId = prop.value
        Else
            getPropertySourceId = ""
        End If
    Else
        getPropertySourceId = ""
    End If
End Function

Sub setPropertySourceId(ByRef appt As Outlook.AppointmentItem, value As String)
    If Not appt Is Nothing Then
        Dim prop As Outlook.UserProperty
        Set prop = appt.UserProperties.Add("SourceID", olText, True)
        prop.value = value
    End If
End Sub

Function getAppointmentHomonymousKey(ByRef appt As Outlook.AppointmentItem) As String
    If Not appt Is Nothing Then
        getAppointmentHomonymousKey = appointmentToString(appt) ' appt.Subject & appt.Start & appt.End
    Else
        getAppointmentHomonymousKey = ""
    End If
End Function

Sub SyncCalendarsParametric(accountA As String, folderA As String, accountB As String, folderB As String, startDateOffset As Integer, endDateOffset As Integer, prefix As String, categoryToSet As String, anonymize As Boolean, dryRun As Boolean)
    Dim ns As Outlook.NameSpace
    Dim calendarA As Outlook.Folder
    Dim calendarB As Outlook.Folder
    Dim apptA As Outlook.AppointmentItem
    Dim apptB As Outlook.AppointmentItem
    Dim apptFound As Outlook.AppointmentItem
    Dim itemsA As Outlook.Items
    Dim itemsB As Outlook.Items
    Dim sourceIDs As Collection
    Dim Item As Object
    Dim found As Boolean
    Dim startDate As Date
    Dim startDateCheck As Date
    Dim endDate As Date
    Dim endDateCheck As Date
    Dim entryIdA As String
    Dim entryIdAFromProperty As String
    Dim entryIdB As String
    Dim skipMeeting As Boolean
    Dim countA, countB As Integer: countA = countB = 0
    Dim mapApptBBySourceID As Object
    Dim mapApptBBySubjectStartEnd As Object
    Dim mapKey As String
    
    ' Inizializzazione
    Set sourceIDs = New Collection
    Set ns = Application.GetNamespace("MAPI")
    Set mapApptBBySourceID = CreateObject("Scripting.Dictionary")
    Set mapApptBBySubjectStartEnd = CreateObject("Scripting.Dictionary")
    Set calendarA = ns.Folders(accountA).Folders(folderA)
    Set calendarB = ns.Folders(accountB).Folders(folderB)

    ' Intervallo di date da sincronizzare
    startDate = Date + startDateOffset
    endDate = Date + endDateOffset
    ' I match nel calendario B vengono ricercati solo nell'intervallo check. Per ragioni di performance
    startDateCheck = startDate - 7
    endDateCheck = endDate + 30

    ' Filtra i calendari per intervallo di date
    Set itemsA = calendarA.Items
    itemsA.IncludeRecurrences = True
    itemsA.Sort "[Start]"
    Set itemsA = itemsA.Restrict("[Start] >= '" & Format(startDate, "ddddd hh:nn AMPM") & "' AND [Start] <= '" & Format(endDate, "ddddd hh:nn AMPM") & "'")
    
    Set itemsB = calendarB.Items
    itemsB.IncludeRecurrences = True
    itemsB.Sort "[Start]"
    Set itemsB = itemsB.Restrict("[Start] >= '" & startDateCheck & "' AND [Start] <= '" & endDateCheck & "'")
    
    ' Conta gli incontri. La proprietà itemsA.Count non funziona e restituisce INTEGER_MAX
    For Each Item In itemsA
        countA = countA + 1
    Next
    For Each Item In itemsB
        countB = countB + 1
    Next

    ' Logga l'avvio della sincronizzazione
    Debug.Print "Sincronizzazione dal " & startDate & " al " & endDate & " avviata il " & Now
    Debug.Print "Numero eventi A=" & accountA & "=" & countA & " --> B=" & accountB & "=" & countB
    
    ' Indicizzazione appuntamenti di B per SourceID e per chiave di omonimia (Subject/Start/End)
    ' Con l'indicizzazione si riducono notevolmente i tempi di esecuzione perché gli eventi del calendario B vengono iterati una sola volta
    For Each apptB In itemsB
        If TypeOf apptB Is Outlook.AppointmentItem Then
            If apptB.Start >= startDateCheck And apptB.Start <= endDateCheck Then
                entryIdB = getPropertySourceId(apptB)
                If entryIdB <> "" Then
                    ' Indicizza per SourceID tutti gli appuntamenti che hanno il SourceID
                    If mapApptBBySourceID.Exists(entryIdB) Then
                        Debug.Print "ERROR!! Chiave duplicata in mapApptBBySourceID: '" & entryIdB & "' per " & appointmentToString(apptB)
                    Else
                        mapApptBBySourceID.Add entryIdB, apptB
                    End If
                Else
                    ' La verifica omonimia viene fatta solo verso gli appuntamenti che non hanno il SourceID, cioé che non sono gestiti da questo script
                    mapKey = getAppointmentHomonymousKey(apptB)
                    If Not mapApptBBySubjectStartEnd.Exists(mapKey) Then
                        ' In questo caso se c'è colisione di chiave non è un errore
                        ' perché non interessa risalire all'appuntamento esatto di B ma è sufficiente sapere se esiste almeno un meeting omonimo
                        mapApptBBySubjectStartEnd.Add mapKey, Nothing
                    End If
                End If
                
            End If
        End If
    Next

    ' Per ogni appuntamento di A
    For Each Item In itemsA
        If TypeOf Item Is Outlook.AppointmentItem Then
            Set apptA = Item

            ' Filtra gli appuntamenti nel calendario A per intervallo di date (doppio check rispetto al precedente)
            If apptA.Start >= startDate And apptA.Start <= endDate Then
                ' Negli incontri ricorrenti EntryID è univoco per tutta serie
                ' e cambia valore solo in un singolo incontro se viene spostato rispetto alla pianificazione della serie.
                ' Per cui serve un ulteriore campo per renderlo univoco per quell'appuntamento
                entryIdA = apptA.EntryID & "#" & apptA.Start
                sourceIDs.Add entryIdA
                Set apptFound = Nothing
                skipMeeting = False
                entryIdAFromProperty = getPropertySourceId(apptA)
                
                ' Copia anche gli incontri annullati (per questo il check è disabilitato)
                'If apptA.MeetingStatus = olMeetingReceivedAndCanceled Then
                '    ' Salta quelli cancellati
                '    skipMeeting = True
                '    ' Debug.Print "Saltato incontro cancellato:  " & appointmentToString(apptA)
                'Else
                If entryIdAFromProperty <> "" Then
                    ' Salta quelli che hanno la property SourceID, cioè che sono stati creati da questo script
                    skipMeeting = True
                    'Debug.Print "Saltato incontro già copiato: " & appointmentToString(apptA)
                ElseIf apptA.AllDayEvent And apptA.BusyStatus = olFree Then
                    ' Salta quelli a giornata intera dove si è liberi perché tipicamente sono appunti o promemoria
                    skipMeeting = True
                    'Debug.Print "Saltato incontro giornaliero: " & appointmentToString(apptA)
                Else
                    ' Recupera il placeholder corrispondente per chiave
                    If mapApptBBySourceID.Exists(entryIdA) Then
                        Set apptFound = mapApptBBySourceID(entryIdA)
                    Else
                        mapKey = getAppointmentHomonymousKey(apptA)
                        If mapApptBBySubjectStartEnd.Exists(mapKey) Then
                            Set apptFound = mapApptBBySubjectStartEnd(mapKey)
                            skipMeeting = True
                        End If
                    End If
                    
                End If

                If skipMeeting Then
                    ' Salta l'elaborazione
                ElseIf apptFound Is Nothing Then
                    ' Crea un nuovo incontro
                    Set apptB = calendarB.Items.Add(olAppointmentItem)
                    Call copyAppointmentA2B(apptA, apptB, entryIdA, anonymize, categoryToSet, prefix, dryRun)
                    If anonymize Then
                        Debug.Print "Creato placeholder anonimo:   " & appointmentToString(apptB) & " - per: " & apptA.Subject
                    Else
                        Debug.Print "Creato placeholder incontro:  " & appointmentToString(apptB)
                    End If
                    
                ElseIf apptFound.Start <> apptA.Start Or apptFound.End <> apptA.End Or (Not anonymize And apptFound.Subject <> prefix & apptA.Subject) Then
                    ' Aggiorna l'incontro
                    apptFound.BusyStatus = apptA.BusyStatus
                    apptFound.Start = apptA.Start
                    apptFound.End = apptA.End
                    If Not anonymize Then
                        apptFound.Subject = prefix & apptA.Subject
                    End If
                    If Not dryRun Then
                        apptFound.Save
                    End If
                    Debug.Print "Aggiornato placeholder:       " & appointmentToString(apptA)

                Else
                    If apptFound.BusyStatus <> apptA.BusyStatus Then
                        ' Aggiorna solo lo stato
                        apptFound.BusyStatus = apptA.BusyStatus
                        If Not dryRun Then
                            apptFound.Save
                        End If
                        Debug.Print "Aggiornato stato placeholder: " & appointmentToString(apptA)
                    Else
                        'Debug.Print "Saltato incontro già copiato: " & appointmentToString(apptA)
                    End If
                End If
            Else
                'Debug.Print "Ignorato (fuori intervallo):  " & appointmentToString(apptA)
            End If
        End If
    Next
    
    ' Fase 2: rimuovi placeholder orfani
    For Each apptB In itemsB
        If TypeOf apptB Is Outlook.AppointmentItem Then
            ' Filtra per data (doppio check)
            If apptB.Start >= startDateCheck And apptB.Start <= endDateCheck Then
                ' Considera solo gli elementi copiati, cioè con SourceID valorizzato
                entryIdB = getPropertySourceId(apptB)
                If entryIdB <> "" Then
                    found = False
                    For Each EntryID In sourceIDs
                        If EntryID = entryIdB Then
                            found = True
                            Exit For
                        End If
                    Next
                    If Not found Then
                        ' Elimina solo i placeholder orfani a partire da startDate
                        ' perché per quelli più vecchi non si può sempre risolvere il corrispondente incontro di A a causa della non corrispondenza dei filtri per data
                        ' Con la condizione "apptB.End < endDate" si eliminerebbero i placeholder di incontri successivi alla data fine che potrebbero essere stati creati da altre sincronizzazioni con intervallo di date più ampio, per cui la condizione è omessa.
                        ' Con la condizione "apptB.Start >= startDate + 1" NON si eliminano i placeholder antecedenti alla data inizio che potrebbero essere stati creati da altre sincronizzazioni con intervallo di date più ampio
                        If apptB.Start >= startDate + 1 Then
                            Debug.Print "Eliminato placeholder orfano: " & appointmentToString(apptB)
                            If Not dryRun Then
                                apptB.Delete
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next

    Debug.Print "Sincronizzazione completata alle " & Time
    Debug.Print ""
End Sub



