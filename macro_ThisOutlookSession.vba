
Sub SyncCalendarsCustomerToPrivate()
    Call SyncCalendarsParametric( _
        "CHANGEME@Customer", "Calendario", _
        "CHANGEME@Private", "Calendario", _
        -7, 7, _
        "CHANGEME PREFIX: ", _
        "CHANGEME CategoryToSetInPrivate", _
        False, _
        False)
End Sub

Sub SyncCalendarsPrivateToCustomer()
    Call SyncCalendarsParametric( _
        "CHANGEME@Private", "Calendario", _
        "CHANGEME@Customer", "Calendario", _
        -7, 7, _
        "Placeholder", _
        "CHANGEME CategoryToSetInCustomer", _
        True, _
        False)
End Sub

Sub copyAppointmentA2B(ByRef apptA As Outlook.AppointmentItem, ByRef apptB As Outlook.AppointmentItem, entryIdA As String, anonymize As Boolean, categoryToSet As String, prefix As String, dryRun As Boolean)
    ' Popola i campi di B
    apptB.Start = apptA.Start
    apptB.End = apptA.End
    apptB.BusyStatus = apptA.BusyStatus
    apptB.Categories = categoryToSet
    If anonymize Then
        apptB.Subject = prefix
    Else
        apptB.Subject = prefix & apptA.Subject
        apptB.Body = apptA.Body
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
    
    ' Legge il SourceId se mancante
    If IsEmpty(entryIdA) Then
        Dim prop As Outlook.UserProperty
        Set prop = apptA.UserProperties.Find("SourceID")
        If Not prop Is Nothing Then
            entryIdA = prop.Value
        End If
    End If
    
    ' Setta il riferimento da B verso A
    If Not IsEmpty(entryIdA) Then
        Set prop = apptB.UserProperties.Add("SourceID", olText, True)
        prop.Value = entryIdA
    End If
    
    ' Salva
    If Not dryRun Then
        apptB.Save
    End If
End Sub

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
    Dim prop As Outlook.UserProperty
    Dim Item As Object
    Dim found As Boolean
    Dim startDate As Date
    Dim startDateCheck As Date
    Dim endDate As Date
    Dim endDateCheck As Date
    Dim entryIdA As String
    Dim entryIdB As String
    Dim skipMeeting As Boolean
    Dim countA, countB As Integer: countA = countB = 0
    
    ' Inizializzazione
    Set sourceIDs = New Collection
    Set ns = Application.GetNamespace("MAPI")

    ' Intervallo di date da sincronizzare
    startDate = Date + startDateOffset
    endDate = Date + endDateOffset
    ' I match nel calendario B vengono ricercati solo nell'intervallo check. Per ragioni di performance
    startDateCheck = startDate - 7
    endDateCheck = endDate + 30

    ' Sostituisci con i valori corretti del proprio account e propria configurazione
    Set calendarA = ns.Folders(accountA).Folders(folderA)
    Set calendarB = ns.Folders(accountB).Folders(folderB)

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

    ' Avvia la sincronizzazione
    Debug.Print "Sincronizzazione dal " & startDate & " al " & endDate & " avviata il " & Now
    Debug.Print "Eventi in A=" & countA & " e in B=" & countB

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
                Set prop = apptA.UserProperties.Find("SourceID")
                
                ' Salta quelli cancellati
                If apptA.MeetingStatus = olMeetingReceivedAndCanceled Then
                    skipMeeting = True
                    'Debug.Print "Saltato incontro cancellato: " & apptA.Subject & " - " & apptA.Start & " - " & apptA.End
                ElseIf Not prop Is Nothing Then
                    skipMeeting = True
                    'Debug.Print "Saltato incontro già copiato: " & apptA.Subject & " - " & apptA.Start & " - " & apptA.End
                ElseIf apptA.AllDayEvent Then
                    skipMeeting = True
                    'Debug.Print "Saltato incontro di giornata intera: " & apptA.Subject & " - " & apptA.Start
                Else
                    ' Cerca se esiste già un placeholder
                    For Each apptB In itemsB
                        If TypeOf apptB Is Outlook.AppointmentItem Then
                            'If apptB.Subject = prefix & apptA.Subject And apptB.Start = apptA.Start And apptB.End = apptA.End Then
                            'If InStr(apptB.Body, textToSearch) > 0 Then
                            If apptB.Start >= startDateCheck And apptB.Start <= endDateCheck Then
                                ' Cerca il match per SourceId
                                Set prop = apptB.UserProperties.Find("SourceID")
                                If Not prop Is Nothing Then
                                    If prop.Value = entryIdA Then
                                        Set apptFound = apptB
                                        Exit For
                                    End If
                                End If
                                ' Cerca se ci sono meeting senza SourceId ma con stesso subject/inizio/fine e in tal caso salta quello corrente perché già gestito a mano
                                If apptA.Subject = apptB.Subject And apptA.Start = apptB.Start And apptA.End = apptB.End Then
                                    'Debug.Print "Saltato incontro omonimo: " & apptA.Subject & " - " & apptA.Start & " - " & apptA.End
                                    skipMeeting = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                    
                End If

                If skipMeeting Then
                    ' Salta l'elaborazione
                ElseIf apptFound Is Nothing Then
                    ' Crea un nuovo incontro
                    Set apptB = calendarB.Items.Add(olAppointmentItem)
                    Call copyAppointmentA2B(apptA, apptB, entryIdA, anonymize, categoryToSet, prefix, dryRun)
                    If anonymize Then
                        Debug.Print "Creato placeholder: " & apptB.Subject & " - " & apptB.Start & " - Da: " & apptA.Subject
                    Else
                        Debug.Print "Creato placeholder: " & apptB.Subject & " - " & apptB.Start
                    End If
                    
                ElseIf apptFound.Start <> apptA.Start Or apptFound.End <> apptA.End Then
                    ' Aggiorna l'incontro spostato ricreandolo da capo
                    Set apptB = calendarB.Items.Add(olAppointmentItem)
                    Call copyAppointmentA2B(apptA, apptB, entryIdA, anonymize, categoryToSet, prefix, dryRun)
                    If Not dryRun Then
                        apptFound.Delete
                    End If
                    Debug.Print "Aggiornato placeholder per: " & apptA.Subject

                Else
                    If apptFound.BusyStatus <> apptA.BusyStatus Then
                        ' Aggiorna solo lo stato
                        apptFound.BusyStatus = apptA.BusyStatus
                        If Not dryRun Then
                            apptFound.Save
                        End If
                        Debug.Print "Aggiornato stato: " & apptA.Subject & " - " & apptA.Start
                    Else
                        'Debug.Print "Già presente: " & apptA.Subject & " - " & apptA.Start
                    End If
                End If
            Else
                'Debug.Print "Ignorato (fuori intervallo): " & apptA.Subject & " - " & apptA.Start
            End If
        End If
    Next
    
    ' Fase 2: rimuovi placeholder orfani
    For Each apptB In calendarB.Items
        If TypeOf apptB Is Outlook.AppointmentItem Then
            If apptB.Start >= startDateCheck And apptB.Start <= endDateCheck Then
                Set prop = apptB.UserProperties.Find("SourceID")
                If Not prop Is Nothing Then
                    idValue = prop.Value
                    On Error Resume Next
                    found = False
                    For Each EntryID In sourceIDs
                        If EntryID = idValue Then
                            found = True
                            Exit For
                        End If
                    Next
                    On Error GoTo 0
                    If Not found Then
                        Debug.Print "Eliminato placeholder orfano: " & apptB.Subject
                        If Not dryRun Then
                            apptB.Delete
                        End If
                    End If
                End If
            End If
        End If
    Next

    Debug.Print "Sincronizzazione completata alle " & Time
End Sub



