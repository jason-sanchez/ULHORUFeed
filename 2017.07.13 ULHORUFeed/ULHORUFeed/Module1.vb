'20131024 - added transcription processing
'20131111 - Added global order comment routine and code to use value in test_master.
'20140121 - mods for running on new test system. ALso don't update placer order number during
'update of test_detail table.
'20140327 - issue with insert in updateDetail. Memo was left out of sql startment.
'20140513 - Addendum handling code in update detail routine.
'20140605 - handle appostrophies in order comment
'20140805 - Mods for W3 Production.
'201409122 - fixed spelling of Abnormal Flags to AbnormalFlags with no space
'20150218 - remove redundant replace function call in checkOrderComment
'20150219 - Add'l mod for comment apostrophies in add detail.
'20150224 - ...and updateDetail.


'20150302 - ULHORUFeed
'20150409 - vs2013 version

'20151222 - Mods for KY2 Production Interface.

Imports System
Imports System.IO
Imports System.Collections
Imports System.Data.SqlClient
Imports System.Diagnostics
Module Module1
    Dim gblLogString As String = ""
    Dim connectionString As String

    Dim dictNVP As New Scripting.Dictionary
    Dim sql As String
    Dim gblSQL As String = ""
    Dim sql2 As String
    Dim myfile As StreamReader
    Dim dir As String
    Dim gblboolFirstNTE As Boolean = True
    Dim globalError As Boolean = False
    
    Dim theError As String
    
    Dim gblTheFile As String = ""
    Dim gblFileName As String = ""
    Dim theArray(0 To 200, 0 To 1) As Integer
    Dim gblOBX As Integer = 0
    Dim gblStrApp As String = ""
    Dim gblStrSecurity As String = ""
    Public objIniFile As New iniFile("c:\newfeeds\HL7Mapper.ini") '20151222
    'Public objIniFile As New iniFile("C:\KY2 Test Environment\HL7Mapper.ini")
    Dim strInputDirectory As String = ""
    Dim strOutputDirectory As String = ""
    Public thefile As FileInfo
    Dim strMapperFile As String = ""

    '20131111
    Dim gblOrderComment As String = ""

    Dim strLogDirectory As String = "" '20140205

    Sub main()
        'declarations for split function
        Dim delimStr As String = "="
        Dim delimiter As Char() = delimStr.ToCharArray()
        Dim theFile As FileInfo
        'declarations for stream reader
        Dim strLine As String = ""
        Dim sql As String = ""

        'setup directory

        strOutputDirectory = objIniFile.GetString("ULHORU", "ULHORUoutputdirectory", "(none)") '
        strMapperFile = objIniFile.GetString("ULHORU", "ULHORUmapper", "(none)")
        '20140205 - add logfile location
        strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")

        Dim dirs As String() = Directory.GetFiles(strOutputDirectory, "NVP.*")

        'declarations and external assignments for database operations
        'connectionString = "server=10.48.64.5\sqlexpress;database=ITWULHCernerTest;uid=sysmax;pwd=Condor!"
        connectionString = "server=10.48.242.249,1433;database=McareULHCerner;uid=sysmax;pwd=Condor!" '20151222

        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        objCommand.Connection = myConnection
        Dim dataReader As SqlDataReader

        For Each dir In dirs

            theFile = New FileInfo(dir)
            gblFileName = theFile.Name

            If theFile.Extension <> ".$#$" Then
                Try
                    globalError = False
                    '1.set up the streamreader to get a file
                    myfile = File.OpenText(dir)
                    'and read the first line
                    'strLine = myfile.ReadLine()
                    Try
                        Do While Not myfile.EndOfStream
                            Dim myArray As String() = Nothing
                            strLine = myfile.ReadLine()
                            If strLine <> "" Then
                                myArray = strLine.Split(delimiter, 2)
                                'add array key and item to hashtable
                                dictNVP.Add(myArray(0), myArray(1))

                            End If
                        Loop
                    Catch ex As Exception
                        'make copy in the problems directory delete any previous ones with same name
                        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "problems\" & theFile.Name)
                        fi2.Delete()
                        theFile.CopyTo(strOutputDirectory & "problems\" & theFile.Name)

                        gblLogString = gblLogString & "Dictionary Error" & " - " & theFile.Name & vbCrLf
                        gblLogString = gblLogString & ex.Message & vbCrLf
                        writeToLog(gblLogString)
                        'get rid of the file so it doesn't mess up the next run.
                        myfile.Close()
                        If theFile.Exists Then
                            theFile.Delete()
                            Exit Sub
                        End If
                    End Try
                    myfile.Close()

                    '===================================================================================================
                    'call subdirectories here

                    '
                    Call BuildArray(dictNVP)
                    Call checkOBX(dictNVP)

                    '20131111
                    Call checkOrderComment(dictNVP)
                    '20131111 end =======================================

                    'gblStrApp = dictNVP.Item("SendingApplication")
                    'gblStrSecurity = dictNVP.Item("Security")
                    sql = "select id from [test_master] where PlacerOrderNumber = '" & dictNVP.Item("Placer Order Number") & "' "
                    sql = sql & "AND patient_accountNumber = '" & extractPanum(dictNVP.Item("Patient_AccountNumber")) & "'"

                    objCommand.CommandText = sql
                    myConnection.Open()
                    dataReader = objCommand.ExecuteReader()
                    Dim addRecord As Boolean = True
                    If dataReader.HasRows Then
                        addRecord = False
                    Else
                        addRecord = True
                    End If
                    myConnection.Close()
                    dataReader.Close()

                    If addRecord Then
                        Call addTest(dictNVP)
                    Else
                        Call updateTest(dictNVP)
                    End If
                Catch ex As Exception
                    myConnection.Close()

                    globalError = True
                    gblLogString = gblLogString & CStr(ex.Message)
                Finally
                    '===================================================================================================
                    dictNVP.RemoveAll()

                    If globalError Then
                        writeToLog(gblLogString)
                        Dim fi2 As FileInfo = New FileInfo(strOutputDirectory & "backup\" & theFile.Name)
                        fi2.Delete()
                        theFile.CopyTo(strOutputDirectory & "backup\" & theFile.Name)
                        theFile.Delete()

                    Else

                        theFile.Delete()

                    End If


                End Try

            End If 'If theFile.Extension <> ".$#$"



        Next



    End Sub
    Public Sub BuildArray(ByVal dictNVP As Scripting.Dictionary)
        Dim counter As Integer
        Dim iOBXSegments As Integer
        Dim iNTESegments As Integer
        Dim arrKeys
        Dim arrItems
        Dim OBXFound As Boolean
        Dim I, J As Integer
        Try
            counter = 0
            iOBXSegments = 0
            gblOBX = 0
            iNTESegments = 0
            OBXFound = False
            '20120207 - increased theArray t0 200
            For I = 0 To 200
                theArray(I, 0) = 0
                theArray(I, 1) = 0
            Next

            I = 1
            J = 1
            'the keys hold the left side of the equal sign i.e. the name in the NVP
            'the items hold the values on the right side of the equal sign
            arrKeys = dictNVP.Keys
            arrItems = dictNVP.Items

            For counter = 0 To dictNVP.Count - 1


                If Left$(arrKeys(counter), 3) = "OBX" Then
                    iOBXSegments = iOBXSegments + 1
                    'OBXFound = True

                End If

                If Left$(arrKeys(counter), 3) = "NTE" Then
                    iNTESegments = iNTESegments + 1

                    'if this is the first NTE and there are no preceeding OBX segments
                    'then it must be associated with the OBR so added it to the master table comments field.
                    If iOBXSegments > 0 Then
                        gblboolFirstNTE = False

                    End If

                    theArray(I, 0) = iOBXSegments
                    theArray(I, 1) = iNTESegments
                    I = I + 1

                End If

            Next

        Catch ex As Exception
            globalError = True
            gblLogString = gblLogString & CStr(ex.Message)
        End Try
    End Sub
    Public Sub checkOBX(ByVal dictNVP As Scripting.Dictionary)

        Dim counter As Integer
        Dim iOBXSegments As Integer

        Dim arrKeys
        Dim arrItems
        counter = 0
        iOBXSegments = 0
        gblOBX = 0
        Try
            'the keys hold the left side of the equal sign i.e. the name in the NVP
            'the items hold the values on the right side of the equal sign
            arrKeys = dictNVP.Keys
            arrItems = dictNVP.Items

            For counter = 0 To dictNVP.Count - 1

                If Left$(arrKeys(counter), 3) = "OBX" Then
                    iOBXSegments = iOBXSegments + 1
                End If

            Next

            'Debug.Print "no of obx segments: " & iOBXSegments

            'set the global value for further processing
            gblOBX = iOBXSegments
        Catch ex As Exception
            globalError = True
            gblLogString = gblLogString & CStr(ex.Message)
        End Try
    End Sub
    
    Public Sub addTest(ByVal dictNVP As Scripting.Dictionary)
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        objCommand.Connection = myConnection
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        Dim sql As String = ""
        Dim dataReader As SqlDataReader
        Dim lastMasterID As Integer = 0
        Dim intMessageType As Integer = 0
        Try

            If isDolby(Replace(dictNVP.Item("ObservationValue"), "'", "''")) Then
                intMessageType = 1
            End If
            Dim varnow = DateTime.Now
            '1.insert master record
            gblSQL = "Insert into [test_master]"
            gblSQL = gblSQL & "(MessageControlID, PatientID_Internal, PatientLastName, PatientFirstName, PatientMiddleNameInitial, "
            gblSQL = gblSQL & "OBR_4_1, OBR_4_2, OBR_4_3, OBR_3, PlacerOrderNumber, "

            gblSQL = gblSQL & "Patient_AccountNumber, ObservationDateTime, KeyInfo, added, Specimen_ReceivedDateTime, SpecimenSource, message_type, messageTypeID, OBX3, comment)"
            '20131022 - added code to remove the region character from the first position of the mrnum
            gblSQL = gblSQL & "values('" & Replace(dictNVP.Item("MessageControlID"), "'", "''") & "'"
            insertNumber(extractMrnum(dictNVP.Item("mrnum")))
            insertString(dictNVP.Item("PatientLastName"))
            insertString(dictNVP.Item("PatientFirstName"))
            insertString(dictNVP.Item("PatientMiddleNameInitial"))

            insertString(dictNVP.Item("Universal Service Identifier ID"))
            insertString(dictNVP.Item("Universal Service Identifier Text"))
            insertString(dictNVP.Item("Universal Service Identifier Coding System"))
            insertString(dictNVP.Item("Filler Order Number"))
            insertString(dictNVP.Item("Placer Order Number"))

            insertString(extractPanum(dictNVP.Item("Patient_AccountNumber")))
            insertString(ConvertDate(dictNVP.Item("ObservationDateTime")))
            insertNumber(dictNVP.Item("ObservationDateTime"))
            insertString(varnow)
            insertString(ConvertDate(dictNVP.Item("Specimen_ReceivedDateTime")))
            insertString(dictNVP.Item("SpecimenSourceCode"))

            insertNumber(intMessageType)
            insertString(dictNVP("DiagnosticService")) '20131017

            '10/28/2004
            insertString(dictNVP.Item("ObservationIdentifier1"))

            '20131111
            'If gblboolFirstNTE Then
            'insertString(dictNVP("Source of Comment") & ": " & dictNVP("Comment"))
            'Else
            'insertString("")
            'End If

            insertString(gblOrderComment)
            '20131111 end =================================================

            gblSQL = gblSQL & ")"
            updatecommand.CommandText = gblSQL
            myConnection.Open()
            updatecommand.ExecuteNonQuery()
            myConnection.Close()

            '2.get the master id for the record entered
            sql = "select max(id) as maxnum from [test_master]"
            objCommand.CommandText = sql
            myConnection.Open()
            dataReader = objCommand.ExecuteReader()
            While dataReader.Read()
                lastMasterID = dataReader.Item(0)
            End While
            myConnection.Close()
            dataReader.Close()

            Call addDetail(dictNVP, lastMasterID)

        Catch ex As Exception
            globalError = True
            gblLogString = gblLogString & CStr(ex.Message)
        End Try
    End Sub

    Public Sub updateTest(ByVal dictNVP As Scripting.Dictionary)
        Dim temp As String = ""
        Dim tempSecurity As String = ""

        Try
            'temp = Trim(dictNVP.Item("SendingApplication"))
            'tempSecurity = Trim(dictNVP.Item("Security"))

            Call updateDetail(dictNVP)
            '
        Catch ex As Exception
            globalError = True
            gblLogString = gblLogString & CStr(ex.Message)
        End Try
    End Sub
   
    Public Sub updateDetail(ByVal dictNVP As Scripting.Dictionary)
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        objCommand.Connection = myConnection
        Dim dataReader As SqlDataReader
        Dim strRStatus As String = ""
        Dim updateResults As Boolean = False

        Dim updateAddendum As Boolean = False

        Dim sql As String = ""
        Dim I As Integer = 0
        Dim strTemp As String = ""
        Dim arrayCounter As Integer = 0
        Dim varTempComment As String = ""
        Dim strNTETemp As String = ""
        Dim masterID As Long = 0
        Dim strMemo As String = ""
        Try


            For I = 1 To gblOBX                           ' number of obx segments in the current LTW file starting at obx_0002
                strTemp = "" & I
                If Len(strTemp) = 1 Then
                    strTemp = "_000" & strTemp
                ElseIf Len(strTemp) = 2 Then
                    strTemp = "_00" & strTemp
                ElseIf Len(strTemp) = 3 Then
                    strTemp = "_0" & strTemp
                End If

                If I = 1 Then strTemp = ""

                For arrayCounter = 0 To UBound(theArray)
                    If theArray(arrayCounter, 1) = 0 And theArray(arrayCounter, 0) = 1 Then
                        varTempComment = varTempComment & Trim(Replace(dictNVP.Item("Comment"), "'", "''")) & vbCrLf '20150224
                    End If

                    If theArray(arrayCounter, 0) = 1 And theArray(arrayCounter, 1) > 1 Then
                        Select Case theArray(arrayCounter, 1)
                            Case Is < 10
                                strNTETemp = "000" & theArray(arrayCounter, 1)
                                varTempComment = varTempComment & Replace(dictNVP.Item("Comment" & "_" & strNTETemp), "'", "''") & vbCrLf
                            Case Is > 10
                                strNTETemp = "00" & theArray(arrayCounter, 1)
                                varTempComment = varTempComment & Replace(dictNVP.Item("Comment" & "_" & strNTETemp), "'", "''") & vbCrLf
                            Case Is > 100
                                strNTETemp = "0" & theArray(arrayCounter, 1)
                                varTempComment = varTempComment & Replace(dictNVP.Item("Comment" & "_" & strNTETemp), "'", "''") & vbCrLf
                        End Select
                    End If

                Next

                '20140507 - Handle Addemdum here. If it exists, update it else add it as a dolby.
                '======================================================================================================================
                '======================================================================================================================
                '======================================================================================================================

                If UCase(dictNVP.Item("ObservationIdentifier1" & strTemp)) = "ADDENDUM" Then

                    sql = "select [Observation_ResultsStatus] from [test_detail] "
                    'sql = sql & "WHERE keyInfo = " & dictNVP.Item("ObservationDateTime")
                    sql = sql & " Where ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & strTemp), "'", "''") & "'"
                    sql = sql & " AND PlacerOrderNumber = '" & Replace(dictNVP.Item("Placer Order Number"), "'", "''") & "' "
                    'sql = sql & " AND ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & "_" & strTemp), "'", "''") & "' "
                    objCommand.CommandText = sql
                    myConnection.Open()
                    dataReader = objCommand.ExecuteReader()
                    If dataReader.HasRows() Then
                        updateAddendum = True
                    Else
                        updateAddendum = False
                    End If
                    myConnection.Close()
                    dataReader.Close()

                    If updateAddendum Then ' update record
                        sql = "update [test_detail] "

                        sql = sql & "Set  ObservationIdentifierText = '" & Replace(dictNVP.Item("ObservationIdentifier2" & strTemp), "'", "''") & "', "
                        sql = sql & "ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & strTemp), "'", "''") & "', "

                        'sql = sql & "Observation_Value = '" & Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''") & "', "
                        '==================================================================================================
                        If isDolby(dictNVP.Item("ObservationValue" & strTemp)) Then

                            strMemo = ""
                            strMemo = strMemo & addMemo(Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''"))
                            sql = sql & "memo = '" & strMemo & "', "
                            sql = sql & "Observation_Value = '** Addendum **', "

                        Else

                            sql = sql & "Observation_Value = '" & Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''") & "', "

                        End If
                        '==================================================================================================
                        sql = sql & "UnitsID = '" & Replace(dictNVP.Item("UnitsID" & strTemp), "'", "''") & "', "
                        sql = sql & "ReferenceRange = '" & Replace(dictNVP.Item("ReferenceRange" & strTemp), "'", "''") & "', "
                        sql = sql & "Observation_ResultsStatus = '" & Replace(dictNVP.Item("ObservationResultStatus" & strTemp), "'", "''") & "', "

                        sql = sql & "OBR_3 = '" & Replace(dictNVP.Item("Filler Order Number"), "'", "''") & "', "
                        'sql = sql & "PlacerOrderNumber = '" & Replace(dictNVP.Item("Placer Order Number"), "'", "''") & "', "
                        sql = sql & "OBR_4 = '" & Replace(dictNVP.Item("Universal Service Identifier Text" & strTemp), "'", "''") & "', "
                        sql = sql & "AbnormalFlags = '" & Replace(dictNVP.Item("AbnormalFlags" & strTemp), "'", "''") & "', " '20140912
                        '12/14/2005 - removed replace function on varTempComment below
                        sql = sql & "nteComment = '" & varTempComment & "', "
                        sql = sql & "updated = '" & Now & "' "
                        'sql = sql & "WHERE keyInfo = " & dictNVP.Item("ObservationDateTime")
                        sql = sql & " WHERE ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & strTemp), "'", "''") & "'"
                        sql = sql & " AND PlacerOrderNumber = '" & Replace(dictNVP.Item("Placer Order Number"), "'", "''") & "' "
                        'sql = sql & " AND ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & "_" & strTemp), "'", "''") & "' "
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()

                    Else ' add record
                        masterID = 0
                        sql = "select id from [test_master] where PlacerOrderNumber = " & dictNVP.Item("Placer Order Number") & " "
                        sql = sql & "AND patient_accountNumber = '" & extractPanum(dictNVP.Item("Patient_AccountNumber")) & "'"
                        objCommand.CommandText = sql
                        myConnection.Open()
                        dataReader = objCommand.ExecuteReader()
                        If dataReader.HasRows() Then
                            While dataReader.Read()
                                masterID = dataReader.GetInt64(0)

                            End While
                        End If
                        myConnection.Close()
                        dataReader.Close()

                        sql = "Insert into [test_detail]"
                        sql = sql & "(masterID, ValueType, Observation_Value, memo, UnitsID, ReferenceRange,"
                        sql = sql & "Observation_ResultsStatus, KeyInfo, Patient_AccountNumber, "
                        sql = sql & "added, nteComment, PlacerOrderNumber, OBR_3, OBR_4, "
                        sql = sql & "ObservationIdentifierID, ObservationIdentifierText, AbnormalFlags, ObservationDateTime)"
                        sql = sql & "values(" & masterID
                        sql = sql & ",'" & Replace(dictNVP.Item("ValueType" & "_" & strTemp), "'", "''") & "'"


                        'sql = sql & ",'" & Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''") & "'"

                        '==================================================================================================
                        If isDolby(dictNVP.Item("ObservationValue" & strTemp)) Then

                            sql = sql & ",'**Addendum**'"
                            strMemo = ""
                            strMemo = strMemo & addMemo(Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''"))

                            sql = sql & ",'" & Replace(strMemo, "'", "''") & "'"
                        Else
                            sql = sql & ",'" & Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''") & "'"
                            sql = sql & ",NULL"
                        End If
                        '==================================================================================================
                        sql = sql & ",'" & Replace(dictNVP.Item("UnitsID" & strTemp), "'", "''") & "'"
                        sql = sql & ",'" & Replace(dictNVP.Item("ReferenceRange" & strTemp), "'", "''") & "'"
                        sql = sql & ",'" & dictNVP.Item("Observation_ResultStatus" & strTemp) & "'"
                        sql = sql & "," & dictNVP.Item("ObservationDateTime")
                        sql = sql & ",'" & Replace(dictNVP.Item("Patient_AccountNumber"), "'", "''") & "'"
                        sql = sql & ",'" & DateTime.Now & "'"
                        '12/14/2005 - removed replace function on varTempComment below
                        sql = sql & ",'" & varTempComment & "'"

                        sql = sql & ",'" & Replace(dictNVP.Item("Placer Order Number"), "'", "''") & "'"
                        sql = sql & ",'" & Replace(dictNVP.Item("Filler Order Number"), "'", "''") & "'"
                        sql = sql & ",'" & Replace(dictNVP.Item("Universal Service Identifier Text"), "'", "''") & "'"
                        sql = sql & ",'" & Replace(dictNVP.Item("ObservationIdentifier1" & strTemp), "'", "''") & "'" ' ObservationIdentifierID
                        sql = sql & ",'" & Replace(dictNVP.Item("ObservationIdentifier2" & strTemp), "'", "''") & "'" ' ObservationIdentifierText
                        sql = sql & ",'" & Replace(dictNVP.Item("AbnormalFlags" & strTemp), "'", "''") & "'" '20140912

                        sql = sql & ",'" & Replace(ConvertDate(dictNVP.Item("ObservationDateTime")), "'", "''") & "'"


                        sql = sql & ")"
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()
                    End If

                End If 'If UCase(dictNVP.Item("ObservationIdentifierID" & strTemp)) = "ADDENDUM"
                '======================================================================================================================
                '======================================================================================================================
                '======================================================================================================================

                If UCase(dictNVP.Item("ObservationResultStatus" & strTemp)) = "C" Then
                    'if the OBX is a C status, update it
                    sql = "update [test_detail] "

                    sql = sql & "Set  ObservationIdentifierText = '" & Replace(dictNVP.Item("ObservationIdentifier2" & strTemp), "'", "''") & "', "
                    sql = sql & "ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & strTemp), "'", "''") & "', "

                    'sql = sql & "Observation_Value = '" & Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''") & "', "
                    '==================================================================================================
                    If isDolby(dictNVP.Item("ObservationValue" & strTemp)) Then

                        strMemo = ""
                        strMemo = strMemo & addMemo(Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''"))
                        sql = sql & "memo = '" & strMemo & "', "
                        sql = sql & "Observation_Value = '** Report **', "

                    Else

                        sql = sql & "Observation_Value = '" & Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''") & "', "

                    End If
                    '==================================================================================================
                    sql = sql & "UnitsID = '" & Replace(dictNVP.Item("UnitsID" & strTemp), "'", "''") & "', "
                    sql = sql & "ReferenceRange = '" & Replace(dictNVP.Item("ReferenceRange" & strTemp), "'", "''") & "', "
                    sql = sql & "Observation_ResultsStatus = '" & Replace(dictNVP.Item("ObservationResultStatus" & strTemp), "'", "''") & "', "

                    sql = sql & "OBR_3 = '" & Replace(dictNVP.Item("Filler Order Number"), "'", "''") & "', "
                    'sql = sql & "PlacerOrderNumber = '" & Replace(dictNVP.Item("Placer Order Number"), "'", "''") & "', "
                    sql = sql & "OBR_4 = '" & Replace(dictNVP.Item("Universal Service Identifier Text" & strTemp), "'", "''") & "', "
                    sql = sql & "AbnormalFlags = '" & Replace(dictNVP.Item("AbnormalFlags" & strTemp), "'", "''") & "', " '20140912
                    '12/14/2005 - removed replace function on varTempComment below
                    sql = sql & "nteComment = '" & varTempComment & "', "
                    sql = sql & "updated = '" & Now & "' "
                    'sql = sql & "WHERE keyInfo = " & dictNVP.Item("ObservationDateTime")
                    sql = sql & " WHERE ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & strTemp), "'", "''") & "'"
                    sql = sql & " AND PlacerOrderNumber = '" & Replace(dictNVP.Item("Placer Order Number"), "'", "''") & "' "
                    'sql = sql & " AND ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & "_" & strTemp), "'", "''") & "' "
                    updatecommand.CommandText = sql
                    myConnection.Open()
                    updatecommand.ExecuteNonQuery()
                    myConnection.Close()
                ElseIf (UCase(dictNVP.Item("ObservationResultStatus" & strTemp)) = "F") Or (UCase(dictNVP.Item("ObservationResultStatus" & strTemp)) = "P") Then

                    sql = "select [Observation_ResultsStatus] from [test_detail] "
                    'sql = sql & "WHERE keyInfo = " & dictNVP.Item("ObservationDateTime")
                    sql = sql & " Where ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & strTemp), "'", "''") & "'"
                    sql = sql & " AND PlacerOrderNumber = '" & Replace(dictNVP.Item("Placer Order Number"), "'", "''") & "' "
                    'sql = sql & " AND ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & "_" & strTemp), "'", "''") & "' "
                    objCommand.CommandText = sql
                    myConnection.Open()
                    dataReader = objCommand.ExecuteReader()
                    If dataReader.HasRows() Then
                        While dataReader.Read
                            strRStatus = dataReader.GetString(0)
                        End While
                        updateResults = True
                    End If
                    myConnection.Close()
                    dataReader.Close()

                    If updateResults Then
                        If strRStatus = "P" Then

                            sql = "update [test_detail] "

                            sql = sql & "Set  ObservationIdentifierText = '" & Replace(dictNVP.Item("ObservationIdentifier2" & strTemp), "'", "''") & "', "
                            sql = sql & "ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & strTemp), "'", "''") & "', "

                            'sql = sql & "Observation_Value = '" & Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''") & "', "
                            '==================================================================================================
                            If isDolby(dictNVP.Item("ObservationValue" & strTemp)) Then


                                strMemo = addMemo(Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''"))
                                sql = sql & "memo = '" & strMemo & "', "
                                sql = sql & "Observation_Value = '** Report **', "

                            Else

                                sql = sql & "Observation_Value = '" & Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''") & "', "

                            End If
                            '==================================================================================================
                            sql = sql & "UnitsID = '" & Replace(dictNVP.Item("UnitsID" & strTemp), "'", "''") & "', "
                            sql = sql & "ReferenceRange = '" & Replace(dictNVP.Item("ReferenceRange" & strTemp), "'", "''") & "', "
                            sql = sql & "Observation_ResultsStatus = '" & Replace(dictNVP.Item("ObservationResultStatus" & strTemp), "'", "''") & "', "

                            sql = sql & "OBR_3 = '" & Replace(dictNVP.Item("Filler Order Number"), "'", "''") & "', "
                            'sql = sql & "PlacerOrderNumber = '" & Replace(dictNVP.Item("Placer Order Number"), "'", "''") & "', "
                            sql = sql & "OBR_4 = '" & Replace(dictNVP.Item("Universal Service Identifier Text"), "'", "''") & "', "
                            sql = sql & "AbnormalFlags = '" & Replace(dictNVP.Item("AbnormalFlags" & strTemp), "'", "''") & "', " '20140912
                            '12/14/2005 - removed replace function on varTempComment below
                            sql = sql & "nteComment = '" & varTempComment & "', "
                            sql = sql & "updated = '" & Now & "' "
                            'sql = sql & "WHERE keyInfo = " & dictNVP.Item("ObservationDateTime")
                            sql = sql & " WHERE ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & strTemp), "'", "''") & "'"
                            sql = sql & " AND PlacerOrderNumber = '" & Replace(dictNVP.Item("Placer Order Number"), "'", "''") & "' "
                            'sql = sql & " AND ObservationIdentifierID = '" & Replace(dictNVP.Item("ObservationIdentifier1" & "_" & strTemp), "'", "''") & "' "

                            updatecommand.CommandText = sql
                            myConnection.Open()
                            updatecommand.ExecuteNonQuery()
                            myConnection.Close()


                        End If 'If rsTest("Observation_ResultsStatus") = "P"

                    Else 'If updateResults Then

                        masterID = 0
                        sql = "select id from [test_master] where PlacerOrderNumber = " & dictNVP.Item("Placer Order Number") & " "
                        sql = sql & "AND patient_accountNumber = '" & extractPanum(dictNVP.Item("Patient_AccountNumber")) & "'"
                        objCommand.CommandText = sql
                        myConnection.Open()
                        dataReader = objCommand.ExecuteReader()
                        If dataReader.HasRows() Then
                            While dataReader.Read()
                                masterID = dataReader.GetInt64(0)

                            End While
                        End If
                        myConnection.Close()
                        dataReader.Close()
                        '20140327 - added memo after observation_value to fix sql error statement
                        sql = "Insert into [test_detail]"
                        sql = sql & "(masterID, ValueType, Observation_Value, memo, UnitsID, ReferenceRange,"
                        sql = sql & "Observation_ResultsStatus, KeyInfo, Patient_AccountNumber, "
                        sql = sql & "added, nteComment, PlacerOrderNumber, OBR_3, OBR_4, "
                        sql = sql & "ObservationIdentifierID, ObservationIdentifierText, AbnormalFlags, ObservationDateTime)"
                        sql = sql & "values(" & masterID
                        sql = sql & ",'" & Replace(dictNVP.Item("ValueType" & "_" & strTemp), "'", "''") & "'"


                        'sql = sql & ",'" & Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''") & "'"

                        '==================================================================================================
                        If isDolby(dictNVP.Item("ObservationValue" & strTemp)) Then

                            sql = sql & ",'**Report**'"
                            strMemo = ""
                            strMemo = strMemo & addMemo(Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''"))

                            sql = sql & ",'" & Replace(strMemo, "'", "''") & "'"
                        Else
                            sql = sql & ",'" & Replace(dictNVP.Item("ObservationValue" & strTemp), "'", "''") & "'"
                            sql = sql & ",NULL"
                        End If
                        '==================================================================================================
                        sql = sql & ",'" & Replace(dictNVP.Item("UnitsID" & strTemp), "'", "''") & "'"
                        sql = sql & ",'" & Replace(dictNVP.Item("ReferenceRange" & strTemp), "'", "''") & "'"
                        sql = sql & ",'" & dictNVP.Item("Observation_ResultStatus" & strTemp) & "'"
                        sql = sql & "," & dictNVP.Item("ObservationDateTime")
                        sql = sql & ",'" & Replace(dictNVP.Item("Patient_AccountNumber"), "'", "''") & "'"
                        sql = sql & ",'" & DateTime.Now & "'"
                        '12/14/2005 - removed replace function on varTempComment below
                        sql = sql & ",'" & varTempComment & "'"

                        sql = sql & ",'" & Replace(dictNVP.Item("Placer Order Number"), "'", "''") & "'"
                        sql = sql & ",'" & Replace(dictNVP.Item("Filler Order Number"), "'", "''") & "'"
                        sql = sql & ",'" & Replace(dictNVP.Item("Universal Service Identifier Text"), "'", "''") & "'"
                        sql = sql & ",'" & Replace(dictNVP.Item("ObservationIdentifier1" & strTemp), "'", "''") & "'" ' ObservationIdentifierID
                        sql = sql & ",'" & Replace(dictNVP.Item("ObservationIdentifier2" & strTemp), "'", "''") & "'" ' ObservationIdentifierText
                        sql = sql & ",'" & Replace(dictNVP.Item("AbnormalFlags" & strTemp), "'", "''") & "'"

                        sql = sql & ",'" & Replace(ConvertDate(dictNVP.Item("ObservationDateTime")), "'", "''") & "'"


                        sql = sql & ")"
                        updatecommand.CommandText = sql
                        myConnection.Open()
                        updatecommand.ExecuteNonQuery()
                        myConnection.Close()
                    End If 'If updateResults Then
                End If
            Next 'For i = 2 To gblOBX
        Catch ex As Exception
            globalError = True
            gblLogString = gblLogString & CStr(ex.Message)
        End Try
    End Sub
    

    Public Function ConvertDate(ByVal datedata As String) As String
        '20100402 - updated
        Dim strYear As String = ""
        Dim strMonth As String = ""
        Dim strDay As String = ""
        Dim strHour As String = ""
        Dim strMinute As String = ""

        If Len(Trim(datedata)) = 8 Then
            strYear = Mid$(datedata, 1, 4)
            strMonth = Mid$(datedata, 5, 2)
            strDay = Mid$(datedata, 7, 2)
            ConvertDate = strMonth & "/" & strDay & "/" & strYear

        ElseIf Len(Trim(datedata)) >= 12 Then
            strYear = Mid$(datedata, 1, 4)
            strMonth = Mid$(datedata, 5, 2)
            strDay = Mid$(datedata, 7, 2)
            strHour = Mid$(datedata, 9, 2)
            strMinute = Mid$(datedata, 11, 2)

            If strHour = "24" Then
                ConvertDate = strMonth & "/" & strDay & "/" & strYear
            Else
                ConvertDate = strMonth & "/" & strDay & "/" & strYear & " " & strHour & ":" & strMinute
            End If


        Else
            ConvertDate = DateTime.Now

        End If


    End Function
    
    Public Sub addDetail(ByVal dictNVP As Scripting.Dictionary, ByVal lastMasterID As Integer)
        '20131024 - add transcription processing
        Dim myConnection As New SqlConnection(connectionString)
        Dim objCommand As New SqlCommand
        Dim updatecommand As New SqlCommand
        updatecommand.Connection = myConnection
        Dim sql As String = ""
        Dim I As Integer = 0
        Dim ArrayCounter As Integer = 0
        Dim varTempComment As String = ""
        Dim strNTETemp As String = ""
        Dim strTemp As String = ""
        Dim strMemo As String = "" '20131024 - added to handle memo for Dolby Transcription
        Try
            For I = 1 To gblOBX

                If I = 1 Then 'And dictNVP.Item("ValueType_0002") <> "TX" Then
                    'put together all the nte comments
                    varTempComment = ""
                    ArrayCounter = 0

                    For ArrayCounter = 0 To UBound(theArray)
                        If theArray(ArrayCounter, 0) = 1 And theArray(ArrayCounter, 1) = 1 Then
                            '20150219 add replace function
                            varTempComment = varTempComment & Trim(Replace(dictNVP.Item("Comment"), "'", "''")) & vbCrLf
                        End If

                        If theArray(ArrayCounter, 0) = 1 And theArray(ArrayCounter, 1) > 1 Then
                            Select Case theArray(ArrayCounter, 1)
                                '12/14/2004 - remove trims and escapre apostrophies for ntecomments
                                Case Is < 10
                                    strNTETemp = "000" & theArray(ArrayCounter, 1)
                                    varTempComment = varTempComment & Replace(dictNVP.Item("Comment" & "_" & strNTETemp), "'", "''") & vbCrLf
                                Case Is > 10
                                    strNTETemp = "00" & theArray(ArrayCounter, 1)
                                    varTempComment = varTempComment & Replace(dictNVP.Item("Comment" & "_" & strNTETemp), "'", "''") & vbCrLf
                                Case Is > 100
                                    strNTETemp = "0" & theArray(ArrayCounter, 1)
                                    varTempComment = varTempComment & Replace(dictNVP.Item("Comment" & "_" & strNTETemp), "'", "''") & vbCrLf
                            End Select
                        End If

                    Next

                    '===============================================================================================================
                    strTemp = "" & I
                    If Len(strTemp) = 1 Then
                        strTemp = "000" & strTemp
                    ElseIf Len(strTemp) = 2 Then
                        strTemp = "00" & strTemp
                    ElseIf Len(strTemp) = 3 Then
                        strTemp = "0" & strTemp
                    End If

                    sql = "Insert into [test_detail]"
                    sql = sql & "(masterID, ValueType, Observation_Value, memo, UnitsID, ReferenceRange,"
                    sql = sql & "Observation_ResultsStatus, KeyInfo, Patient_AccountNumber, "
                    sql = sql & "added, nteComment, PlacerOrderNumber, OBR_3, OBR_4, "
                    sql = sql & "ObservationIdentifierID, ObservationIdentifierText, AbnormalFlags, ObservationDateTime)"
                    sql = sql & "values(" & lastMasterID
                    sql = sql & ",'" & dictNVP.Item("ValueType") & "'"
                    '============================================================================================================
                    If isDolby(Replace(dictNVP.Item("ObservationValue"), "'", "''")) Then
                        sql = sql & ",'** Transcription **'"
                        strMemo = strMemo & addMemo(Replace(dictNVP.Item("ObservationValue"), "'", "''"))
                        sql = sql & ", '" & strMemo & "'"
                    Else
                        sql = sql & ",'" & Replace(dictNVP.Item("ObservationValue"), "'", "''") & "'"
                        sql = sql & ",NULL "
                    End If
                    '============================================================================================================
                    sql = sql & ",'" & Replace(dictNVP.Item("UnitsID"), "'", "''") & "'"
                    sql = sql & ",'" & Replace(dictNVP.Item("ReferenceRange"), "'", "''") & "'"
                    sql = sql & ",'" & Replace(dictNVP.Item("ObservationResultStatus"), "'", "''") & "'"
                    sql = sql & "," & dictNVP.Item("ObservationDateTime")
                    sql = sql & ",'" & Replace(extractPanum(dictNVP.Item("Patient_AccountNumber")), "'", "''") & "'"
                    sql = sql & ",'" & DateTime.Now & "'"

                    sql = sql & ",'" & varTempComment & "'"

                    sql = sql & ",'" & Replace(dictNVP.Item("Placer Order Number"), "'", "''") & "'"
                    sql = sql & ",'" & Replace(dictNVP.Item("Filler Order Number"), "'", "''") & "'"
                    sql = sql & ",'" & Replace(dictNVP.Item("Universal Service Identifier Text"), "'", "''") & "'"
                    sql = sql & ",'" & Replace(dictNVP.Item("ObservationIdentifier1"), "'", "''") & "'" ' ObservationIdentifierID
                    sql = sql & ",'" & Replace(dictNVP.Item("ObservationIdentifier2"), "'", "''") & "'" ' ObservationIdentifierText
                    sql = sql & ",'" & Replace(dictNVP.Item("AbnormalFlags"), "'", "''") & "'"
                    sql = sql & ",'" & Replace(ConvertDate(dictNVP.Item("ObservationDateTime")), "'", "''") & "'"


                    sql = sql & ")"


                    updatecommand.CommandText = sql
                    myConnection.Open()
                    updatecommand.ExecuteNonQuery()
                    myConnection.Close()
                    '===============================================================================================================

                ElseIf I > 1 Then 'And dictNVP.Item("ValueType" & "_" & strTemp) <> "TX" Then

                    'put together all the nte comments
                    varTempComment = ""
                    ArrayCounter = 0

                    For ArrayCounter = 0 To UBound(theArray)
                        If theArray(ArrayCounter, 0) = I And theArray(ArrayCounter, 1) = 1 Then
                            varTempComment = varTempComment & Trim(dictNVP.Item("Comment")) & vbCrLf
                        End If

                        If theArray(ArrayCounter, 0) = I And theArray(ArrayCounter, 1) > 1 Then
                            Select Case theArray(ArrayCounter, 1)
                                '12/14/2004 - remove trims and escapre apostrophies for ntecomments
                                Case Is < 10
                                    strNTETemp = "000" & theArray(ArrayCounter, 1)
                                    varTempComment = varTempComment & Replace(dictNVP.Item("Comment" & "_" & strNTETemp), "'", "''") & vbCrLf
                                Case Is > 10
                                    strNTETemp = "00" & theArray(ArrayCounter, 1)
                                    varTempComment = varTempComment & Replace(dictNVP.Item("Comment" & "_" & strNTETemp), "'", "''") & vbCrLf
                                Case Is > 100
                                    strNTETemp = "0" & theArray(ArrayCounter, 1)
                                    varTempComment = varTempComment & Replace(dictNVP.Item("Comment" & "_" & strNTETemp), "'", "''") & vbCrLf
                            End Select
                        End If

                    Next


                    strTemp = "" & I
                    If Len(strTemp) = 1 Then
                        strTemp = "000" & strTemp
                    ElseIf Len(strTemp) = 2 Then
                        strTemp = "00" & strTemp
                    ElseIf Len(strTemp) = 3 Then
                        strTemp = "0" & strTemp
                    End If

                    gblSQL = "Insert into [test_detail]"
                    gblSQL = gblSQL & "(masterID, ValueType, Observation_Value, memo, UnitsID, ReferenceRange,"
                    gblSQL = gblSQL & "Observation_ResultsStatus, KeyInfo, Patient_AccountNumber, "
                    gblSQL = gblSQL & "added, nteComment, PlacerOrderNumber, OBR_3, OBR_4, "
                    gblSQL = gblSQL & "ObservationIdentifierID, ObservationIdentifierText, AbnormalFlags, ObservationDateTime)"
                    gblSQL = gblSQL & "values(" & lastMasterID
                    insertString(dictNVP.Item("ValueType" & "_" & strTemp))

                    '==================================================================================================
                    If isDolby(dictNVP.Item("ObservationValue" & "_" & strTemp)) Then
                        insertString("** Transcription **")
                        strMemo = ""
                        strMemo = strMemo & addMemo(Replace(dictNVP.Item("ObservationValue" & "_" & strTemp), "'", "''"))
                        insertString(strMemo)
                    Else
                        insertString(dictNVP.Item("ObservationValue" & "_" & strTemp))
                        insertString("")
                    End If
                    '==================================================================================================

                    insertString(dictNVP.Item("UnitsID" & "_" & strTemp))
                    insertString(dictNVP.Item("ReferenceRange" & "_" & strTemp))
                    insertString(dictNVP.Item("ObservationResultStatus" & "_" & strTemp))
                    gblSQL = gblSQL & "," & dictNVP.Item("ObservationDateTime")
                    insertString(extractPanum(dictNVP.Item("Patient_AccountNumber")))
                    gblSQL = gblSQL & ",'" & DateTime.Now & "'"

                    '12/14/2005 - removed replace function for vartempcomment below
                    insertString(varTempComment)

                    insertString(dictNVP.Item("Placer Order Number"))
                    insertString(dictNVP.Item("Filler Order Number"))
                    insertString(dictNVP.Item("Universal Service Identifier Text"))
                    insertString(dictNVP.Item("ObservationIdentifier1" & "_" & strTemp)) ' ObservationIdentifierID
                    insertString(dictNVP.Item("ObservationIdentifier2" & "_" & strTemp)) ' ObservationIdentifierText
                    insertString(dictNVP.Item("AbnormalFlags" & "_" & strTemp)) '220140912
                    insertString(ConvertDate(dictNVP.Item("ObservationDateTime")))


                    gblSQL = gblSQL & ")"

                    updatecommand.CommandText = gblSQL
                    myConnection.Open()
                    updatecommand.ExecuteNonQuery()
                    myConnection.Close()
                    '===============================================================================================================
                End If


            Next 'For i = 1 To gblOBX


            '20131024 now all obx fields have been processed, so update the memo field in the test_master table if this is a dolby record
            'If Len(strMemo) > 0 Then
            'sql = "update [test_master] set memo = '" & strMemo & "' Where ID = " & lastMasterID
            'updatecommand.CommandText = sql
            'myConnection.Open()
            ' updatecommand.ExecuteNonQuery()
            'myConnection.Close()
            ' End If

        Catch ex As Exception
            globalError = True
            gblLogString = gblLogString & CStr(ex.Message)
        End Try
    End Sub
    Public Sub insertString(ByVal theString As String)
        If theString <> "" Then
            gblSQL = gblSQL & ",'" & Replace(theString, "'", "''") & "' "
        Else
            gblSQL = gblSQL & ",NULL "
        End If

    End Sub

    Public Sub insertNumber(ByVal theString As String)

        If IsNumeric(theString) Then
            gblSQL = gblSQL & ", " & theString & " "
        Else
            gblSQL = gblSQL & ",NULL "
        End If

    End Sub
    Public Sub writeToLog2(ByVal logText As String)
        Dim myLog As New EventLog()
        Try
            ' check for the existence of the log that the user wants to create.
            ' Create the source, if it does not already exist.
            If Not EventLog.SourceExists("ITWORU") Then
                EventLog.CreateEventSource("ITWORU", "ITWORU")
            End If

            ' Create an EventLog instance and assign its source.

            myLog.Source = "ITWORU"

            ' Write an informational entry to the event log.    
            myLog.WriteEntry(logText)

        Finally
            myLog.Close()
        End Try
    End Sub
    Public Sub writeTolog(ByVal strMsg As String)
        '20140205 - use a text file to log errors instead of the event log
        Dim file As System.IO.StreamWriter
        Dim tempLogFileName As String = strLogDirectory & "ULHORUFeed_log.txt"
        file = My.Computer.FileSystem.OpenTextFileWriter(tempLogFileName, True)
        file.WriteLine(DateTime.Now & " : " & strMsg)
        file.Close()
    End Sub
    Public Function extractMrnum(ByVal mrnum As String) As String
        '20131022 - extract the mrnum by stripping off the region caracter in first position if it exists
        If IsNumeric(Mid(mrnum, 1, 1)) Then
            extractMrnum = mrnum
        Else
            extractMrnum = Mid(mrnum, 2)
        End If
    End Function
    Public Function extractPanum(ByVal panum As String) As String
        '20131022 - extract the mrnum by stripping off the region caracter in first position if it exists
        If IsNumeric(Mid(panum, 1, 1)) Then
            extractPanum = panum
        Else
            extractPanum = Mid(panum, 2)
        End If
    End Function
    Public Function isDolby(ByVal obsValue As String) As Boolean
        Dim intTildeCount As Integer = 0
        isDolby = False

        If InStr(obsValue, "~") >= 1 Then
            'this is a transcription 
            isDolby = True
        Else
            isDolby = False
        End If

    End Function

    Public Function addMemo(ByVal obsValue As String) As String
        '20131024 - build memo string from observer value
        Dim memoArray() As String
        Dim tempString As String = ""
        addMemo = ""

        memoArray = Split(obsValue, "~")
        For i As Integer = 0 To memoArray.Length - 1
            If (memoArray(i) = " " Or memoArray(i) = "") Then
                tempString = tempString & vbCrLf
            Else
                tempString = tempString & memoArray(i) & " "
            End If
        Next

        addMemo = tempString

    End Function

    Public Sub checkOrderComment(ByVal dictNVP As Scripting.Dictionary)
        Dim counter As Integer
        Dim isitOBR As Boolean = False
        Dim isitOBX As Boolean = False
        'Dim isitORC As Boolean = False
        Dim arrKeys
        Dim arrItems
        Dim strCommentSource As String = ""
        Dim strComment As String = ""
        counter = 0
        gblOrderComment = ""
        Try
            'the keys hold the left side of the equal sign i.e. the name in the NVP
            'the items hold the values on the right side of the equal sign
            arrKeys = dictNVP.Keys
            arrItems = dictNVP.Items

            For counter = 0 To dictNVP.Count - 1

                'If Left$(arrKeys(counter), 12) = "OrderControl" Then
                'isitORC = True
                'End If


                If Left$(arrKeys(counter), 3) = "OBR" Then
                    isitOBR = True

                End If

                If Left$(arrKeys(counter), 3) = "OBX" Then
                    isitOBX = True
                End If

                If Left$(arrKeys(counter), 15) = "SourceOfComment" Then
                    If isitOBR And Not isitOBX Then
                        strCommentSource = dictNVP.Item(arrKeys(counter))
                    End If
                End If

                If Left$(arrKeys(counter), 7) = "Comment" Then
                    If isitOBR And Not isitOBX Then
                        strCommentSource = dictNVP.Item(arrKeys(counter))

                        If strCommentSource <> "" Then
                            '20150218 remove repalce code - redundant
                            'gblOrderComment = gblOrderComment & Replace(strComment, "'", "''") & vbCrLf 'Replace(theString, "'", "''")
                            gblOrderComment = gblOrderComment & strComment & vbCrLf
                        Else
                            gblOrderComment = ""
                        End If
                    End If
                End If

            Next


        Catch ex As Exception
            globalError = True
            gblLogString = gblLogString & CStr(ex.Message)
        End Try
    End Sub
End Module
