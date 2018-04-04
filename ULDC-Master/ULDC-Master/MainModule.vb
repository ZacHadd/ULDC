Imports System.IO
Imports Microsoft.Office.Interop


'*******************************************************************'
'                  Project:  User Login Data Combiner (ULDC)        '
'                   Author:  Zachary Hadd                           '
'            Last Modified:  4/4/2018                               '
'*******************************************************************'
' Purpose:                                                          '
'    This application was created to combine two very specific files'
' together as a single source of user login data. This will allow   '
' for the data to be used for reporting to see how many users are   '
' logging in to the system and when.                                '
'*******************************************************************'
' Global Variables:                                                 '
'   strEXCELFILE - String - Holds the filepath to the excel file    '
'   strINPUTFILE - String - Holds the filepath to the txt input file'
'   strOUTPUTFILE - String - Holds the path where the output file   '
'                            will be created.                       '
'*******************************************************************'
' Classes:                                                          '
'___________________________________________________________________'
'clsUser - Represents a user. Attributes describe each user         '
'   Attributes                                                      '
'      strID ------------- Nexteer users Z-ID.                      '
'      strName ----------- Nexteer name (Last, First Middle).       '
'      strLicenseLevel --- The Teamcenter license level of the user.'
'      strCountry -------- Nexteer user's country.                  '
'   Methods                                                         '
'      PrintAttributes() - Displays user attributes.                '
'*******************************************************************'
' Methods:                                                          '
'___________________________________________________________________'
'createExcelInstance:                                               '
'      Parameters:                                                  '
'         None.                                                     '
'      Description:                                                 '
'         Sets up the excel file to be processed.                   '
'      Returns:                                                     '
'         ExcelFile as Excel.Application.Workbook                   '
'___________________________________________________________________'
'GetUserCount:                                                      '
'      Parameters:                                                  '
'         (anExcelApplication as Excel.Application)                 '
'      Description:                                                 '
'         Counts how many users are in the excel file.              '
'      Returns:                                                     '
'         Count as Integer                                          '
'___________________________________________________________________'
'FillUserDictionary:                                                '
'      Parameters:                                                  '
'         (userCount as integer, anExcelApplication as              '
'          Excel.Application, usersDictionary as SortedDictionary)  '
'      Description:                                                 '
'          Fills the userDictionary with the user information from  '
'          the user excel file.                                     '
'___________________________________________________________________'
'ProcessData:                                                       '
'      Parameters:                                                  '
'         (usersDictionary as SortedDictionary)                     '
'      Description:                                                 '
'         Processes the login data from the txt file for each user  '
'         and prints out the information one row at a time to the   '
'         output file.                                              '
'*******************************************************************'

Module MainModule

    Dim strEXCELFILE As String = "c:\CUD\Nexteer_active_PROD_02242018_tester.xlsx"
    Dim strINPUTFILE As String = "c:\CUD\logs_7DAY_TEST.txt"
    Dim strOUTPUTFILE As String = "c:\CUD\CUD_.txt"

    Sub Main()

        ' Using a function for creating an instance of and 
        ' opening an excel document. This is to clean up the code
        ' in main so it is more readable.
        Dim xApp As Excel.Application = createExcelInstance()

        ' Using a function that counts the lines in the excel document
        ' to get a count of how many users are in the document.
        Dim userCount As Integer = GetUserCount(xApp)

        ' Create a dictionary for easy referencing of known users
        ' by using thier id as the key (type string). Then filling 
        ' the dictionary using the FillUserDictionary Method.
        Dim dicUsers As New SortedDictionary(Of String, clsUser)
        FillUserDictionary(userCount, xApp, dicUsers)

        ProcessData(dicUsers)

        ' Cleaning up possible loose ends
        xApp.Quit()
        xApp = Nothing

    End Sub

    '*********************************************************************************'
    '                    IMPORTANT NOTE ON EXCEL INSTANCES                            '
    '_________________________________________________________________________________'
    '    In this function a new excel application instance is created. You'll notice  '
    ' under the task manager processes that an new EXCEL.EXE will be started, even    '
    ' if there is already an instance running. This can be changed to only allow      '
    ' one instance to run at one time, however there are some big disadvantages.      '
    '                                                                                 '
    '    The first and biggest reason to make our own instance is that if there are   '
    ' multiple workbooks open under one instance, each workbook needs to be taken     '
    ' into consideration while parcing through our desired workbook of user info.     '
    ' This means that the time to process the user info will dramatically increase    '
    ' based on the number of excel files that are openned.                            '
    '     The second reason we want to make a seperate instance is because we cannot  '
    ' quit a shared instance unless we are ok with closing every open excel sheet the '
    ' user is currently working on. That is not good. Even if you chose to leave it   '
    ' open this means that when the user closes all of the open excel files the       '
    ' instance will remain open in the background taking up Memory because we didn't  '
    ' close it.                                                                       '
    '                                                                                 '
    '     DO NOT USE AN EXISTING EXCEL.APPLICATION OBJECT TO OPEN THE INPUT DATA      '
    '                                                                                 '
    '         A new instance was created for a reason so don't change it.             '
    '                                                                                 '
    '*********************************************************************************'
    Function createExcelInstance()
        Dim app As New Excel.Application

        Try
            app.Workbooks.Open(strEXCELFILE)
        Catch ex As Exception
            Debug.WriteLine(vbCrLf + vbCrLf + "ERROR" + vbCrLf)
            Debug.WriteLine(strEXCELFILE + " does not exist or could not be opened." +
                            vbCrLf + vbCrLf)
            Debug.WriteLine(ex)
            app.Quit()
            app = Nothing
            Environment.Exit(0)
        End Try

        Return app
    End Function

    Function GetUserCount(ByVal excelSheet As Excel.Application)
        Debug.WriteLine(vbCrLf + "Running GetUserCount")

        Dim usersCount As Integer = 0
        ' Need to skip the first row because its a header row.
        Dim rowCount As Integer = 2

        'While there are more users left add on to the user count
        While Not excelSheet.Cells(rowCount, 2).text = ""
            usersCount += 1
            rowCount += 1
        End While

        Debug.WriteLine(vbCrLf + "Number of users in excel doc: " + CStr(usersCount))
        Return usersCount
    End Function

    Sub FillUserDictionary(ByVal userCount As Integer, ByVal excelSheet As Excel.Application, ByRef NewUser As SortedDictionary(Of String, clsUser))
        Debug.WriteLine(vbCrLf + "Running FillUserDictionary")
        For row = 2 To userCount + 1
            Dim user As New clsUser(excelSheet.Cells(row, 2).value(), excelSheet.Cells(row, 3).value(), excelSheet.Cells(row, 5).value(), excelSheet.Cells(row, 6).value())
            NewUser.Add(user.ID, user)
            Select Case row
                Case CInt(userCount / 10)
                    Debug.Write("10%  ")
                Case CInt(userCount / 4)
                    Debug.Write("25%   ")
                Case CInt(userCount / 2)
                    Debug.Write("50%   ")
                Case CInt(userCount * 0.7)
                    Debug.Write("75%  ")
                Case CInt(userCount * 0.9)
                    Debug.Write("90%")
            End Select
        Next
        Debug.Write("100%" + vbCrLf)
        Debug.WriteLine("FillUserDicitonary Complete.")

    End Sub

    Sub ProcessData(ByVal Users As SortedDictionary(Of String, clsUser))
        Dim inFile As StreamReader = Nothing
        Dim outFile As StreamWriter = Nothing

        Try
            inFile = New StreamReader(strINPUTFILE)
            Debug.WriteLine(vbCrLf + "Successfully opened stream reader.")
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
            MessageBox.Show("No logs file in CUD folder at location:" &
                            vbCrLf & vbCrLf & strINPUTFILE)
        End Try
        Try
            outFile = New IO.StreamWriter(strOUTPUTFILE.Insert(strOUTPUTFILE.IndexOf("_") + 1, DateTime.Now().ToString("yyyyMMdd'_'hhmmss")))
            Debug.WriteLine(vbCrLf + "Successfully created output file.")
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
        End Try

        Dim strTemp As String
        Dim strLoginDateTime As String = ""

        Debug.WriteLine(vbCrLf + "Starting ProcessData")
        'skip 3 blank lines
        For i = 1 To 3
            inFile.ReadLine()
        Next

        While Not inFile.EndOfStream
            strTemp = CStr(inFile.ReadLine())

            If Not (strTemp.StartsWith("Mon") Or strTemp.StartsWith("Tue") Or strTemp.StartsWith("Wed") Or strTemp.StartsWith("Thu") Or strTemp.StartsWith("Fri") Or strTemp.StartsWith("Sat") Or strTemp.StartsWith("Sun")) Then
                strTemp = strTemp.Substring(4, 6)
                If Not (strTemp = "infodb" Or strTemp = "dcprox" Or strTemp = "projpr" Or strTemp = "stcadm") Then
                    If Users.ContainsKey(strTemp) Then
                        outFile.WriteLine(Users(strTemp).PrintAttributes + strLoginDateTime)
                    End If
                End If
            Else
                strLoginDateTime = strTemp.Substring(4)
                strTemp = inFile.ReadLine()
                strLoginDateTime += " " & strTemp
                'skip unnecessary lines
                For i = 1 To 12
                    inFile.ReadLine()
                Next
            End If
        End While
        Debug.WriteLine(vbCrLf + "Successfully output data to " + strOUTPUTFILE)
    End Sub

End Module
