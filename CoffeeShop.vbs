Option Explicit
'
' CoffeeShop.vbs (VBScript) Exercise by Mark Fairpo.
'
' Run: CoffeeShop.vbs [FileType1.csv] [FileType2.csv]
'
' Notes: Files can be selected in any order as detects the file type.
'
' Assumes: For "most popular blend" as each Drink has an unknown measure of Roast, that a "blend" is a unique pairing.
' Assumes: "what percentage of customers choose from the extras" is a customer per Drink, as customers may have many transaction rows.
' Assumes: "cocnut" is a typo "coconut", but this may also be a unique but miscoded till button (see Global Constants).
' Assumes: Own code desired. Bubble-sorting Dictionary Objects is sourceable from public domain (e.g. Rhino Developer).
'
Const CSV_DELIM = ",", PAIR_DELIM = "/", SUBST_DELIM = "|", SUBST_PAIRS = "null/None|, /,|cocnut/coconut"
Const FILES_DEFAULT = "Coffee shop price list.csv|coffee shop Monday sales.csv", QUIT_ON_ERROR = False
Dim objFSO, g_ScriptDir, dicSales, dicExtras, dicPrice, colArgs
Dim g_SumRows, g_SumDrinks, g_SumSales, g_Msg
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set dicSales = CreateObject("Scripting.Dictionary")
Set dicExtras = CreateObject("Scripting.Dictionary")
Set dicPrice = CreateObject("Scripting.Dictionary")
Set colArgs = WScript.Arguments.Unnamed
Main
WScript.Echo g_Msg

Sub Main()
    Dim strKey, strDrink, intErr, strFilename
    g_ScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
    ' Process any command-line arguments
    For Each strFilename In colArgs
        intErr = intProcessCSV(strFilename)
    Next
    ' Show file open dialogue while either Dictionary Object is empty
    Do While dicSales.Count = 0 Or dicPrice.Count = 0
        strFilename = strSelectFile
        If strFilename <> "" Then
            intErr = intProcessCSV(strFilename)
        End If
    Loop
    Output "BLENDS:"
    For Each strKey In dicSales.Keys
        ' Each dictionary Key is formatted as Drink/Blend with its Item as Quantity
        strDrink = Split(strKey, PAIR_DELIM)(0)
        If dicPrice.Exists(strDrink) Then
            ' Running calculation of Sales total and output blends per popularity
            g_SumSales = g_SumSales + (dicSales(strKey) * dicPrice(strDrink))
            Output strKey & vbTab & dicSales(strKey) & "x" & vbTab & FormatPercent(dicSales(strKey) / g_SumDrinks)
        Else
            ErrHandler "Price was not found: " & strDrink
        End If
    Next
    Output ""
    Output "SALES:" & vbTab & g_SumDrinks & "x Drinks" & vbTab & FormatCurrency(g_SumSales)
    Output ""
    Output "EXTRAS PER DRINK:"
    For Each strKey In dicExtras.Keys
        Output strKey & vbTab & dicExtras(strKey) & "x" & vbTab & FormatPercent(dicExtras(strKey) / g_SumDrinks)
    Next
End Sub

Function intProcessCSV(strFilename)
' About: Detects and processes CSV file based on number of columns
' Returns: 0 or Line Number of error
    Dim intLineNum, strLine, arrReadAll, arrCols, intHeaderCols, strKey1, strKey2, vNumber, intSadLine
    ' Add default path of g_ScriptDir
    intSadLine = 0
    arrReadAll = arrFileToArray(strFilename)
    If IsArray(arrReadAll) Then
        For Each strLine In arrReadAll
            intLineNum = intLineNum + 1
            arrCols = Split(strLine, CSV_DELIM)
            If intLineNum = 1 Then
                intHeaderCols = UBound(arrCols)
                If intHeaderCols = 1 Then
                    ' Logic for Prices File
                    dicPrice.RemoveAll
                ElseIf intHeaderCols = 3 Then
                    ' Logic for Till File
                    dicSales.RemoveAll
                    dicExtras.RemoveAll
                    g_SumDrinks = 0
                    g_SumRows = 0
                End If
            Else
                If UBound(arrCols) = intHeaderCols Then
                    If intHeaderCols = 1 Then
                        ' Logic for Prices File
                        vNumber = arrCols(1)
                        ' Price column is floating-point Pounds without a mandatory decimal point
                        If IsNumeric(vNumber) Then
                            vNumber = CDbl(vNumber)
                            strKey1 = arrCols(0)
                            If Not dicPrice.Exists(strKey1) Then
                                dicPrice(strKey1) = vNumber
                            Else
                                intSadLine = intLineNum
                            End If
                        Else
                            intSadLine = intLineNum
                        End If
                    ElseIf intHeaderCols = 3 Then
                        ' Logic for Till File
                        vNumber = arrCols(2)
                        ' Quantity is an integer
                        If IsNumeric(vNumber) Then
                            vNumber = CLng(vNumber)
                            ' Each dictionary Key is formatted as Drink/Blend with its Item as Quantity
                            strKey1 = arrCols(0) & PAIR_DELIM & arrCols(1)
                            ' Extras are stored in a seperate dictionary
                            strKey2 = arrCols(3)
                            If Not dicSales.Exists(strKey1) Then
                                dicSales(strKey1) = vNumber
                            Else
                                dicSales(strKey1) = dicSales(strKey1) + vNumber
                            End If
                            If Not dicExtras.Exists(strKey2) Then
                                dicExtras(strKey2) = vNumber
                            Else
                                dicExtras(strKey2) = dicExtras(strKey2) + vNumber
                            End If
                            ' Running calculation of total Drinks and data Rows
                            g_SumDrinks = g_SumDrinks + vNumber
                            g_SumRows = intLineNum - 1
                        Else
                            intSadLine = intLineNum
                        End If
                    Else
                        intSadLine = intLineNum
                    End If
                Else
                    intSadLine = intLineNum
                End If
            End If
        Next
    Else
        intSadLine = 1
    End If
    If intSadLine > 0 Then
        ErrHandler "Problem with CSV file """ & Mid(strFilename, InStrRev(strFilename, "\") + 1) & """ on line " & intSadLine
    End If
    intProcessCSV = intSadLine
End Function

Function arrFileToArray(strFilename)
' About: Returns a raw text file as an array (IsArray); or vbEmpty if errors
    Const ForReading = 1
    Dim objFile, vData, arrLines, strPair
    If InStr(strFilename, "\") = 0 Then
        strFilename = objFSO.BuildPath(g_ScriptDir, strFilename)
    End If
    If objFSO.FileExists(strFilename) Then
        If objFSO.GetFile(strFilename).Size > 0 Then
            Set objFile = objFSO.OpenTextFile(strFilename, ForReading)
            vData = objFile.ReadAll
            objFile.Close
            ' Clean text with replacements
            For Each strPair In Split(SUBST_PAIRS, SUBST_DELIM)
                vData = Replace(vData, Split(strPair, PAIR_DELIM)(0), Split(strPair, PAIR_DELIM)(1))
            Next
            ' Convert text into an array
            arrLines = Split(vData, vbNewLine)
            ' Trim whitespace from the tail
            Do While UBound(arrLines) > 0 And Trim(arrLines(UBound(arrLines))) = ""
                ReDim Preserve arrLines(UBound(arrLines) - 1)
            Loop
            arrFileToArray = arrLines
        End If
    End If
End Function

Sub Output(strMsg)
' About: Standardises console output by building a global variable
    g_Msg = IIf(g_Msg = "", strMsg, g_Msg & vbNewLine & strMsg)
End Sub

Sub ErrHandler(strMsg)
' About: Standardises error output.
    WScript.Echo "Error - " & strMsg
    If QUIT_ON_ERROR Then WScript.Quit 1
End Sub

Function IIf(blnExpression, vTruePart, vFalsePart)
' About: If expression is true returns TruePart or else FalsePart
    IIf = vFalsePart: If blnExpression Then IIf = vTruePart
End Function

Function strSelectFile()
    ' File Browser via HTA
    ' Author:   Rudi Degrande, modifications by Denis St-Pierre and Rob van der Woude
    ' Features: Works in Windows Vista and up (Should also work in XP).
    '           Fairly fast.
    '           All native code/controls (No 3rd party DLL/ XP DLL).
    ' Caveats:  Cannot define default starting folder.
    '           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
    '           Dialog title says "Choose file to upload".
    ' Source:   https://social.technet.microsoft.com/Forums/scriptcenter/en-US/a3b358e8-15ae-4ba3-bca5-ec349df65ef6/windows7-vbscript-open-file-dialog-box-fakepath?forum=ITCG
    Dim objExec, strMSHTA, wshShell
    strSelectFile = ""
    strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
             & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
    Set wshShell = CreateObject("WScript.Shell")
    Set objExec = wshShell.Exec(strMSHTA)
    strSelectFile = objExec.StdOut.ReadLine()
End Function
