Attribute VB_Name = "basAlgorithm"
Option Explicit

Dim lValue As Long      'Key ASCII VALUE
Dim bCount As Byte
Dim bSeedCount As Byte  'Seed ASCII VALUE
Dim lSeed As Long
Dim Key As String       'String to hold the key

Public Function StringLength(sText As String)
    Dim x As Integer
    Dim y As Integer
    
    x = Len(sText)
    y = x Mod 2
    If x < 10 Then
        If y = 0 Then
            'If length of string is even then use this key.
            Key = "#%4&09kYxK"
        Else
            'else use this string if the length of the string is odd.
            Key = ",./8l;lejn"
        End If
    ElseIf x > 9 And x < 20 Then
        If y = 0 Then
            'If length of string is even then use this key.
            Key = "a5d87b42a98a130a.+*&"
        Else
            'else use this string if the length of the string is odd.
            Key = "bnkieytk89()73j0-DxS"
        End If
    ElseIf x > 19 And x < 39 Then
        If y = 0 Then
            'If length of string is even then use this key.
            Key = "96483lp03oKIJD)*(&#dZ34,]{[\|`!~(&$)N3"
        Else
            'else use this string if the length of the string is odd.
            Key = "&#dZ24,]{p0xoK48`!3thlN35d8\|a.+*~lp07"
        End If
    Else
        If y = 0 Then
            'If length of string is even then use this key.
            Key = ",.L;'a*xa32'\][=0|*_&#n,./zkph-31DasdesDFNnIO.,nzse439821B"
        Else
            'else use this string if the length of the string is odd.
            Key = ",sDFNnIO.,n./zL;'a_&#nkph-3*ezse1Dasd*xa32'\][=0B,.|439821"
        End If
    End If
    
End Function

Public Function GetKeyValue()
    Dim sKey As String
    Dim lKeyLen As Long
    
    lKeyLen = Len(Key)
    
    If bCount = 0 Then
        bCount = lKeyLen
    End If
    
    sKey = Right$(Key, bCount)
    lValue = Asc(sKey)
    bCount = bCount - 1
End Function

Public Function GetSeedValue(lTitleLength As Long)
    
    Dim sSeed As String
    Dim lSeedLen As Long
    
    'Compare length of title to get the seed.
    Select Case lTitleLength
        Case lTitleLength = 1 To 10
            lTitleLength = 12
        Case lTitleLength = 11 To 20
            lTitleLength = 22
        Case lTitleLength = 21 To 30
            lTitleLength = 61
        Case Else
            lTitleLength = 59
    End Select
    
    lSeedLen = Len(lTitleLength)

    If bSeedCount = 0 Then
        bSeedCount = lSeedLen
    End If
    
    sSeed = Right$(lTitleLength, bSeedCount)
    lSeed = Asc(sSeed)
    
    bSeedCount = bSeedCount - 1
End Function

Public Function MyEncrypt(sText As String, sTitle As String, _
    Encrypt As Boolean)
Attribute MyEncrypt.VB_Description = "This will allow you to encrypt / decrypt string values"
    
    StringLength sText
    
    Dim lTmp As Long
    Dim sTmp As String
    Dim sEncrypt As String
    Dim lLen As Long
    Dim lTitleLength As Long
    
    'Get the length of the string
    lLen = Len(sText)
    lTitleLength = Len(sTitle)
    sTmp = sText
    If Encrypt = True Then
        bCount = 0
        bSeedCount = 0
        Do Until lLen = 0
            GetKeyValue
            GetSeedValue (lTitleLength)
            
            lTmp = Asc(sTmp)
            lTmp = lTmp + lValue - lSeed
            If lTmp > 255 Then
                lTmp = lTmp - 256
            End If
            sText = Chr(lTmp)
            
            lLen = lLen - 1
            sTmp = Right$(strToConvert, lLen)
            sEncrypt = sEncrypt & sText
        Loop
    Else
        bCount = 0
        bSeedCount = 0
        Do Until lLen = 0
            GetKeyValue
            GetSeedValue (lTitleLength)
            
            lTmp = Asc(sTmp)
            lTmp = lTmp - lValue + lSeed
                If lTmp < 0 Then
                    lTmp = lTmp + 256
                End If
            sTmp = Chr(lTmp)
            sEncrypt = sEncrypt & sTmp
            lLen = lLen - 1
            sTmp = Right$(strToConvert, lLen)
        Loop
    End If
    strEnDecrypted = sEncrypt
End Function
