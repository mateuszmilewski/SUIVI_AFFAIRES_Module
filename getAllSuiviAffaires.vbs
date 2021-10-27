Option Explicit

'The MIT License (MIT)
'
'Copyright (c) 2021 FORREST
' Mateusz Milewski mateusz.milewski@mpsa.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'

' the module from mixSoreppYDi3

' new on 2021-09-21 / 2021-09-20

' this module going through folders of shared drive moreless like it was during wizard and collectors time of LESS family tools
' but I will make very new implementation w/o any extra class handlers - trying to handle everyting within 1 module (this one)
' at the end - this is not so complicated - and big advantage of this solution that at least debugging will work on whole scope


Public Sub getAllSuiviAffaires()


    Application.Calculation = xlCalculationManual

 
    Dim sharedDriveRootLink As String, mainPrefixFolder As String, sAff1 As String, sAffFlex As String, filenamePattern As String, toPrecisePattern As String
    sharedDriveRootLink = "//yvshn002.inetpsa.com/_LSI/Applilsi/HEBERGE/bdd/APPLIS_PRY/Sechel/"
    mainPrefixFolder = "VS"
    sAff1 = "SUIVI AFFAIRES"
    sAffFlex = "su*aff*"
    filenamePattern = "SuiviAffaires*xlsx"
    toPrecisePattern = "SuiviAffaires??.xlsx"
    
    Dim collectionOfTheExcelFiles As New Dictionary
    
    
    initCfg sharedDriveRootLink, mainPrefixFolder, sAffFlex, filenamePattern, collectionOfTheExcelFiles
    makeLoopForGatheringPotentialExcelFiles collectionOfTheExcelFiles
    
    
    Application.Calculation = xlCalculationAutomatic
    
    
    MsgBox "ready!"
End Sub


Private Sub initCfg(link As String, vs As String, ptrn As String, filenamePattern As String, ByRef d As Dictionary)

    Dim fs As Variant
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim fld As Variant, subFld As Variant, subSubFld As Variant, refForSuiviAff As Variant, tmpFld As Variant
    Set refForSuiviAff = Nothing
    Set fld = fs.GetFolder(link)
    
    For Each subFld In fld.SubFolders
        
        
        
        If subFld.Name Like vs & " *" Then
        
            Debug.Print subFld.Name
        
            ' we are going inside
            For Each tmpFld In subFld.SubFolders
                
                If UCase(tmpFld.Name) Like UCase(ptrn) Then
                    Set refForSuiviAff = tmpFld
                    Exit For
                End If
            Next tmpFld
            
            If Not refForSuiviAff Is Nothing Then
                ' 2ns step to gather exactly the file!
                ' ----------------------------------------------------
                
                Dim fajl As Variant
                For Each fajl In refForSuiviAff.Files
                    
                    
                    
                    If UCase(fajl.Name) Like UCase(filenamePattern) Then
                        Debug.Print fajl.Path & " " & fajl.Name
                        
                        If d.Exists(Trim(fajl.Name)) Then
                            recurWithAdding_I d, fajl, Trim(fajl.Name) & "I"
                        Else
                            d.Add Trim(fajl.Name), fajl ' path having already important info for opening excel file
                        End If
                        ' Exit For
                    End If
                Next fajl
                
                ' ----------------------------------------------------
            End If
            
        End If
    Next subFld
    
    
End Sub

Private Sub recurWithAdding_I(d As Dictionary, f As Variant, inm As String)
    
    If d.Exists(Trim(inm)) Then
        recurWithAdding_I d, f, Trim(inm) & "I"
    Else
        Debug.Print "recur success with " & Trim(inm)
        d.Add Trim(inm), f
    End If
End Sub

Private Sub makeLoopForGatheringPotentialExcelFiles(ByRef d As Dictionary)

    Dim tmpWrk As Workbook
    Dim tmpSh As Worksheet
    
    Dim outWrk As Workbook, outSh As Worksheet, allInOneSh As Worksheet
    Set outWrk = Workbooks.Add
    Set outSh = outWrk.Sheets.Add
    Set allInOneSh = outWrk.Sheets.Add
    outSh.Cells(1, 1).Value = "COLLECTED FILES"
    outSh.Cells(1, 2).Value = "VALID"
    outSh.Cells(1, 3).Value = "MATCH"
    outSh.Cells(1, 4).Value = "SOURCE"
    
    Dim wiersz As Long
    wiersz = 2
    

    If d.Count > 0 Then
        Dim kij As Variant
        
        For Each kij In d.Keys
            Set tmpWrk = Nothing
            Set tmpWrk = Workbooks.Open(d(kij).Path, False, True)
            
            Do
                DoEvents
            Loop While tmpWrk Is Nothing
            
            
            ' -----
            DoEvents
            
            Debug.Print " tmpWrk.ActiveSheet.Cells(3, 2).Value == " & tmpWrk.ActiveSheet.Cells(3, 2).Value
            
            
            Set tmpSh = tmpWrk.ActiveSheet
            
            If Trim(tmpSh.Cells(3, 2).Value) = "Affaire" Then
                outSh.Cells(wiersz, 2).Value = "OK"
            Else
                outSh.Cells(wiersz, 2).Value = "NOK"
            End If
            
            Application.DisplayAlerts = True
            Debug.Print tmpWrk.Name
            tmpSh.Copy outSh
            Application.DisplayAlerts = False
            
            outSh.Cells(wiersz, 1).Value = Trim(kij)
            outSh.Cells(wiersz, 3).Value = outWrk.ActiveSheet.Name
            outSh.Cells(wiersz, 4).Value = Trim(d(kij).Path)
            wiersz = wiersz + 1
            
            Set tmpSh = Nothing
            tmpWrk.Close False
            
            
            
            
        Next kij
    End If

End Sub
