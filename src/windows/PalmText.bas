Attribute VB_Name = "PalmText"
'
'
' PaLM Text - Version 0.0.1
'
'
Option Explicit

Sub PalmText()


  If Selection.Type = wdSelectionIP Then
    Exit Sub
  End If
  
  If Selection.Text = ChrW$(13) Then
    Exit Sub
  End If

  

  Dim apiKey As String
  apiKey = Environ("GOOGLE_PALM_API_KEY")
  
  Dim apiEndpointURL As String
  apiEndpointURL = "https://generativelanguage.googleapis.com/v1beta2/"
  
  Dim modelName As String
  modelName = "models/text-bison-001"
  
  Dim generationMethod As String
  generationMethod = "generateText"
  
  Dim textPrompt As String
  textPrompt = Replace(Selection, ChrW$(13), "")
  
  Dim fullApiEndpointUrl As String
  fullApiEndpointUrl = apiEndpointURL & modelName & ":" & generationMethod & "?key=" & apiKey

  
  Dim jsonData As String
  jsonData = "{""prompt"": {""text"": """ & textPrompt & """}}"

  
  Dim curlHttp As Object
  Set curlHttp = CreateObject("MSXML2.serverXMLHTTP")
  
  

  With curlHttp
    .Open "POST", fullApiEndpointUrl, False
    .SetRequestHeader "Content-type", "application/json"
    .Send (jsonData)
    
    Dim status As Integer
    status = .status
    Dim statusText As String
    statusText = .statusText
    

    If status <> 200 Then
      MsgBox Prompt:="We have encountered the following error: " & status & " " & statusText
      Exit Sub
    End If
    
    Dim response As String
    response = .ResponseText
    
  End With
  

  
  

  Dim startPosition As Integer
  startPosition = InStr(1, response, Chr(34) & "output" & Chr(34)) + 11

  
  Dim endPosition As Integer
  endPosition = InStr(1, response, Chr(34) & "safetyRatings" & Chr(34)) - 9

  
  Dim textResponseLength As Integer
  textResponseLength = endPosition - startPosition
    
  Dim textResponseOutput As String
  textResponseOutput = Mid(response, startPosition, textResponseLength)

  
  

  Dim textResponseOutputFormatted As String
  textResponseOutputFormatted = Replace(textResponseOutput, "\n", vbCrLf)
 
  
  With Selection
    .Collapse Direction:=wdCollapseEnd
    .InsertAfter vbCr & textResponseOutputFormatted
    .Font.Name = "Roboto"
    .Font.Size = 9
    .Font.ColorIndex = wdTeal
    .Paragraphs.Alignment = wdAlignParagraphJustify
    .InsertAfter vbCr
    .Collapse Direction:=wdCollapseEnd
  End With
  
  
  Set curlHttp = Nothing


End Sub
