Attribute VB_Name = "askGPT"
Function askGPT(prompt_text As String) As String
    
    ' Copyright (C) Vaughan Wynne-Jones https://futurewatch.ai
    ' MIT License in License.txt
    ' Basically you can use it if you give me credit.
    ' You can help out by subscribing to @futurewatch-ai on Youtube
    ' https://www.youtube.com/@futurewatch-ai
    
    
    Dim API_KEY As String
    
    'Enter your API Key here:
    API_KEY = "YOUR_API_KEY_HERE"
    
        
    Dim request As Object
    Dim response As String
    Dim endPoint As String
    Dim maxTokens As Integer
    Dim SystemPrompt As String
    Dim Examples As String
    Dim Instructions As String
    
    SystemPrompt = "Digital Excel Assistant"
    Examples = "Return answers that are succinct but highly accurate"
    prompt_text = Replace(prompt_text, Chr(34), "")
    
    endPoint = "https://api.openai.com/v1/chat/completions"
    Model = "gpt-4-0314"
    maxTokens = 1024
    
    Set request = CreateObject("MSXML2.XMLHTTP")
    request.Open "POST", endPoint, False
    request.setRequestHeader "Content-Type", "application/json"
    request.setRequestHeader "Authorization", "Bearer " & API_KEY
    request.send "{" _
             & Chr(34) & "messages" & Chr(34) & ": [ {" _
             & Chr(34) & "role" & Chr(34) & ": " & Chr(34) & "system" & Chr(34) & ", " _
             & Chr(34) & "content" & Chr(34) & ": " & Chr(34) & SystemPrompt & Chr(34) _
             & "}, {" _
             & Chr(34) & "role" & Chr(34) & ": " & Chr(34) & "assistant" & Chr(34) & ", " _
             & Chr(34) & "content" & Chr(34) & ": " & Chr(34) & Examples & Chr(34) _
             & "},{ " _
             & Chr(34) & "role" & Chr(34) & ": " & Chr(34) & "user" & Chr(34) & ", " _
             & Chr(34) & "content" & Chr(34) & ": " & Chr(34) & prompt_text & Chr(34) _
             & "} ]," _
             & Chr(34) & "model" & Chr(34) & ": " _
             & Chr(34) & Model & Chr(34) & ", " _
             & Chr(34) & "max_tokens" & Chr(34) & ": " & maxTokens _
             & "}"
             
    
    response = request.responseText
    Dim answer As String
    answer = Mid(response, InStr(response, "content"":""") + Len("content"":"""), InStr(InStr(response, "content"":""") + Len("content"":"""), response, """") - InStr(response, "content"":""") - Len("content"":"""))
    askGPT = answer
    
    
End Function

