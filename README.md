# ChatGPT-Excel
Chat GPT Function for Excel

Instructions:

1) Obtain an API Key from Openai.com
2) Start Excel, then open the Visual Basic Editor (alt-F11)
3) Import the Module askGPT.bas
4) Enter your API Key in the API_KEY variable.
5) Save the spreadsheet as a macro-enabled Excel workbook. Or you can unhide Personal and put it in there and save that as Macro Enabled.

To call the function in Excel, its simply =askGPT(A1)  
Where A1 is the cell containing the prompt.

For more advanced work, say checking a table for example, you could use:

=askgpt("Does this table add up: "&TEXTJOIN(",",FALSE,C29:F33))

..Where C29:F33 is the table. TEXTJOIN turns a table into a CSV string. It works. You might want to play with the output though to make it prettier.

Here's a video of it working: https://github.com/Vaughanwj/ChatGPT-Excel/commit/0022d124cdde75b48b08ff61a1e6dce822a57e62

Expect there to be a delay in copy and pasting the function. 

any questions, vaughanwj@futurewatch.ai

Please suscribe to our youtube channel! 
youtube.com/@futurewatch-ai

Thanks, have fun!

