'Code Taken From https://www.wiseowl.co.uk/blog/s393/scrape-website-html.htm

Sub ImportStackOverflowData()
    'to refer to the running copy of Internet Explorer
    Dim ie As InternetExplorer
    'to refer to the HTML document returned
    Dim html As HTMLDocument
    'open Internet Explorer in memory, and go to website
    Set ie = New InternetExplorer
    ie.Visible = False
    ie.navigate "http://stackoverflow.com/"
    'Wait until IE is done loading page
    Do While ie.readyState <> READYSTATE_COMPLETE
        Application.StatusBar = "Trying to go to StackOverflow ..."
        DoEvents
    Loop
    'show text of HTML document returned
    Set html = ie.document
    'MsgBox html.DocumentElement.innerHTML
    'close down IE and reset status bar
    Set ie = Nothing
    Application.StatusBar = ""
    'clear old data out and put titles in
    Cells.Clear
    'put heading across the top of row 3
    Range("A3").Value = "Question id"
    Range("B3").Value = "Votes"
    Range("C3").Value = "Views"
    Range("D3").Value = "Person"
    refHTML html
End Sub

Sub refHTML(html As HTMLDocument)
    Dim QuestionList As IHTMLElement
    Dim Questions As IHTMLElementCollection
    Dim Question As IHTMLElement
    Dim RowNumber As Long
    Dim QuestionId As String
    Dim QuestionFields As IHTMLElementCollection
    Dim QuestionField As IHTMLElement
    Dim votes As String
    Dim views As String
    Dim QuestionFieldLinks As IHTMLElementCollection


    Set QuestionList = html.getElementById("question-mini-list")
    Set Questions = QuestionList.Children
    
    RowNumber = 4
    For Each Question In Questions
    'if this is the tag containing the question details, process it
    If Question.className = "question-summary narrow" Then
        'first get and store the question id in first column
        QuestionId = Replace(Question.ID, "question-summary-", "")
        Cells(RowNumber, 1).Value = CLng(QuestionId)
        'get a list of all of the parts of this question,
        'and loop over them
        Set QuestionFields = Question.all
        For Each QuestionField In QuestionFields
            'if this is the question's votes, store it (get rid of any surrounding text)
            If QuestionField.className = "votes" Then
                votes = Replace(QuestionField.innerText, "votes", "")
                votes = Replace(votes, "vote", "")
                Cells(RowNumber, 2).Value = Trim(votes)
            End If
            'likewise for views (getting rid of any text)
            If QuestionField.className = "views" Then
                views = QuestionField.innerText
                views = Replace(views, "views", "")
                views = Replace(views, "view", "")
                Cells(RowNumber, 3).Value = Trim(views)
            End If
            'if this is the bit where author's name is ...
            If QuestionField.className = "started" Then
                'get a list of all elements within, and store the
                'text in the second one
                Set QuestionFieldLinks = QuestionField.all
                Cells(RowNumber, 4).Value = QuestionFieldLinks(2).innerHTML
                Cells(RowNumber, 5).Value = QuestionFieldLinks(1).innerHTML
            End If
    Next QuestionField
    'go on to next row of worksheet
    RowNumber = RowNumber + 1
    End If
    Next
Set html = Nothing
End Sub
