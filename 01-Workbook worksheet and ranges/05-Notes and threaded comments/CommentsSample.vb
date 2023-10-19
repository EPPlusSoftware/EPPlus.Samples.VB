Imports OfficeOpenXml
Imports OfficeOpenXml.ThreadedComments
Imports System
Imports System.Drawing

Namespace EPPlusSamples.WorkbookWorksheetAndRanges
    Public Module CommentsSample
        Public Sub Run()
            Console.WriteLine("Running sample 1.5-Comments/Notes and Threaded Comments")
            Using package = New ExcelPackage()
                ' Comments/Notes
                Dim sheet1 = package.Workbook.Worksheets.Add("Comments")
                AddComments(sheet1)
                ' Threaded comments
                Dim sheet2 = package.Workbook.Worksheets.Add("ThreadedComments")
                AddAndReadThreadedComments(sheet2)
                package.SaveAs(FileUtil.GetCleanFileInfo("1.05-Comments.xlsx"))
            End Using
            Console.WriteLine("Sample 1.5 created {0}", FileUtil.OutputDir.Name)
            Console.WriteLine()
        End Sub

        Private Sub AddComments(ByVal ws As ExcelWorksheet)
            Console.WriteLine("Sample 1.5 - Comment/Note")
            'Add Comments using the range class
            Dim comment = ws.Cells("A3").AddComment("Jan Källman:" & vbCrLf, "JK")
            comment.Font.Bold = True
            Dim rt = comment.RichText.Add("This column contains the extensions.")
            rt.Bold = False
            comment.AutoFit = True

            'Add a comment using the Comment collection
            comment = ws.Comments.Add(ws.Cells("B3"), "This column contains the size of the files.", "JK")
            'This sets the size and position. (The position is only when the comment is visible)
            comment.From.Column = 7
            comment.From.Row = 3
            comment.To.Column = 16
            comment.To.Row = 8
            comment.BackgroundColor = Color.White
            comment.RichText.Add(vbCrLf & "To format the numbers use the Numberformat-property like:" & vbCrLf)

            ws.Cells("B3:B42").Style.Numberformat.Format = "#,##0"

            'Format the code using the RichText Collection
            Dim rc = comment.RichText.Add("//Format the Size and Count column" & vbCrLf)
            rc.FontName = "Courier New"
            rc.Color = Color.FromArgb(0, 128, 0)
            rc = comment.RichText.Add("ws.Cells[")
            rc.Color = Color.Black
            rc = comment.RichText.Add("""B3:B42""")
            rc.Color = Color.FromArgb(123, 21, 21)
            rc = comment.RichText.Add("].Style.Numberformat.Format = ")
            rc.Color = Color.Black
            rc = comment.RichText.Add("""#,##0""")
            rc.Color = Color.FromArgb(123, 21, 21)
            rc = comment.RichText.Add(";")
            rc.Color = Color.Black
            Console.WriteLine("Comment added")
            Console.WriteLine()
        End Sub

        Private Sub AddAndReadThreadedComments(ByVal sheet As ExcelWorksheet)
            Dim persons = sheet.ThreadedComments.Persons
            ' Add a threaded comment author
            Dim user1 = persons.Add("Ernest Peter Plus")



            ' add a threaded comment to cell A1
            Dim thread = sheet.Cells("A1").AddThreadedComment()
            thread.AddComment(user1.Id, "My first comment")
            ' threaded comments can also be added via the worksheet:
            thread.AddComment(user1.Id, "My second comment")

            ' A workbook might have been opened by previous users that you will find in the ThreadedComments collection, could be from the AD and/or Office365.
            ' Let's add another fictive user using the user id format of Office365.
            Dim user2 = persons.Add("John Doe", "S::john.doe@somecompany.com::e3e726c6-1401-473b-bc95-cb3e1c892d99", IdentityProvider.Office365)

            ' The Thread.Sleep(50) statements below is just to avoid that comments get the same timestamp when this sample runs

            Threading.Thread.Sleep(50)
            ' now we can add comments with mentions
            thread.AddComment(user2.Id, "Really great comments there, {0}", user1)
            Threading.Thread.Sleep(50)
            thread.AddComment(user1.Id, "Many thanks {0}!", user2)

            ' A third person joins
            Dim user3 = persons.Add("IT Support")
            ' you can add multiple mentions in one comment like this
            Threading.Thread.Sleep(50)
            thread.AddComment(user3.Id, "Hello {0} and {1}, how can I help?", user1, user2)

            Console.WriteLine("*** reading threaded comments ***")
            ' Read threaded comments in a cell:
            For Each comment In sheet.Cells("A1").ThreadedComment.Comments
                Dim author = persons(comment.PersonId)
                Console.WriteLine("{0} wrote at {1}", author.DisplayName, comment.DateCreated.ToString())
                Console.WriteLine(comment.Text)
                If comment.Mentions IsNot Nothing Then
                    For Each mention In comment.Mentions
                        Dim personMentioned = persons(mention.MentionPersonId)
                        Console.WriteLine("{0} was mentioned in a comment", personMentioned.DisplayName)
                        Console.WriteLine("Identity provider: {0}", personMentioned.ProviderId.ToString())
                    Next
                End If
                Console.WriteLine("***************************")
            Next

            ' finally close the thread (can be opened again with the ReopenThread method)
            thread.ResolveThread()
            If thread.IsResolved Then
                Console.WriteLine("The thread is now resolved!")
            End If

            ' for backward compatibility a comment/note is created in a cell containing a threaded comment
            ' if threaded comments is not supported the user will see this comment instead
            Dim legacyComment = sheet.Cells("A1").Comment
            Console.WriteLine("Legacy comment text: {0}", legacyComment.Text)

            ' add a thread in cell B1, add a comment
            Dim thread2 = sheet.ThreadedComments.Add("B1")
            Dim c = thread2.AddComment(user1.Id, "Hello")
            Console.WriteLine("B1 now contains a thread with {0} comment", thread2.Comments.Count)
            ' remove the comment
            thread2.Remove(c)
            If thread2.Comments.Count = 0 Then Console.WriteLine("Thread is now empty")
        End Sub
    End Module
End Namespace
