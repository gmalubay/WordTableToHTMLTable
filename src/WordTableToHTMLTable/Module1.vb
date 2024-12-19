Imports System.IO
Imports System.Xml.Linq
Imports Word = Microsoft.Office.Interop.Word
Imports System.Text.RegularExpressions

Module Module1

    Sub Main()


        Console.WriteLine()
        Console.WriteLine("*****************************************************************************")
        Console.WriteLine("            Word Table to HTML table Ver. 01.01.02 - Apr.06.2024")
        Console.WriteLine("*****************************************************************************")
        Console.WriteLine()



        Dim inputPath As New DirectoryInfo(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\INPUT")
        Dim outputPath As New DirectoryInfo(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\REPORT")
        Dim Mword As New Word.Application

        If Not Directory.Exists(inputPath.FullName) Then Directory.CreateDirectory(inputPath.FullName)
        If Not Directory.Exists(outputPath.FullName) Then Directory.CreateDirectory(outputPath.FullName)

        Dim iMaxProgress As Integer = inputPath.GetFiles("*.doc*", SearchOption.AllDirectories).Count
        Dim iProgress As Integer = 0

        Dim procs As Process() = Process.GetProcessesByName("WINWORD")
        Dim myHashTable As New Hashtable
        Dim iProcess As Integer = 0

        Console.ForegroundColor = ConsoleColor.Cyan

        For Each docFile As FileInfo In inputPath.GetFiles("*.doc*", SearchOption.AllDirectories)
            Console.WriteLine(Space(3) + docFile.Name)

            Dim htmlDocFname As String = Path.Combine(outputPath.FullName, Path.GetFileNameWithoutExtension(docFile.Name))

            If Regex.IsMatch(docFile.Name, "^~") Then
                'RenderProgressBar(iMaxProgress, iProgress, 1, Console.CursorTop)
                iProgress += 1
                Continue For
            End If


            Try
                Mword.Documents.Open(docFile.FullName, False)

                insertStylingPreTag(Mword)

                Mword.ActiveDocument.SaveAs(htmlDocFname & ".html", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFilteredHTML)

                'wordDoc.SaveAs(htmlDocFname & ".html", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFilteredHTML)

            Catch ex As Exception
                MsgBox("There is an error in converting" & vbNewLine & ex.Message & vbNewLine & vbNewLine & ex.StackTrace)
            Finally
                Mword.ActiveDocument.Close()
            End Try


            Dim sConvertedDoc As String = File.ReadAllText(htmlDocFname & ".html")

            Dim strPage As String = String.Empty
            Dim strTitle As String = String.Empty
            Dim strSubtitle As String = String.Empty
            Dim strTitleAlign As String = String.Empty
            Dim strSubtitleAlign As String = String.Empty


            For Each mPara As Match In Regex.Matches(sConvertedDoc, "(<p[^>]*>)(.+?)(</p>)", RegexOptions.Singleline)
                Dim sPara As String = mPara.Groups(2).Value.Trim
                sPara = sPara.Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ")
                sPara = Regex.Replace(sPara, "\s\s+", " ")

                For Each tdElm As Match In Regex.Matches(sPara, "<[^>]+>")
                    If tdElm.Value.Contains("sup") = False Then
                        sPara = Regex.Replace(sPara, Regex.Escape(tdElm.Value), "")
                    End If
                Next


                If Not String.IsNullOrEmpty(sPara) Then
                    Dim alignMatch As Match = Regex.Match(mPara.Groups(1).Value, "(align=)(right|center)")

                    If Regex.IsMatch(sPara, "\d+\s\STAT\.?\s\d+") Then
                        strPage = sPara
                    ElseIf sPara.Contains("&lt;title&gt;") Then
                        strTitle = sPara.Replace("&lt;title&gt;", "").Replace("&lt;/title&gt;", "")
                        strTitleAlign = alignMatch.Groups(2).Value
                    ElseIf sPara.Contains("&lt;subtitle&gt;") Then
                        strSubtitle = sPara.Replace("&lt;subtitle&gt;", "").Replace("&lt;/subtitle&gt;", "")
                        strSubtitleAlign = alignMatch.Groups(2).Value
                    End If
                End If
            Next



            For Each mTable As Match In Regex.Matches(sConvertedDoc, "<table[^>]+?>.+</table>", RegexOptions.Singleline)
                'Dim eleTable As XElement = XElement.Load(sTable)
                Dim sTable As String = mTable.Value
                Dim colCount As Integer = 0

                sTable = Regex.Replace(sTable, "(?<=<table)[^>]+(?=>)", "")

                sTable = Regex.Replace(sTable, "(?<=<tr)[^>]+(?=>)", "")
                'sTable = Regex.Replace(sTable, "(?<=<td[^s]+)style=[^>]+(?=>)", "")


                '***** <thead> .... </thead>
                For Each trLineMatch As Match In Regex.Matches(sTable, "(<tr[^>]*>)(.+?)(</tr>)", RegexOptions.Singleline)
                    Dim blnThead As Boolean = False
                    If Regex.IsMatch(trLineMatch.Groups(2).Value, "<h[1-7][^>]*>") Then
                        blnThead = True
                    End If

                    If blnThead = True Then
                        Dim theadReplcemnt As String = String.Empty
                        Dim trStyle As String = GetTRTheadStyle(trLineMatch)
                        If String.IsNullOrEmpty(trStyle) Then
                            theadReplcemnt = "<tr class=""header"">"
                        Else
                            theadReplcemnt = String.Format("<tr class=""header"" style=""{0}"">", trStyle)
                        End If

                        For Each tdLineMatch As Match In Regex.Matches(trLineMatch.Groups(2).Value, "(<td[^>]*>)(.+?)(</td>)", RegexOptions.Singleline)
                            Dim tdData As String = tdLineMatch.Groups(2).Value.Replace(vbCrLf, " ")
                            colCount += 1
                            For Each tdElm As Match In Regex.Matches(tdData, "<[^>]+>")
                                If tdElm.Value.Contains("sup") = False Then
                                    tdData = Regex.Replace(tdData, Regex.Escape(tdElm.Value), "")
                                End If
                            Next



                            If Regex.IsMatch(tdData.Trim, "(&lt;/(b|i|u|sc|bi)&gt;)$") And Regex.IsMatch(tdData.Trim, "^(&lt;(b|i|u|sc|bi)&gt;)") Then
                                tdData = Regex.Replace(tdData.Trim, "(&lt;/?(b|i|u|sc|bi)&gt;)", "")
                            End If




                            Dim tdStyle As String = GetTDStyle(tdLineMatch)

                            theadReplcemnt += String.Format("{2}<th {0}>{1}</th>", tdStyle, tdData.Trim, vbCrLf)
                        Next

                        theadReplcemnt = String.Concat("<thead>", vbCrLf, theadReplcemnt, vbCrLf, "</tr>", vbCrLf, "</thead>")

                        sTable = sTable.Replace(trLineMatch.Value, theadReplcemnt)

                        Exit For
                    End If
                Next
                '**************************


                For Each tdMatch As Match In Regex.Matches(sTable, "(<td[^>]*>)(.*?)(</td>)", RegexOptions.Singleline)
                    Dim tdStyle As String = GetTDStyle(tdMatch)
                    Dim tdData As String = tdMatch.Groups(2).Value
                    If Regex.IsMatch(tdData, "\.\.\.+") Then
                        tdData = Regex.Replace(tdData, "\.\.\.+", "")
                    End If





                    Dim sReplacement As String = String.Format("<td {0}>{1}</td>", tdStyle, tdData)

                    Dim colspan As Match = Regex.Match(tdMatch.Value, " colspan=\d+")
                    Dim rowspan As Match = Regex.Match(tdMatch.Value, " rowspan=\d+")

                    If colspan.Success Then
                        sReplacement = sReplacement.Replace("<td ", String.Concat("<td ", colspan.Value.Trim, " "))
                    End If
                    If rowspan.Success Then
                        sReplacement = sReplacement.Replace("<td ", String.Concat("<td ", rowspan.Value.Trim, " "))
                    End If
                    sTable = sTable.Replace(tdMatch.Value, sReplacement)
                Next


                sTable = Regex.Replace(sTable, "<[/]?(font|span|xml|del|ins|p|br|o)[^>]*?>", "")

                ' sTable = Symbol2Unicode(sTable)

                sTable = sTable.Replace("&lt;bi&gt;", "<span class=""EmphasisTypeBoldItalic"">")
                sTable = sTable.Replace("&lt;b&gt;", "<span class=""EmphasisTypeSmallCaps"">")
                sTable = sTable.Replace("&lt;i&gt;", "<span class=""EmphasisTypeBold"">")
                sTable = sTable.Replace("&lt;sc&gt;", "<span class=""EmphasisTypeItalic"">")
                sTable = sTable.Replace("&lt;u&gt;", "<span class=""EmphasisTypeUnderline"">")
                sTable = Regex.Replace(sTable, "&lt;/(b|i|u|sc|bi)&gt;", "</span>")


                For Each tdmatch As Match In Regex.Matches(sTable, "(?Is)(<td [^>]+>)(.*?)(</td>)")
                    Dim tdtxt As String = tdmatch.Groups(2).Value.Replace(vbCrLf, " ").Trim
                    sTable = sTable.Replace(tdmatch.Value, tdmatch.Groups(1).Value & tdtxt & tdmatch.Groups(3).Value)
                Next


                '***** footnote reference ****
                Dim arrsup As New ArrayList
                For Each supmatch As Match In Regex.Matches(sTable, "<sup>([^\<]+)</sup>")
                    Dim suptxt As String = supmatch.Groups(1).Value
                    If arrsup.Contains(supmatch.Value) = False Then
                        sTable = Regex.Replace(sTable, supmatch.Value, "<ref xmlns=""http://schemas.gpo.gov/xml/uslm"" class=""footnoteRef"" idref=""fntable" & suptxt & """>" & supmatch.Value & "</ref>")
                        arrsup.Add(supmatch.Value)
                    End If
                Next
                '***************

                If sTable.Contains("</thead>") Then
                    sTable = sTable.Replace("</thead>", "</thead>" & vbCrLf & "<tbody>")
                Else
                    sTable = sTable.Replace("<table>", "<table>" & vbCrLf & "<tbody>")
                End If

                sTable = sTable.Replace("</table>", "</tbody>" & vbCrLf & "</table>")


                '****** Table Footer ******
                Dim txtFoot As String = sConvertedDoc.Substring(sConvertedDoc.IndexOf("</table>"))
                txtFoot = txtFoot.Replace("</table>", "").Trim

                Dim strTableFoot As String = String.Empty
                If Not String.IsNullOrEmpty(txtFoot) Then
                    strTableFoot = CreateTableFootnote(txtFoot, colCount)
                    If Not String.IsNullOrEmpty(strTableFoot) Then
                        sTable = sTable.Replace("</table>", strTableFoot & vbCrLf & "</table>")
                    End If
                End If
                '******


                If Not String.IsNullOrEmpty(strTitle) Then
                    For Each supmatch As Match In Regex.Matches(strTitle, "<sup>([^\<]+)</sup>")
                        Dim suptxt As String = supmatch.Groups(1).Value
                        strTitle = Regex.Replace(strTitle, supmatch.Value, "<ref xmlns=""http://schemas.gpo.gov/xml/uslm"" class=""footnoteRef"" idref=""fntable" & suptxt & """>" & supmatch.Value & "</ref>")
                    Next

                    If strTitleAlign = "center" Then
                        strTitleAlign = "centered"
                    Else
                        strTitleAlign = "rightAlign"
                    End If
                    strTitle = String.Format("<p xmlns=""http://schemas.gpo.gov/xml/uslm"" role=""title"" class=""{0}"">{1}</p>", strTitleAlign, strTitle)
                End If

                If Not String.IsNullOrEmpty(strSubtitle) Then
                    For Each supmatch As Match In Regex.Matches(strSubtitle, "<sup>([^\<]+)</sup>")
                        Dim suptxt As String = supmatch.Groups(1).Value
                        strSubtitle = Regex.Replace(strSubtitle, supmatch.Value, "<ref xmlns=""http://schemas.gpo.gov/xml/uslm"" class=""footnoteRef"" idref=""fntable" & suptxt & """>" & supmatch.Value & "</ref>")
                    Next

                    If strSubtitleAlign = "center" Then
                        strSubtitleAlign = "centered"
                    Else
                        strSubtitleAlign = "rightAlign"
                    End If
                    strSubtitle = String.Format("{2}<p xmlns=""http://schemas.gpo.gov/xml/uslm"" role=""subtitle"" class=""{0}"">{1}</p>", strSubtitleAlign, strSubtitle, vbCrLf)
                End If


                If Not String.IsNullOrEmpty(strPage) Then
                    strPage = String.Format("<page xmlns=""http://schemas.gpo.gov/xml/uslm"">{0}</page>", strPage)
                End If


                sTable = sTable.Replace("<table>", String.Format("{3}{0}<table>{0}<caption>{0}{1}{2}{0}</caption>", vbCrLf, strTitle, strSubtitle, strPage))


                sTable = sTable.Replace("<table>", "<table xmlns=""http://www.w3.org/1999/xhtml"" width=""100%"" style=""border-collapse:collapse"">")



                sTable = sTable.Replace("&amp;", "&")

                File.WriteAllText(htmlDocFname & ".txt", sTable)
            Next




            'Directory.Delete(Path.GetFullPath(htmlDocFname) & "_files", True)
            If File.Exists(htmlDocFname & ".html") Then File.Delete(htmlDocFname & ".html")

                'RenderProgressBar(iMaxProgress, iProgress, 1, Console.CursorTop)
                iProgress += 1
            Next

            Mword = Nothing

        Console.ForegroundColor = ConsoleColor.White

        For Each proc As Process In procs
            'If Not myHashTable.ContainsKey(proc.Id) Then
            proc.CloseMainWindow()
            If Not proc.HasExited Then
                proc.Kill()
            End If
            proc.Close()
            'End If
        Next proc

        If iMaxProgress = 0 Then
            Console.WriteLine("  There are no available doc(x) For convertion. Try Again Later.")
            Console.WriteLine("***************************************************************")

        Else

            Console.WriteLine()
            Console.WriteLine("Done Converting Doc(x) File(s)")
            Console.WriteLine("Converted Doc(x) File(s) located @: " & vbNewLine & outputPath.FullName)
            Console.WriteLine()
            Console.WriteLine("Show Converted File's containing folder? (Y/N)")

            Dim sResponse As String = Console.ReadLine()

            If sResponse = "Y" Or sResponse = "y" Then
                Process.Start(outputPath.FullName)
            End If
            Console.WriteLine("***************************************************************")
        End If

        Console.ReadLine()

    End Sub

    Function CreateTableFootnote(ByVal TableFootnote As String, ByVal colCount As String) As String
        Dim tablefoot As String = String.Empty

        Dim collOutput As New ArrayList
        Dim collFnote As New ArrayList
        Dim txtFnote As String = TableFootnote
        txtFnote = Regex.Replace(txtFnote, "<span[^>]+>", "").Replace("</span>", "")
        txtFnote = Regex.Replace(txtFnote, "<p[^>]+>&nbsp;</p>", "").Trim
        txtFnote = txtFnote.Replace(ChrW(65533), "")
        Dim fn As String = String.Empty

        For Each mPara As Match In Regex.Matches(txtFnote, "(<p[^>]*>)(.+?)(</p>)", RegexOptions.Singleline)
            If mPara.Value.Contains("<sup>") Then
                If Not String.IsNullOrEmpty(fn) Then
                    collFnote.Add(fn)
                End If

                fn = String.Empty
                fn = mPara.Value
            ElseIf Not String.IsNullOrEmpty(fn) Then
                fn += " " & mPara.Value
            End If
        Next
        If Not String.IsNullOrEmpty(fn) Then
            collFnote.Add(fn)
        End If

        For Each fnote In collFnote
            Dim fnum As String = Regex.Match(fnote, "<sup>(.*?)</sup>").Groups(1).Value.Trim

            Dim txtalign As String = String.Empty
            Dim alignMatch As Match = Regex.Match(fnote, "(align=)(right|center)")
            If alignMatch.Success = False Then
                alignMatch = Regex.Match(fnote, "text\-align\:([^\;]+)\;")
            End If
            If alignMatch.Success Then
                txtalign = alignMatch.Groups(1).Value
            Else
                txtalign = "left"
            End If

            Dim fdata As String = fnote
            For Each Elm As Match In Regex.Matches(fdata, "<[^>]+>")
                If Elm.Value.Contains("sup") = False Then
                    fdata = Regex.Replace(fdata, Regex.Escape(Elm.Value), "")
                End If
            Next
            fdata = fdata.Replace(vbCrLf, " ").Trim
            fdata = fdata.Replace("<sup> ", "<sup>").Replace(" </sup>", "</sup>")
            fdata = Regex.Replace(fdata, "\s\s+", " ")

            collOutput.Add(String.Format("<tr>{4}<td colspan=""{0}"" style=""text-align:{1}; text-indent:1em; font-size:6pt"">" &
                                    "<footnote xmlns=""http://schemas.gpo.gov/xml/uslm"" id=""fntable{2}"">{3}</footnote></td>{4}</tr>",
                                    colCount, txtalign, fnum, fdata, vbCrLf))


        Next


        If collOutput.Count > 0 Then

            tablefoot = String.Format("<tfoot>{1}{0}{1}</tfoot>", String.Join(vbCrLf, collOutput.ToArray), vbCrLf)


        End If



        Return tablefoot


    End Function



    Function GetTRTheadStyle(ByVal trLineMatch As Match) As String
        Dim strStyle As String = String.Empty
        Dim trbordrbottom As Match = Regex.Match(trLineMatch.Groups(1).Value, "border-bottom:([^;]+)(;|')")
        Dim trbordrtop As Match = Regex.Match(trLineMatch.Groups(1).Value, "border-top:([^;]+)(;|')")
        Dim trbordrleft As Match = Regex.Match(trLineMatch.Groups(1).Value, "border-left:([^;]+)(;|')")
        Dim trbordright As Match = Regex.Match(trLineMatch.Groups(1).Value, "border-right:([^;]+)(;|')")

        Dim arrTr As New ArrayList

        If trbordrbottom.Success = True Then
            If trbordrbottom.Value.Contains("none") = False Then
                arrTr.Add("border-bottom:1pc solid black")
            End If
        End If
        If trbordrtop.Success = True Then
            If trbordrtop.Value.Contains("none") = False Then
                arrTr.Add("border-top:1pc solid black")
            End If
        End If

        If arrTr.Count <> 0 Then
            strStyle = String.Join("; ", arrTr.ToArray)
        End If

        Return strStyle
    End Function


    Function GetTDStyle(ByVal tdLineMatch As Match) As String

        Dim tdStyle As String = String.Empty

        Dim valign As Match = Regex.Match(tdLineMatch.Groups(1).Value, "valign=(middle|top|bottom)(;|'|\s)")
        Dim alignMatch As Match = Regex.Match(tdLineMatch.Groups(2).Value, "(align=)(right|center)(;|'|\s)")
        Dim borderbottom As Match = Regex.Match(tdLineMatch.Groups(1).Value, "border-bottom:([^;]+)(;|'|\s)")
        Dim bordertop As Match = Regex.Match(tdLineMatch.Groups(1).Value, "border-top:([^;]+)(;|'|\s)")
        Dim borderleft As Match = Regex.Match(tdLineMatch.Groups(1).Value, "border-left:([^;]+)(;|'|\s)")
        Dim borderrigth As Match = Regex.Match(tdLineMatch.Groups(1).Value, "border-right:([^;]+)(;|'|\s)")

        Dim border As Match = Regex.Match(tdLineMatch.Groups(1).Value, "border:([^;]+)(;|'|\s)")

        Dim arrTd As New ArrayList


        Dim tmpborder As New ArrayList
        Dim arrborder As New ArrayList



        If alignMatch.Success Then
            arrTd.Add(String.Format("text-align:{0}", alignMatch.Groups(2).Value))
        Else
            arrTd.Add("text-align:left")
        End If
        If valign.Success Then
            arrTd.Add(String.Format("vertical-align:{0}", valign.Groups(1).Value))
        End If

        If borderleft.Success Then
            If borderleft.Groups(1).Value.Contains("none") = False Then
                arrborder.Add("border-left:1px solid black")
            End If
        Else
            tmpborder.Add("border-left:1px solid black")
        End If


        If borderrigth.Success Then
            If borderrigth.Groups(1).Value.Contains("none") = False Then
                arrborder.Add("border-right:1px solid black")
            End If
        Else
            tmpborder.Add("border-right:1px solid black")
        End If
        If bordertop.Success Then
            If bordertop.Groups(1).Value.Contains("none") = False Then
                arrborder.Add("border-top:1px solid black")
            End If
        Else
            tmpborder.Add("border-top:1px solid black")
        End If
        If borderbottom.Success Then
            If borderbottom.Groups(1).Value.Contains("none") = False Then
                arrborder.Add("border-bottom:1px solid black")
            End If
        Else
            tmpborder.Add("border-bottom:1px solid black")
        End If



        If tmpborder.Count = 0 Then
        Else
            If border.Success = True Then
                If border.Groups(1).Value.Contains("none") = False Then
                    For Each tb In tmpborder
                        If arrborder.Contains(tb) = False Then
                            arrborder.Add(tb)
                        End If
                    Next

                End If

            End If
        End If

        If arrborder.Count = 0 Then
        Else
            arrTd.AddRange(arrborder)
        End If



        tdStyle = String.Format("style=""{0}""", String.Join("; ", arrTd.ToArray))


        If Regex.IsMatch(tdLineMatch.Groups(2).Value, "\.\.\.+") Then
            tdStyle += " leaders=""yes"""
        End If


        Return tdStyle
    End Function




    Sub insertStylingPreTag(ByVal Mword As Word.Application)

        Mword.Visible = False

        Mword.Selection.HomeKey(Word.WdUnits.wdStory)

        Mword.Selection.Find.ClearFormatting()
        Mword.Selection.Find.Font.Italic = True
        Mword.Selection.Find.Font.Bold = False
        With Mword.Selection.Find
            .Text = ""
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute()
        End With

        Dim sPrevSelection As String = ""

        While Mword.Selection.Find.Found
            'Mword.Selection.Find.Replacement.Text = "<i>" & Mword.Selection.Text & "</i>"

            If Not Mword.Selection.Text.Contains(sPrevSelection) Or sPrevSelection = "" Then
                sPrevSelection = Mword.Selection.Text
                Mword.Selection.InsertBefore("<i>")
                Mword.Selection.InsertAfter("</i>")
                'Mword.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                Mword.Selection.Find.Execute()

            Else
                Exit While
            End If
        End While

        Mword.Selection.HomeKey(Word.WdUnits.wdStory)

        Mword.Selection.Find.ClearFormatting()
        Mword.Selection.Find.Font.Bold = True
        Mword.Selection.Find.Font.Italic = False
        With Mword.Selection.Find
            .Text = ""
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute()
        End With

        sPrevSelection = ""

        While Mword.Selection.Find.Found
            'Mword.Selection.Find.Replacement.Text = "<i>" & Mword.Selection.Text & "</i>"

            If Not Mword.Selection.Text.Contains(sPrevSelection) Or sPrevSelection = "" Then
                sPrevSelection = Mword.Selection.Text
                Mword.Selection.InsertBefore("<b>")
                Mword.Selection.InsertAfter("</b>")
                'Mword.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                Mword.Selection.Find.Execute()

            Else
                Exit While
            End If
        End While

        Mword.Selection.HomeKey(Word.WdUnits.wdStory)

        Mword.Selection.Find.ClearFormatting()
        Mword.Selection.Find.Font.Underline = True
        With Mword.Selection.Find
            .Text = ""
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute()
        End With

        sPrevSelection = ""

        While Mword.Selection.Find.Found
            'Mword.Selection.Find.Replacement.Text = "<i>" & Mword.Selection.Text & "</i>"

            If Not Mword.Selection.Text.Contains(sPrevSelection) Or sPrevSelection = "" Then
                sPrevSelection = Mword.Selection.Text
                Mword.Selection.InsertBefore("<u>")
                Mword.Selection.InsertAfter("</u>")
                'Mword.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                Mword.Selection.Find.Execute()

            Else
                Exit While
            End If
        End While

        Mword.Selection.HomeKey(Word.WdUnits.wdStory)

        Mword.Selection.Find.ClearFormatting()
        Mword.Selection.Find.Font.SmallCaps = True
        With Mword.Selection.Find
            .Text = ""
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute()
        End With

        sPrevSelection = ""

        While Mword.Selection.Find.Found
            'Mword.Selection.Find.Replacement.Text = "<i>" & Mword.Selection.Text & "</i>"

            If Not Mword.Selection.Text.Contains(sPrevSelection) Or sPrevSelection = "" Then
                sPrevSelection = Mword.Selection.Text
                Mword.Selection.InsertBefore("<sc>")
                Mword.Selection.InsertAfter("</sc>")
                'Mword.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                Mword.Selection.Find.Execute()

            Else
                Exit While
            End If
        End While

        Mword.Selection.HomeKey(Word.WdUnits.wdStory)

        Mword.Selection.Find.ClearFormatting()
        Mword.Selection.Find.Font.Bold = True
        Mword.Selection.Find.Font.Italic = True
        With Mword.Selection.Find
            .Text = ""
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute()
        End With

        sPrevSelection = ""

        While Mword.Selection.Find.Found
            'Mword.Selection.Find.Replacement.Text = "<i>" & Mword.Selection.Text & "</i>"

            If Not Mword.Selection.Text.Contains(sPrevSelection) Or sPrevSelection = "" Then
                sPrevSelection = Mword.Selection.Text
                Mword.Selection.InsertBefore("<bi>")
                Mword.Selection.InsertAfter("</bi>")
                'Mword.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                Mword.Selection.Find.Execute()

            Else
                Exit While
            End If
        End While

    End Sub


    Sub RenderProgressBar(ByVal intMaxValue As Integer, ByVal intProgress As Integer, ByVal intLeftPos As Integer, ByVal intTopPos As Integer)

        Dim strResult As String
        Dim intPercent As Integer

        If (intMaxValue > 0) Then
            intPercent = Math.Round((intProgress / intMaxValue) * 100, 0)
            If intPercent >= 1 Then
                strResult = "| " & intPercent & "% |" & StrDup(intPercent \ 2, "#") & StrDup(50 - (intPercent \ 2), " ") & "|"
            Else
                strResult = "| " & intPercent & "% |" & StrDup(50, " ") & "|"
            End If
        Else
            intPercent = intProgress
            strResult = "| " & intPercent & " |"
        End If
        Console.CursorVisible = False
        Console.SetCursorPosition(intLeftPos, intTopPos)
        Console.Write(strResult)
        Console.CursorVisible = True

    End Sub

    Public Function Symbol2Unicode(ByVal ContentText As String) As String
        REM Symbol -> AscW(Symbol) -> convert to Hex -> make Unicode of symbol
        REM -------------------------------------------------------------------

        Dim tmpwholetext As String = ContentText
        Dim Str_symbol As String = " A-Za-z0-9\!~`@#\$%\^\&\*\(\)\[\]\{\}'"";\:/\?\>\<,\.\\|\-=\+_\r\n\s" & vbTab & vbCr & vbCrLf
        Dim SplMatch As Match = Regex.Match(tmpwholetext, "([^" & Str_symbol & "])")
        While SplMatch.Success
            Dim splChr As String = SplMatch.Groups(1).ToString
            If splChr <> "" Then
                Dim charNo As String = AscW(splChr)
                charNo = Hex(charNo)
                While charNo.Length < 4
                    charNo = "0" & charNo
                End While
                Dim Repl_str As String = "&#x" & charNo & ";"
                ContentText = Replace(ContentText, splChr, Repl_str)
            End If
            tmpwholetext = Replace(tmpwholetext, splChr, "Z")
            SplMatch = Regex.Match(tmpwholetext, "([^" & Str_symbol & "])")
        End While
        ContentText = Regex.Replace(ContentText, "[\&]+", "&")
        Return ContentText
    End Function

End Module
