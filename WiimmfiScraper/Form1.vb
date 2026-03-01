Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports HtmlAgilityPack
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Drawing

Public Class Form1
    Inherits Form

    Private isRunning As Boolean = False
    Private outputFile As String = "room_data.json"

    ' UI Elemente
    Private txtUrl As TextBox
    Private txtFC As TextBox
    Private WithEvents btnStart As Button
    Private lblStatus As Label
    Private WithEvents Timer1 As Timer
    Private WithEvents webView As WebView2

    Public Sub New()
        ' Fenster-Eigenschaften
        Me.Text = "Wiimmfi Scraper (Auto-Generated)"
        Me.Size = New Size(800, 600)

        ' UI Elemente initialisieren
        txtUrl = New TextBox() With { .Location = New Point(10, 10), .Width = 400, .Text = "https://wiimmfi.de/stats/mkw/room/p1" }
        txtFC = New TextBox() With { .Location = New Point(10, 40), .Width = 200, .Text = "1234-5678-9012" }
        btnStart = New Button() With { .Location = New Point(220, 38), .Text = "Starten" }
        lblStatus = New Label() With { .Location = New Point(10, 70), .Width = 400, .Text = "Bereit." }
        
        webView = New WebView2() With { .Location = New Point(10, 100), .Size = New Size(760, 450), .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right }
        Timer1 = New Timer() With { .Interval = 30000 }

        AddHandler Me.Load, AddressOf Form1_Load

        Me.Controls.Add(txtUrl)
        Me.Controls.Add(txtFC)
        Me.Controls.Add(btnStart)
        Me.Controls.Add(lblStatus)
        Me.Controls.Add(webView)
    End Sub

    Private Async Sub Form1_Load(sender As Object, e As EventArgs)
        Await webView.EnsureCoreWebView2Async(Nothing)
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        If Not isRunning Then
            isRunning = True
            btnStart.Text = "Stoppen"
            lblStatus.Text = "Lade Seite..."
            webView.CoreWebView2.Navigate(txtUrl.Text)
            Timer1.Start()
        Else
            isRunning = False
            btnStart.Text = "Starten"
            Timer1.Stop()
            lblStatus.Text = "Gestoppt."
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If isRunning AndAlso webView.CoreWebView2 IsNot Nothing Then
            lblStatus.Text = "Lade Seite neu..."
            webView.CoreWebView2.Reload()
        End If
    End Sub

    Private Async Sub webView_NavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs) Handles webView.NavigationCompleted
        If Not isRunning Then Return

        lblStatus.Text = "Analysiere Daten..."
        Try
            Dim html As String = Await webView.ExecuteScriptAsync("document.documentElement.outerHTML;")
            html = Regex.Unescape(html)
            If html.Length >= 2 Then
                html = html.Substring(1, html.Length - 2)
            End If
            ParseHtmlAndSave(html)
        Catch ex As Exception
            lblStatus.Text = "Fehler: " & ex.Message
        End Try
    End Sub

    Private Sub ParseHtmlAndSave(html As String)
        Dim doc As New HtmlAgilityPack.HtmlDocument()
        doc.LoadHtml(html)

        Dim foundRoom As Boolean = False
        Dim rData As New Dictionary(Of String, Object)
        Dim targetFC As String = txtFC.Text.Trim()

        Dim tables = doc.DocumentNode.SelectNodes("//table")
        If tables IsNot Nothing Then
            For Each table In tables
                Dim tableText As String = table.InnerText
                
                If tableText.Contains(targetFC) Then
                    foundRoom = True
                    rData.Add("fc", targetFC)
                    rData.Add("status", "In Room")

                    Dim tableTextLower = tableText.ToLower()
                    Dim ccFound As String = "Unknown CC"
                    Dim ccs = New String() {"150cc", "100cc", "50cc", "200cc"}
                    For Each cc In ccs
                        If tableTextLower.Contains(cc) Then
                            ccFound = cc.ToUpper()
                            Exit For
                        End If
                    Next
                    If tableTextLower.Contains("mirror") Then ccFound = "Mirror"
                    rData.Add("cc", ccFound)

                    Dim trackFound As String = "Unknown Track"
                    Dim trackMatch = Regex.Match(tableText, "Tracks:\s*([^,
]+)")
                    If trackMatch.Success Then
                        trackFound = trackMatch.Groups(1).Value.Trim()
                    End If
                    rData.Add("track", trackFound)

                    Dim rows = table.SelectNodes(".//tr")
                    Dim playerCount As Integer = 0
                    Dim driver As String = "Unknown"
                    Dim vehicle As String = "Unknown"

                    If rows IsNot Nothing Then
                        For Each row In rows
                            If row.SelectNodes(".//td") IsNot Nothing Then
                                playerCount += 1
                                If row.InnerText.Contains(targetFC) Then
                                    Dim imgs = row.SelectNodes(".//img")
                                    If imgs IsNot Nothing Then
                                        Dim titles = imgs.Select(Function(img) img.GetAttributeValue("title", "")).
                                                          Where(Function(t) Not String.IsNullOrEmpty(t)).ToList()
                                        If titles.Count >= 2 Then
                                            driver = titles(0)
                                            vehicle = titles(1)
                                        ElseIf titles.Count = 1 Then
                                            driver = titles(0)
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If

                    rData.Add("player_count", playerCount)
                    rData.Add("driver", driver)
                    rData.Add("vehicle", vehicle)
                    Exit For
                End If
            Next
        End If

        If foundRoom Then
            Dim json = JsonConvert.SerializeObject(rData, Formatting.Indented)
            File.WriteAllText(outputFile, json)
            lblStatus.Text = $"Gefunden! ({rData("track")} - {rData("cc")})"
        Else
            File.WriteAllText(outputFile, "{}")
            lblStatus.Text = "Spieler nicht im Raum. JSON geleert."
        End If
    End Sub
End Class
