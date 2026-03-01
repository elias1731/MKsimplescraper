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
    Private WithEvents btnStart As Button
    Private lblStatus As Label
    Private WithEvents Timer1 As Timer
    Private WithEvents webView As WebView2

    Public Sub New()
        Me.Text = "Wiimmfi Scraper (Auto-Generated)"
        Me.Size = New Size(800, 600)

        ' UI Elemente initialisieren
        txtUrl = New TextBox() With { .Location = New Point(10, 10), .Width = 600, .Text = "https://wiimmfi.de/stats/mkwx/room/p1" }
        btnStart = New Button() With { .Location = New Point(620, 8), .Text = "Starten" }
        lblStatus = New Label() With { .Location = New Point(10, 40), .Width = 600, .Text = "Bereit." }
        
        webView = New WebView2() With { .Location = New Point(10, 70), .Size = New Size(760, 480), .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right }
        Timer1 = New Timer() With { .Interval = 30000 }

        AddHandler Me.Load, AddressOf Form1_Load

        Me.Controls.Add(txtUrl)
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
        
        ' 1. Friend Code aus der <h1> Überschrift lesen
        Dim targetFC As String = ""
        Dim fcMatch = Regex.Match(html, "Room of profile\s*([\d-]+)", RegexOptions.IgnoreCase)
        If fcMatch.Success Then
            targetFC = fcMatch.Groups(1).Value.Trim()
        End If

        If String.IsNullOrEmpty(targetFC) Then
            File.WriteAllText(outputFile, "{}")
            lblStatus.Text = "Fehler: Kein Friend Code in der Kopfzeile gefunden."
            Return
        End If

        ' 2. Track und CC über Regex auslesen
        Dim ccFound As String = "Unknown CC"
        Dim ccMatch = Regex.Match(html, "(150cc|100cc|50cc|200cc|Mirror)", RegexOptions.IgnoreCase)
        If ccMatch.Success Then
            ccFound = ccMatch.Groups(1).Value
        End If

        Dim trackFound As String = "Unknown Track"
        ' Sucht nach "Last track: <a... > TrackName </a>"
        Dim trackMatch = Regex.Match(html, "Last track:\s*<a[^>]*>(.*?)<\/a>", RegexOptions.IgnoreCase)
        If trackMatch.Success Then
            ' Wir entfernen das " (Nintendo)", falls vorhanden
            trackFound = trackMatch.Groups(1).Value.Replace(" (Nintendo)", "").Trim()
        End If

        ' 3. Spieler in der Tabelle suchen
        Dim tables = doc.DocumentNode.SelectNodes("//table")
        If tables IsNot Nothing Then
            For Each table In tables
                If table.InnerText.Contains(targetFC) Then
                    foundRoom = True
                    rData.Add("fc", targetFC)
                    rData.Add("status", "In Room")
                    rData.Add("cc", ccFound)
                    rData.Add("track", trackFound)

                    Dim rows = table.SelectNodes(".//tr")
                    Dim playerCount As Integer = 0
                    Dim driver As String = "Unknown"
                    Dim vehicle As String = "Unknown"

                    If rows IsNot Nothing Then
                        For Each row In rows
                            ' Zähle Spieler: Überprüfe ob es eine Zelle mit einem Friend Code Link gibt
                            Dim fcNode = row.SelectSingleNode(".//td/a[contains(@href, '/list/p')]")
                            If fcNode IsNot Nothing Then
                                playerCount += 1
                            End If

                            ' Befinden wir uns in der Zeile unseres gesuchten FCs?
                            If row.InnerText.Contains(targetFC) Then
                                ' Suche nach td Attributen mit data-tooltip oder title die ein "@" enthalten
                                Dim dvNode = row.SelectSingleNode(".//td[contains(@data-tooltip, '@') or contains(@title, '@')]")
                                If dvNode IsNot Nothing Then
                                    Dim tooltip As String = dvNode.GetAttributeValue("data-tooltip", "")
                                    If String.IsNullOrEmpty(tooltip) Then
                                        tooltip = dvNode.GetAttributeValue("title", "")
                                    End If
                                    
                                    If tooltip.Contains("@") Then
                                        Dim parts = tooltip.Split("@"c)
                                        If parts.Length >= 2 Then
                                            driver = parts(0).Trim()
                                            vehicle = parts(1).Trim()
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

        ' Daten schreiben
        If foundRoom Then
            Dim json = JsonConvert.SerializeObject(rData, Formatting.Indented)
            File.WriteAllText(outputFile, json)
            lblStatus.Text = $"Gefunden! ({rData("track")} - {rData("cc")} - Spieler: {rData("player_count")})"
        Else
            File.WriteAllText(outputFile, "{}")
            lblStatus.Text = "Spieler nicht (mehr) im Raum. JSON geleert."
        End If
    End Sub
End Class
