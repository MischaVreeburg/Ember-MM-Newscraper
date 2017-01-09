' ################################################################################
' #                             EMBER MEDIA MANAGER                              #
' ################################################################################
' ################################################################################
' # This file is part of Ember Media Manager.                                    #
' #                                                                              #
' # Ember Media Manager is free software: you can redistribute it and/or modify  #
' # it under the terms of the GNU General Public License as published by         #
' # the Free Software Foundation, either version 3 of the License, or            #
' # (at your option) any later version.                                          #
' #                                                                              #
' # Ember Media Manager is distributed in the hope that it will be useful,       #
' # but WITHOUT ANY WARRANTY; without even the implied warranty of               #
' # MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                #
' # GNU General Public License for more details.                                 #
' #                                                                              #
' # You should have received a copy of the GNU General Public License            #
' # along with Ember Media Manager.  If not, see <http://www.gnu.org/licenses/>. #
' ################################################################################

Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports EmberAPI
Imports NLog
Imports WinForms = System.Windows.Forms
Imports Ember_Media_Manager
Imports Xceed.Wpf.Toolkit
Imports System.Windows
Imports System
Imports System.ComponentModel
Imports System.Windows.Threading

Class frmMainWindow
    Public Property tmrLoad_TVEpisode As New WinForms.Timer With {.Interval = 100}
    Public Property tmrLoad_TVSeason As New WinForms.Timer With {.Interval = 100}
    Public Property tmrWait_TVSeason As New WinForms.Timer With {.Interval = 100}
    Public Property tmrLoad_TVShow As New WinForms.Timer With {.Interval = 100}
    Public Property tmrWait_MovieSet As New WinForms.Timer With {.Interval = 250}
    Public Property tmrLoad_MovieSet As New WinForms.Timer With {.Interval = 100}
    Public Property tmrWait_TVShow As New WinForms.Timer With {.Interval = 250}
    Public Property tmrWait_TVEpisode As New WinForms.Timer With {.Interval = 250}
    Public Property tmrWait_Movie As New WinForms.Timer With {.Interval = 250}
    Public Property tmrLoad_Movie As New WinForms.Timer With {.Interval = 100}
    Public Property tmrSearch_MovieSets As New WinForms.Timer With {.Interval = 250}
    Public Property tmrSearchWait_MovieSets As New WinForms.Timer With {.Interval = 250}
    Public Property tmrSearch_Movies As New WinForms.Timer With {.Interval = 250}
    Public Property tmrSearchWait_Movies As New WinForms.Timer With {.Interval = 250}
    Public Property tmrKeyBuffer As New WinForms.Timer With {.Interval = 2000}
    Public Property tmrAppExit As New WinForms.Timer With {.Interval = 1000}
    Public Property tmrSearchWait_Shows As New WinForms.Timer With {.Interval = 250}
    Public Property tmrSearch_Shows As New WinForms.Timer With {.Interval = 250}
    Public Property tmrRunTasks As New WinForms.Timer With {.Interval = 100}



    Public Property ilColumnIcons As New WinForms.ImageList
    Public Property ilMoviesInSet As New WinForms.ImageList With {.ColorDepth = WinForms.ColorDepth.Depth8Bit, .ImageSize = New System.Drawing.Size(16, 16)}




    Public Shared Sub DoEvents()
        Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, New Action(Sub()

                                                                                        End Sub))
    End Sub

#Region "Fields"

    Shared logger As Logger = LogManager.GetCurrentClassLogger()

    Friend WithEvents bwCheckVersion As New ComponentModel.BackgroundWorker
    Friend WithEvents bwCleanDB As New ComponentModel.BackgroundWorker
    Friend WithEvents bwDownloadPic As New ComponentModel.BackgroundWorker
    Friend WithEvents bwLoadImages_Movie As New ComponentModel.BackgroundWorker
    Friend WithEvents bwLoadImages_MovieSet As New ComponentModel.BackgroundWorker
    Friend WithEvents bwLoadImages_MovieSetMoviePosters As New ComponentModel.BackgroundWorker
    Friend WithEvents bwLoadImages_TVEpisode As New ComponentModel.BackgroundWorker
    Friend WithEvents bwLoadImages_TVSeason As New ComponentModel.BackgroundWorker
    Friend WithEvents bwLoadImages_TVShow As New ComponentModel.BackgroundWorker
    Friend WithEvents bwMovieScraper As New ComponentModel.BackgroundWorker
    Friend WithEvents bwMovieSetScraper As New ComponentModel.BackgroundWorker
    Friend WithEvents bwReload_Movies As New ComponentModel.BackgroundWorker
    Friend WithEvents bwReload_MovieSets As New ComponentModel.BackgroundWorker
    Friend WithEvents bwReload_TVShows As New ComponentModel.BackgroundWorker
    Friend WithEvents bwRewrite_Movies As New ComponentModel.BackgroundWorker
    Friend WithEvents bwTVScraper As New ComponentModel.BackgroundWorker
    Friend WithEvents bwTVEpisodeScraper As New ComponentModel.BackgroundWorker
    Friend WithEvents bwTVSeasonScraper As New ComponentModel.BackgroundWorker

    Public fCommandLine As New CommandLine

    Private TaskList As New List(Of Task)
    Private TasksDone As Boolean = True

    Private alActors As New List(Of String)
    Private FilterPanelIsRaised_Movie As Boolean = False
    Private FilterPanelIsRaised_MovieSet As Boolean = False
    Private FilterPanelIsRaised_TVShow As Boolean = False
    Private InfoPanelState_Movie As Integer = 0 '0 = down, 1 = mid, 2 = up
    Private InfoPanelState_MovieSet As Integer = 0 '0 = down, 1 = mid, 2 = up
    Private InfoPanelState_TVShow As Integer = 0 '0 = down, 1 = mid, 2 = up

    Private bsMovies As New WinForms.BindingSource
    Private bsMovieSets As New WinForms.BindingSource
    Private bsTVEpisodes As New WinForms.BindingSource
    Private bsTVSeasons As New WinForms.BindingSource
    Private bsTVShows As New WinForms.BindingSource

    Private dtMovies As New DataTable
    Private dtMovieSets As New DataTable
    Private dtTVEpisodes As New DataTable
    Private dtTVSeasons As New DataTable
    Private dtTVShows As New DataTable

    Private fScanner As New Scanner
    Private fTaskManager As New TaskManager

    Private GenreImage As System.Drawing.Image
    Private InfoCleared As Boolean = False
    Private LoadingDone As Boolean = False
    Private MainActors As New Images
    Private MainBanner As New Images
    Private MainCharacterArt As New Images
    Private MainClearArt As New Images
    Private MainClearLogo As New Images
    Private MainDiscArt As New Images
    Private MainFanart As New Images
    Private MainFanartSmall As New Images
    Private MainLandscape As New Images
    Private MainPoster As New Images
    Private pbGenre() As WinForms.PictureBox = Nothing
    Private pnlGenre() As WinForms.Panel = Nothing
    Private ReportDownloadPercent As Boolean = False
    Private sHTTP As New HTTP

    'Loading Delays
    Private currRow_Movie As Integer = -1
    Private currRow_MovieSet As Integer = -1
    Private currRow_TVEpisode As Integer = -1
    Private currRow_TVSeason As Integer = -1
    Private currRow_TVShow As Integer = -1
    Private currList As Integer = 0
    Private currThemeType As Theming.ThemeType
    Private prevRow_Movie As Integer = -1
    Private prevRow_MovieSet As Integer = -1
    Private prevRow_TVEpisode As Integer = -1
    Private prevRow_TVSeason As Integer = -1
    Private prevRow_TVShow As Integer = -1

    'list movies
    Private currList_Movies As String = "movielist" 'default movie list SQLite view
    Private listViews_Movies As New Dictionary(Of String, String)

    'list moviesets
    Private currList_MovieSets As String = "setslist" 'default moviesets list SQLite view
    Private listViews_MovieSets As New Dictionary(Of String, String)

    'list shows
    Private currList_TVShows As String = "tvshowlist" 'default tv show list SQLite view
    Private listViews_TVShows As New Dictionary(Of String, String)

    'filter movies
    Private bDoingSearch_Movies As Boolean = False
    Private FilterArray_Movies As New List(Of String)
    Private filDataField_Movies As String = String.Empty
    Private filSearch_Movies As String = String.Empty
    Private filSource_Movies As String = String.Empty
    Private filYear_Movies As String = String.Empty
    Private filGenre_Movies As String = String.Empty
    Private filCountry_Movies As String = String.Empty
    Private filMissing_Movies As String = String.Empty
    Private filTag_Movies As String = String.Empty
    Private currTextSearch_Movies As String = String.Empty
    Private prevTextSearch_Movies As String = String.Empty

    'filter moviesets
    Private bDoingSearch_MovieSets As Boolean = False
    Private FilterArray_MovieSets As New List(Of String)
    Private filSearch_MovieSets As String = String.Empty
    Private filMissing_MovieSets As String = String.Empty
    Private currTextSearch_MovieSets As String = String.Empty
    Private prevTextSearch_MovieSets As String = String.Empty

    'filter shows
    Private bDoingSearch_TVShows As Boolean = False
    Private FilterArray_TVShows As New List(Of String)
    Private filSearch_TVShows As String = String.Empty
    Private filSource_TVShows As String = String.Empty
    Private filGenre_TVShows As String = String.Empty
    Private filMissing_TVShows As String = String.Empty
    Private filTag_TVShows As String = String.Empty
    Private currTextSearch_TVShows As String = String.Empty
    Private prevTextSearch_TVShows As String = String.Empty

    'Theme Information
    Private _bannermaxheight As Integer = 160
    Private _bannermaxwidth As Integer = 285
    Private _characterartmaxheight As Integer = 160
    Private _characterartmaxwidth As Integer = 160
    Private _clearartmaxheight As Integer = 160
    Private _clearartmaxwidth As Integer = 285
    Private _clearlogomaxheight As Integer = 160
    Private _clearlogomaxwidth As Integer = 285
    Private _discartmaxheight As Integer = 160
    Private _discartmaxwidth As Integer = 160
    Private _postermaxheight As Integer = 160
    Private _postermaxwidth As Integer = 160
    Private _fanartsmallmaxheight As Integer = 160
    Private _fanartsmallmaxwidth As Integer = 285
    Private _landscapemaxheight As Integer = 160
    Private _landscapemaxwidth As Integer = 285
    Private tTheme As New Theming
    Private _genrepanelcolor As Drawing.Color = Drawing.Color.Gainsboro
    Private _ipmid As Integer = 280
    Private _ipup As Integer = 500
    Private CloseApp As Boolean = False

    Private _SelectedScrapeType As String = String.Empty
    Private _SelectedScrapeTypeMode As String = String.Empty
    Private _SelectedContentType As String = String.Empty

    Private oldStatus As String = String.Empty

    Private KeyBuffer As String = String.Empty

    Private currMovie As Database.DBElement
    Private currMovieSet As Database.DBElement
    Private currTV As Database.DBElement

#End Region 'Fields

#Region "Delegates"

    Delegate Sub DelegateEvent_Movie(ByVal eType As Enums.ScraperEventType, ByVal Parameter As Object)
    Delegate Sub DelegateEvent_MovieSet(ByVal eType As Enums.ScraperEventType, ByVal Parameter As Object)
    Delegate Sub DelegateEvent_TVShow(ByVal eType As Enums.ScraperEventType, ByVal Parameter As Object)

    Delegate Sub Delegate_dtListAddRow(ByVal dTable As DataTable, ByVal dRow As DataRow)
    Delegate Sub Delegate_dtListRemoveRow(ByVal dTable As DataTable, ByVal dRow As DataRow)
    Delegate Sub Delegate_dtListUpdateRow(ByVal dRow As DataRow, ByVal v As DataRow)

    Delegate Sub Delegate_ChangeToolStripLabel(control As WinForms.ToolStripLabel,
                                               bVisible As Boolean,
                                               strValue As String)
    Delegate Sub Delegate_ChangeToolStripProgressBar(control As WinForms.ToolStripProgressBar,
                                                     bVisible As Boolean,
                                                     iMaximum As Integer,
                                                     iMinimum As Integer,
                                                     iValue As Integer,
                                                     tStyle As WinForms.ProgressBarStyle)

    Delegate Sub MySettingsShow(ByVal dlg As dlgSettings)

#End Region 'Delegates

#Region "Properties"

    Public Property GenrePanelColor() As Drawing.Color
        Get
            Return _genrepanelcolor
        End Get
        Set(ByVal value As Drawing.Color)
            _genrepanelcolor = value
        End Set
    End Property

    Public Property IPMid() As Integer
        Get
            Return _ipmid
        End Get
        Set(ByVal value As Integer)
            _ipmid = value
        End Set
    End Property

    Public Property IPUp() As Integer
        Get
            Return _ipup
        End Get
        Set(ByVal value As Integer)
            _ipup = value
        End Set
    End Property

    Public Property BannerMaxHeight() As Integer
        Get
            Return _bannermaxheight
        End Get
        Set(ByVal value As Integer)
            _bannermaxheight = value
        End Set
    End Property

    Public Property BannerMaxWidth() As Integer
        Get
            Return _bannermaxwidth
        End Get
        Set(ByVal value As Integer)
            _bannermaxwidth = value
        End Set
    End Property

    Public Property CharacterArtMaxHeight() As Integer
        Get
            Return _characterartmaxheight
        End Get
        Set(ByVal value As Integer)
            _characterartmaxheight = value
        End Set
    End Property

    Public Property CharacterArtMaxWidth() As Integer
        Get
            Return _characterartmaxwidth
        End Get
        Set(ByVal value As Integer)
            _characterartmaxwidth = value
        End Set
    End Property

    Public Property ClearArtMaxHeight() As Integer
        Get
            Return _clearartmaxheight
        End Get
        Set(ByVal value As Integer)
            _clearartmaxheight = value
        End Set
    End Property

    Public Property ClearArtMaxWidth() As Integer
        Get
            Return _clearartmaxwidth
        End Get
        Set(ByVal value As Integer)
            _clearartmaxwidth = value
        End Set
    End Property

    Public Property ClearLogoMaxHeight() As Integer
        Get
            Return _clearlogomaxheight
        End Get
        Set(ByVal value As Integer)
            _clearlogomaxheight = value
        End Set
    End Property

    Public Property ClearLogoMaxWidth() As Integer
        Get
            Return _clearlogomaxwidth
        End Get
        Set(ByVal value As Integer)
            _clearlogomaxwidth = value
        End Set
    End Property

    Public Property DiscArtMaxHeight() As Integer
        Get
            Return _discartmaxheight
        End Get
        Set(ByVal value As Integer)
            _discartmaxheight = value
        End Set
    End Property

    Public Property DiscArtMaxWidth() As Integer
        Get
            Return _discartmaxwidth
        End Get
        Set(ByVal value As Integer)
            _discartmaxwidth = value
        End Set
    End Property

    Public Property PosterMaxHeight() As Integer
        Get
            Return _postermaxheight
        End Get
        Set(ByVal value As Integer)
            _postermaxheight = value
        End Set
    End Property

    Public Property PosterMaxWidth() As Integer
        Get
            Return _postermaxwidth
        End Get
        Set(ByVal value As Integer)
            _postermaxwidth = value
        End Set
    End Property

    Public Property FanartSmallMaxHeight() As Integer
        Get
            Return _fanartsmallmaxheight
        End Get
        Set(ByVal value As Integer)
            _fanartsmallmaxheight = value
        End Set
    End Property

    Public Property FanartSmallMaxWidth() As Integer
        Get
            Return _fanartsmallmaxwidth
        End Get
        Set(ByVal value As Integer)
            _fanartsmallmaxwidth = value
        End Set
    End Property

    Public Property LandscapeMaxHeight() As Integer
        Get
            Return _landscapemaxheight
        End Get
        Set(ByVal value As Integer)
            _landscapemaxheight = value
        End Set
    End Property

    Public Property LandscapeMaxWidth() As Integer
        Get
            Return _landscapemaxwidth
        End Get
        Set(ByVal value As Integer)
            _landscapemaxwidth = value
        End Set
    End Property

#End Region 'Properties

#Region "Methods"

    Public Sub ClearInfo()
        If bwDownloadPic.IsBusy Then bwDownloadPic.CancelAsync()
        If bwLoadImages_Movie.IsBusy Then bwLoadImages_Movie.CancelAsync()
        If bwLoadImages_MovieSet.IsBusy Then bwLoadImages_MovieSet.CancelAsync()
        If bwLoadImages_MovieSetMoviePosters.IsBusy Then bwLoadImages_MovieSetMoviePosters.CancelAsync()
        If bwLoadImages_TVShow.IsBusy Then bwLoadImages_TVShow.CancelAsync()
        If bwLoadImages_TVSeason.IsBusy Then bwLoadImages_TVSeason.CancelAsync()
        If bwLoadImages_TVEpisode.IsBusy Then bwLoadImages_TVEpisode.CancelAsync()

        While bwDownloadPic.IsBusy OrElse bwLoadImages_Movie.IsBusy OrElse bwLoadImages_MovieSet.IsBusy OrElse
                    bwLoadImages_TVShow.IsBusy OrElse bwLoadImages_TVSeason.IsBusy OrElse bwLoadImages_TVEpisode.IsBusy OrElse
                    bwLoadImages_MovieSetMoviePosters.IsBusy
            DoEvents()
            Threading.Thread.Sleep(50)
        End While

        If pbFanArt.Image IsNot Nothing Then
            pbFanArt.Image.Dispose()
            pbFanArt.Image = Nothing
        End If
        MainFanart.Clear()

        If pbBanner.Image IsNot Nothing Then
            pbBanner.Image.Dispose()
            pbBanner.Image = Nothing
        End If
        pnlBanner.Visible = False
        MainBanner.Clear()

        If pbCharacterArt.Image IsNot Nothing Then
            pbCharacterArt.Image.Dispose()
            pbCharacterArt.Image = Nothing
        End If
        pnlCharacterArt.Visible = False
        MainCharacterArt.Clear()

        If pbClearArt.Image IsNot Nothing Then
            pbClearArt.Image.Dispose()
            pbClearArt.Image = Nothing
        End If
        pnlClearArt.Visible = False
        MainClearArt.Clear()

        If pbClearLogo.Image IsNot Nothing Then
            pbClearLogo.Image.Dispose()
            pbClearLogo.Image = Nothing
        End If
        pnlClearLogo.Visible = False
        MainClearLogo.Clear()

        If pbPoster.Image IsNot Nothing Then
            pbPoster.Image.Dispose()
            pbPoster.Image = Nothing
        End If
        pnlPoster.Visible = False
        MainPoster.Clear()

        If pbFanartSmall.Image IsNot Nothing Then
            pbFanartSmall.Image.Dispose()
            pbFanartSmall.Image = Nothing
        End If
        pnlFanartSmall.Visible = False
        MainFanartSmall.Clear()

        If pbLandscape.Image IsNot Nothing Then
            pbLandscape.Image.Dispose()
            pbLandscape.Image = Nothing
        End If
        pnlLandscape.Visible = False
        MainLandscape.Clear()

        If pbDiscArt.Image IsNot Nothing Then
            pbDiscArt.Image.Dispose()
            pbDiscArt.Image = Nothing
        End If
        pnlDiscArt.Visible = False
        MainDiscArt.Clear()

        'remove all current genres
        Try
            For iDel As Integer = 0 To pnlGenre.Count - 1
                scMain.Panel2.Controls.Remove(pbGenre(iDel))
                scMain.Panel2.Controls.Remove(pnlGenre(iDel))
            Next
        Catch
        End Try

        If pbMPAA.Image IsNot Nothing Then
            pbMPAA.Image = Nothing
        End If
        pnlMPAA.Visible = False

        lblFanartSmallSize.Text = String.Empty
        lblTitle.Text = String.Empty
        lblOriginalTitle.Text = String.Empty
        lblPosterSize.Text = String.Empty
        lblRating.Text = String.Empty
        lblRuntime.Text = String.Empty
        lblStudio.Text = String.Empty
        pnlTop250.Visible = False
        lblTop250.Text = String.Empty
        pbStar1.Image = Nothing
        pbStar2.Image = Nothing
        pbStar3.Image = Nothing
        pbStar4.Image = Nothing
        pbStar5.Image = Nothing
        pbStar6.Image = Nothing
        pbStar7.Image = Nothing
        pbStar8.Image = Nothing
        pbStar9.Image = Nothing
        pbStar10.Image = Nothing
        ToolTips.SetToolTip(pbStar1, "")
        ToolTips.SetToolTip(pbStar2, "")
        ToolTips.SetToolTip(pbStar3, "")
        ToolTips.SetToolTip(pbStar4, "")
        ToolTips.SetToolTip(pbStar5, "")
        ToolTips.SetToolTip(pbStar6, "")
        ToolTips.SetToolTip(pbStar7, "")
        ToolTips.SetToolTip(pbStar8, "")
        ToolTips.SetToolTip(pbStar9, "")
        ToolTips.SetToolTip(pbStar10, "")

        lstActors.Items.Clear()
        If alActors IsNot Nothing Then
            alActors.Clear()
            alActors = Nothing
        End If
        If pbActors.Image IsNot Nothing Then
            pbActors.Image.Dispose()
            pbActors.Image = Nothing
        End If
        MainActors.Clear()
        lblDirectors.Text = String.Empty
        lblReleaseDate.Text = String.Empty
        txtCertifications.Text = String.Empty
        txtIMDBID.Text = String.Empty
        txtFilePath.Text = String.Empty
        txtOutline.Text = String.Empty
        txtPlot.Text = String.Empty
        txtTMDBID.Text = String.Empty
        txtTrailerPath.Text = String.Empty
        lblTagline.Text = String.Empty
        If pbMPAA.Image IsNot Nothing Then
            pbMPAA.Image.Dispose()
            pbMPAA.Image = Nothing
        End If
        pbStudio.Image = Nothing
        pbVideo.Image = Nothing
        pbVType.Image = Nothing
        pbAudio.Image = Nothing
        pbResolution.Image = Nothing
        pbChannels.Image = Nothing
        pbAudioLang0.Image = Nothing
        pbAudioLang1.Image = Nothing
        pbAudioLang2.Image = Nothing
        pbAudioLang3.Image = Nothing
        pbAudioLang4.Image = Nothing
        pbAudioLang5.Image = Nothing
        pbAudioLang6.Image = Nothing
        ToolTips.SetToolTip(pbAudioLang0, "")
        ToolTips.SetToolTip(pbAudioLang1, "")
        ToolTips.SetToolTip(pbAudioLang2, "")
        ToolTips.SetToolTip(pbAudioLang3, "")
        ToolTips.SetToolTip(pbAudioLang4, "")
        ToolTips.SetToolTip(pbAudioLang5, "")
        ToolTips.SetToolTip(pbAudioLang6, "")
        pbSubtitleLang0.Image = Nothing
        pbSubtitleLang1.Image = Nothing
        pbSubtitleLang2.Image = Nothing
        pbSubtitleLang3.Image = Nothing
        pbSubtitleLang4.Image = Nothing
        pbSubtitleLang5.Image = Nothing
        pbSubtitleLang6.Image = Nothing
        ToolTips.SetToolTip(pbSubtitleLang0, "")
        ToolTips.SetToolTip(pbSubtitleLang1, "")
        ToolTips.SetToolTip(pbSubtitleLang2, "")
        ToolTips.SetToolTip(pbSubtitleLang3, "")
        ToolTips.SetToolTip(pbSubtitleLang4, "")
        ToolTips.SetToolTip(pbSubtitleLang5, "")
        ToolTips.SetToolTip(pbSubtitleLang6, "")

        txtMetaData.Text = String.Empty

        lvMoviesInSet.Items.Clear()
        ilMoviesInSet.Images.Clear()

        InfoCleared = True

        DoEvents()
    End Sub

    Private Function CheckColumnHide_Movies(ByVal ColumnName As String) As Boolean
        Dim lsColumn As Settings.ListSorting = Master.eSettings.MovieGeneralMediaListSorting.FirstOrDefault(Function(l) l.Column = ColumnName)
        Return If(lsColumn Is Nothing, True, lsColumn.Hide)
    End Function

    Private Function CheckColumnHide_MovieSets(ByVal ColumnName As String) As Boolean
        Dim lsColumn As Settings.ListSorting = Master.eSettings.MovieSetGeneralMediaListSorting.FirstOrDefault(Function(l) l.Column = ColumnName)
        Return If(lsColumn Is Nothing, True, lsColumn.Hide)
    End Function

    Private Function CheckColumnHide_TVEpisodes(ByVal ColumnName As String) As Boolean
        Dim lsColumn As Settings.ListSorting = Master.eSettings.TVGeneralEpisodeListSorting.FirstOrDefault(Function(l) l.Column = ColumnName)
        Return If(lsColumn Is Nothing, True, lsColumn.Hide)
    End Function

    Private Function CheckColumnHide_TVSeasons(ByVal ColumnName As String) As Boolean
        Dim lsColumn As Settings.ListSorting = Master.eSettings.TVGeneralSeasonListSorting.FirstOrDefault(Function(l) l.Column = ColumnName)
        Return If(lsColumn Is Nothing, True, lsColumn.Hide)
    End Function

    Private Function CheckColumnHide_TVShows(ByVal ColumnName As String) As Boolean
        Dim lsColumn As Settings.ListSorting = Master.eSettings.TVGeneralShowListSorting.FirstOrDefault(Function(l) l.Column = ColumnName)
        Return If(lsColumn Is Nothing, True, lsColumn.Hide)
    End Function

    Public Sub LoadMedia(ByVal Scan As Structures.ScanOrClean, Optional ByVal SourceID As Long = -1, Optional ByVal Folder As String = "")
        Try
            SetStatus(Master.eLang.GetString(116, "Performing Preliminary Tasks (Gathering Data)..."))
            tspbLoading.ProgressBar.Style = ProgressBarStyle.Marquee
            tspbLoading.Visible = True

            DoEvents()

            SetControlsEnabled(False)

            fScanner.CancelAndWait()

            If Scan.MovieSets Then
                prevRow_MovieSet = -1
                dgvMovieSets.DataSource = Nothing
            End If

            fScanner.Start(Scan, SourceID, Folder)

        Catch ex As Exception
            LoadingDone = True
            EnableFilters_Movies(True)
            EnableFilters_MovieSets(True)
            EnableFilters_Shows(True)
            SetControlsEnabled(True)
            logger.Error(ex, New StackFrame().GetMethod().Name)
        End Try
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMainHelpAbout.Click
        Using dAbout As New dlgAbout
            dAbout.ShowDialog()
        End Using
    End Sub

    Private Sub mnuMainToolsExportMovies_Click(sender As Object, e As EventArgs) Handles mnuMainToolsExportMovies.Click
        Try
            Dim table As New DataTable
            Dim ds As New DataSet
            Using SQLcommand As SQLite.SQLiteCommand = Master.DB.MyVideosDBConn.CreateCommand()
                SQLcommand.CommandText = "SELECT * FROM movie INNER JOIN MoviesVStreams ON (MoviesVStreams.MovieID = movie.idMovie) INNER JOIN MoviesAStreams ON (MoviesAStreams.MovieID = movie.idMovie);"
                Using SQLreader As SQLite.SQLiteDataReader = SQLcommand.ExecuteReader()
                    ds.Tables.Add(table)
                    ds.EnforceConstraints = False
                    table.Load(SQLreader)
                End Using
            End Using

            Dim saveFileDialog1 As New WinForms.SaveFileDialog()
            saveFileDialog1.FileName = "export_movies" + ".xml"
            saveFileDialog1.Filter = "xml files (*.xml)|*.xml"
            saveFileDialog1.FilterIndex = 2
            saveFileDialog1.RestoreDirectory = True

            If saveFileDialog1.ShowDialog() = WinForms.DialogResult.OK Then
                table.WriteXml(saveFileDialog1.FileName)
            End If
        Catch ex As Exception
            logger.Error(ex, New StackFrame().GetMethod().Name)
        End Try
    End Sub

    Private Sub mnuMainToolsExportTvShows_Click(sender As Object, e As EventArgs) Handles mnuMainToolsExportTvShows.Click
        Try
            Dim table As New DataTable
            Using SQLcommand As SQLite.SQLiteCommand = Master.DB.MyVideosDBConn.CreateCommand()
                SQLcommand.CommandText = "Select * from tvshow;"
                Using SQLreader As SQLite.SQLiteDataReader = SQLcommand.ExecuteReader()
                    'Load the SqlDataReader object to the DataTable object as follows. 
                    table.Load(SQLreader)
                End Using
            End Using

            Dim saveFileDialog1 As New WinForms.SaveFileDialog()
            saveFileDialog1.FileName = "export_tvshows" + ".xml"
            saveFileDialog1.Filter = "xml files (*.xml)|*.xml"
            saveFileDialog1.FilterIndex = 2
            saveFileDialog1.RestoreDirectory = True

            If saveFileDialog1.ShowDialog() = WinForms.DialogResult.OK Then
                table.WriteXml(saveFileDialog1.FileName)
            End If
        Catch ex As Exception
            logger.Error(ex, New StackFrame().GetMethod().Name)
        End Try
    End Sub

    Private Sub ApplyTheme(ByVal tType As Theming.ThemeType)
        pnlInfoPanel.SuspendLayout()

        currThemeType = tType

        tTheme.ApplyTheme(tType)

        Dim currMainTabTag As Structures.MainTabType = DirectCast(tcMain.SelectedTab.Tag, Structures.MainTabType)

        Select Case If(currMainTabTag.ContentType = Enums.ContentType.Movie, InfoPanelState_Movie, If(currMainTabTag.ContentType = Enums.ContentType.MovieSet, InfoPanelState_MovieSet, InfoPanelState_TVShow))
            Case 1
                If btnMid.Visible Then
                    pnlInfoPanel.Height = _ipmid
                    btnUp.Enabled = True
                    btnMid.Enabled = False
                    btnDown.Enabled = True
                ElseIf btnUp.Visible Then
                    pnlInfoPanel.Height = _ipup
                    If currMainTabTag.ContentType = Enums.ContentType.Movie Then
                        InfoPanelState_Movie = 2
                    ElseIf currMainTabTag.ContentType = Enums.ContentType.MovieSet Then
                        InfoPanelState_MovieSet = 2
                    ElseIf currMainTabTag.ContentType = Enums.ContentType.TV Then
                        InfoPanelState_TVShow = 2
                    End If
                    btnUp.Enabled = False
                    btnMid.Enabled = True
                    btnDown.Enabled = True
                Else
                    pnlInfoPanel.Height = 25
                    If currMainTabTag.ContentType = Enums.ContentType.Movie Then
                        InfoPanelState_Movie = 0
                    ElseIf currMainTabTag.ContentType = Enums.ContentType.MovieSet Then
                        InfoPanelState_MovieSet = 0
                    ElseIf currMainTabTag.ContentType = Enums.ContentType.TV Then
                        InfoPanelState_TVShow = 0
                    End If
                    btnUp.Enabled = True
                    btnMid.Enabled = True
                    btnDown.Enabled = False
                End If
            Case 2
                If btnUp.Visible Then
                    pnlInfoPanel.Height = _ipup
                    btnUp.Enabled = False
                    btnMid.Enabled = True
                    btnDown.Enabled = True
                ElseIf btnMid.Visible Then
                    pnlInfoPanel.Height = _ipmid

                    If currMainTabTag.ContentType = Enums.ContentType.Movie Then
                        InfoPanelState_Movie = 1
                    ElseIf currMainTabTag.ContentType = Enums.ContentType.MovieSet Then
                        InfoPanelState_MovieSet = 1
                    ElseIf currMainTabTag.ContentType = Enums.ContentType.TV Then
                        InfoPanelState_TVShow = 1
                    End If

                    btnUp.Enabled = True
                    btnMid.Enabled = False
                    btnDown.Enabled = True
                Else
                    pnlInfoPanel.Height = 25
                    If currMainTabTag.ContentType = Enums.ContentType.Movie Then
                        InfoPanelState_Movie = 0
                    ElseIf currMainTabTag.ContentType = Enums.ContentType.MovieSet Then
                        InfoPanelState_MovieSet = 0
                    ElseIf currMainTabTag.ContentType = Enums.ContentType.TV Then
                        InfoPanelState_TVShow = 0
                    End If
                    btnUp.Enabled = True
                    btnMid.Enabled = True
                    btnDown.Enabled = False
                End If
            Case Else
                pnlInfoPanel.Height = 25
                If currMainTabTag.ContentType = Enums.ContentType.Movie Then
                    InfoPanelState_Movie = 0
                ElseIf currMainTabTag.ContentType = Enums.ContentType.MovieSet Then
                    InfoPanelState_MovieSet = 0
                ElseIf currMainTabTag.ContentType = Enums.ContentType.TV Then
                    InfoPanelState_TVShow = 0
                End If

                btnUp.Enabled = True
                btnMid.Enabled = True
                btnDown.Enabled = False
        End Select

        pbActLoad.Visible = False
        pbActors.Image = My.Resources.actor_silhouette
        pbMILoading.Visible = False

        pnlInfoPanel.ResumeLayout()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        btnCancel.Visible = False
        lblCanceling.Visible = True
        prbCanceling.Visible = True

        If bwMovieScraper.IsBusy Then bwMovieScraper.CancelAsync()
        If bwMovieSetScraper.IsBusy Then bwMovieSetScraper.CancelAsync()
        If bwReload_Movies.IsBusy Then bwReload_Movies.CancelAsync()
        If bwReload_MovieSets.IsBusy Then bwReload_MovieSets.CancelAsync()
        If bwReload_TVShows.IsBusy Then bwReload_TVShows.CancelAsync()
        If bwRewrite_Movies.IsBusy Then bwRewrite_Movies.CancelAsync()
        If bwTVEpisodeScraper.IsBusy Then bwTVEpisodeScraper.CancelAsync()
        If bwTVScraper.IsBusy Then bwTVScraper.CancelAsync()
        If bwTVSeasonScraper.IsBusy Then bwTVSeasonScraper.CancelAsync()
        While bwMovieScraper.IsBusy OrElse bwReload_Movies.IsBusy OrElse bwMovieSetScraper.IsBusy OrElse bwReload_MovieSets.IsBusy OrElse
            bwReload_TVShows.IsBusy OrElse bwRewrite_Movies.IsBusy OrElse bwTVEpisodeScraper.IsBusy OrElse bwTVScraper.IsBusy OrElse
            bwTVSeasonScraper.IsBusy
            DoEvents()
            Threading.Thread.Sleep(50)
        End While
    End Sub

    Private Sub btnClearFilters_Movies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearFilters_Movies.Click
        ClearFilters_Movies(True)
    End Sub

    Private Sub btnClearFilters_MovieSets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearFilters_MovieSets.Click
        ClearFilters_MovieSets(True)
    End Sub

    Private Sub btnClearFilters_Shows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearFilters_Shows.Click
        ClearFilters_Shows(True)
    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Dim currMainTabTag As Structures.MainTabType = DirectCast(tcMain.SelectedTab.Tag, Structures.MainTabType)
        tcMain.Focus()
        If currMainTabTag.ContentType = Enums.ContentType.Movie Then
            InfoPanelState_Movie = 0
        ElseIf currMainTabTag.ContentType = Enums.ContentType.MovieSet Then
            InfoPanelState_MovieSet = 0
        ElseIf currMainTabTag.ContentType = Enums.ContentType.TV Then
            InfoPanelState_TVShow = 0
        End If
        MoveInfoPanel()
    End Sub

    Private Sub btnFilterDown_Movies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterDown_Movies.Click
        FilterPanelIsRaised_Movie = False
        FilterMovement_Movies()
    End Sub

    Private Sub btnFilterDown_MovieSets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterDown_MovieSets.Click
        FilterPanelIsRaised_MovieSet = False
        FilterMovement_MovieSets()
    End Sub

    Private Sub btnFilterDown_Shows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterDown_Shows.Click
        FilterPanelIsRaised_TVShow = False
        FilterMovement_Shows()
    End Sub

    Private Sub btnFilterUp_Movies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterUp_Movies.Click
        FilterPanelIsRaised_Movie = True
        FilterMovement_Movies()
    End Sub

    Private Sub btnFilterUp_MovieSets_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterUp_MovieSets.Click

        FilterPanelIsRaised_MovieSet = True
        FilterMovement_MovieSets()
    End Sub

    Private Sub btnFilterUp_Shows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterUp_Shows.Click
        FilterPanelIsRaised_TVShow = True
        FilterMovement_Shows()
    End Sub

    Private Sub btnMarkAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMarkAll.Click
        Dim currMainTabTag = ModulesManager.Instance.RuntimeObjects.MediaTabSelected
        CreateTask(currMainTabTag.ContentType, Enums.SelectionType.All, Enums.TaskManagerType.SetMarkedState, True, String.Empty)
    End Sub

    Private Sub btnUnmarkAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUnmarkAll.Click
        Dim currMainTabTag = ModulesManager.Instance.RuntimeObjects.MediaTabSelected
        CreateTask(currMainTabTag.ContentType, Enums.SelectionType.All, Enums.TaskManagerType.SetMarkedState, False, String.Empty)
    End Sub

    Private Sub btnMid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMid.Click
        Dim currMainTabTag As Structures.MainTabType = DirectCast(tcMain.SelectedTab.Tag, Structures.MainTabType)
        tcMain.Focus()
        If currMainTabTag.ContentType = Enums.ContentType.Movie Then
            InfoPanelState_Movie = 1
        ElseIf currMainTabTag.ContentType = Enums.ContentType.MovieSet Then
            InfoPanelState_MovieSet = 1
        ElseIf currMainTabTag.ContentType = Enums.ContentType.TV Then
            InfoPanelState_TVShow = 1
        End If
        MoveInfoPanel()
    End Sub

    Private Sub btnMIRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMetaDataRefresh.Click
        Dim currMainTabTag As Structures.MainTabType = DirectCast(tcMain.SelectedTab.Tag, Structures.MainTabType)

        If currMainTabTag.ContentType = Enums.ContentType.Movie Then
            If dgvMovies.SelectedRows.Count = 1 Then
                Dim ScrapeModifiers As New Structures.ScrapeModifiers
                Functions.SetScrapeModifiers(ScrapeModifiers, Enums.ModifierType.MainMeta, True)
                CreateScrapeList_Movie(Enums.ScrapeType.SelectedAuto, Master.DefaultOptions_Movie, ScrapeModifiers)
            End If
        ElseIf currMainTabTag.ContentType = Enums.ContentType.TV Then
            If dgvMovies.SelectedRows.Count = 1 AndAlso Not String.IsNullOrEmpty(currTV.Filename) Then
                Dim ScrapeModifiers As New Structures.ScrapeModifiers
                Functions.SetScrapeModifiers(ScrapeModifiers, Enums.ModifierType.EpisodeMeta, True)
                CreateScrapeList_TVEpisode(Enums.ScrapeType.SelectedAuto, Master.DefaultOptions_TV, ScrapeModifiers)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Launch video using system default player
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilePlay.Click
        Functions.Launch(txtFilePath.Text, True)
        'Try
        '    If Not String.IsNullOrEmpty(Me.txtFilePath.Text) Then
        '        If File.Exists(Me.txtFilePath.Text) Then
        '            If Master.isWindows Then
        '                Process.Start(String.Concat("""", Me.txtFilePath.Text, """"))
        '            Else
        '                Using Explorer As New Process
        '                    Explorer.StartInfo.FileName = "xdg-open"
        '                    Explorer.StartInfo.Arguments = String.Format("""{0}""", Me.txtFilePath.Text)
        '                    Explorer.Start()
        '                End Using
        '            End If

        '        End If
        '    End If
        'Catch ex As Exception
        '    logger.Error(ex, New StackFrame().GetMethod().Name)
        'End Try
    End Sub
    ''' <summary>
    ''' Launch trailer using system default player
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnTrailerPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTrailerPlay.Click
        If txtTrailerPath.Text.StartsWith("plugin://plugin.video.youtube") Then
            Functions.Launch(StringUtils.ConvertFromKodiTrailerFormatToYouTubeURL(txtTrailerPath.Text), True)
        Else
            Functions.Launch(txtTrailerPath.Text, True)
        End If
    End Sub
    ''' <summary>
    ''' sorts the movielist by adding date
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>this filter is inverted (DESC first) to get the newest title on the top of the list</remarks>
    Private Sub btnFilterSortDateAdded_Movies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterSortDateAdded_Movies.Click
        If dgvMovies.RowCount > 0 Then
            btnFilterSortRating_Movies.Tag = String.Empty
            btnFilterSortRating_Movies.Image = Nothing
            btnFilterSortDateModified_Movies.Tag = String.Empty
            btnFilterSortDateModified_Movies.Image = Nothing
            btnFilterSortReleaseDate_Movies.Tag = String.Empty
            btnFilterSortReleaseDate_Movies.Image = Nothing
            btnFilterSortTitle_Movies.Tag = String.Empty
            btnFilterSortTitle_Movies.Image = Nothing
            btnFilterSortYear_Movies.Tag = String.Empty
            btnFilterSortYear_Movies.Image = Nothing
            If btnFilterSortDateAdded_Movies.Tag.ToString = "DESC" Then
                btnFilterSortDateAdded_Movies.Tag = "ASC"
                btnFilterSortDateAdded_Movies.Image = My.Resources.asc
                dgvMovies.Sort(dgvMovies.Columns("DateAdded"), System.ComponentModel.ListSortDirection.Ascending)
            Else
                btnFilterSortDateAdded_Movies.Tag = "DESC"
                btnFilterSortDateAdded_Movies.Image = My.Resources.desc
                dgvMovies.Sort(dgvMovies.Columns("DateAdded"), System.ComponentModel.ListSortDirection.Descending)
            End If

            SaveSorting_Movies()
        End If
    End Sub
    ''' <summary>
    ''' sorts the movielist by last modification date
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>this filter is inverted (DESC first) to get the latest modified title on the top of the list</remarks>
    Private Sub btnFilterSortDateModified_Movies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterSortDateModified_Movies.Click
        If dgvMovies.RowCount > 0 Then
            btnFilterSortDateAdded_Movies.Tag = String.Empty
            btnFilterSortDateAdded_Movies.Image = Nothing
            btnFilterSortRating_Movies.Tag = String.Empty
            btnFilterSortRating_Movies.Image = Nothing
            btnFilterSortReleaseDate_Movies.Tag = String.Empty
            btnFilterSortReleaseDate_Movies.Image = Nothing
            btnFilterSortTitle_Movies.Tag = String.Empty
            btnFilterSortTitle_Movies.Image = Nothing
            btnFilterSortYear_Movies.Tag = String.Empty
            btnFilterSortYear_Movies.Image = Nothing
            If btnFilterSortDateModified_Movies.Tag.ToString = "DESC" Then
                btnFilterSortDateModified_Movies.Tag = "ASC"
                btnFilterSortDateModified_Movies.Image = My.Resources.asc
                dgvMovies.Sort(dgvMovies.Columns("DateModified"), System.ComponentModel.ListSortDirection.Ascending)
            Else
                btnFilterSortDateModified_Movies.Tag = "DESC"
                btnFilterSortDateModified_Movies.Image = My.Resources.desc
                dgvMovies.Sort(dgvMovies.Columns("DateModified"), System.ComponentModel.ListSortDirection.Descending)
            End If

            SaveSorting_Movies()
        End If
    End Sub
    ''' <summary>
    ''' sorts the movielist by sort title
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnFilterSortTitle_Movies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterSortTitle_Movies.Click
        If dgvMovies.RowCount > 0 Then
            btnFilterSortDateAdded_Movies.Tag = String.Empty
            btnFilterSortDateAdded_Movies.Image = Nothing
            btnFilterSortDateModified_Movies.Tag = String.Empty
            btnFilterSortDateModified_Movies.Image = Nothing
            btnFilterSortRating_Movies.Tag = String.Empty
            btnFilterSortRating_Movies.Image = Nothing
            btnFilterSortReleaseDate_Movies.Tag = String.Empty
            btnFilterSortReleaseDate_Movies.Image = Nothing
            btnFilterSortYear_Movies.Tag = String.Empty
            btnFilterSortYear_Movies.Image = Nothing
            If btnFilterSortTitle_Movies.Tag.ToString = "ASC" Then
                btnFilterSortTitle_Movies.Tag = "DSC"
                btnFilterSortTitle_Movies.Image = My.Resources.desc
                dgvMovies.Sort(dgvMovies.Columns("SortedTitle"), System.ComponentModel.ListSortDirection.Descending)
            Else
                btnFilterSortTitle_Movies.Tag = "ASC"
                btnFilterSortTitle_Movies.Image = My.Resources.asc
                dgvMovies.Sort(dgvMovies.Columns("SortedTitle"), System.ComponentModel.ListSortDirection.Ascending)
            End If

            SaveSorting_Movies()
        End If
    End Sub
    ''' <summary>
    ''' sorts the tvshowlist by sort title
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnFilterSortTitle_Shows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterSortTitle_Shows.Click
        If dgvTVShows.RowCount > 0 Then
            'Me.btnFilterSortDateAdded_Shows.Tag = String.Empty
            'Me.btnFilterSortDateAdded_Shows.Image = Nothing
            'Me.btnFilterSortDateModified_Shows.Tag = String.Empty
            'Me.btnFilterSortDateModified_Shows.Image = Nothing
            'Me.btnFilterSortRating_Shows.Tag = String.Empty
            'Me.btnFilterSortRating_Shows.Image = Nothing
            'Me.btnFilterSortYear_Shows.Tag = String.Empty
            'Me.btnFilterSortYear_Shows.Image = Nothing
            If btnFilterSortTitle_Shows.Tag.ToString = "ASC" Then
                btnFilterSortTitle_Shows.Tag = "DSC"
                btnFilterSortTitle_Shows.Image = My.Resources.desc
                dgvTVShows.Sort(dgvTVShows.Columns("SortedTitle"), System.ComponentModel.ListSortDirection.Descending)
            Else
                btnFilterSortTitle_Shows.Tag = "ASC"
                btnFilterSortTitle_Shows.Image = My.Resources.asc
                dgvTVShows.Sort(dgvTVShows.Columns("SortedTitle"), System.ComponentModel.ListSortDirection.Ascending)
            End If

            SaveFilter_Shows()
        End If
    End Sub
    ''' <summary>
    ''' sorts the movielist by rating
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>this filter is inverted (DESC first) to get the highest rated title on the top of the list</remarks>
    Private Sub btnFilterSortRating_Movies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterSortRating_Movies.Click
        If dgvMovies.RowCount > 0 Then
            btnFilterSortDateAdded_Movies.Tag = String.Empty
            btnFilterSortDateAdded_Movies.Image = Nothing
            btnFilterSortDateModified_Movies.Tag = String.Empty
            btnFilterSortDateModified_Movies.Image = Nothing
            btnFilterSortReleaseDate_Movies.Tag = String.Empty
            btnFilterSortReleaseDate_Movies.Image = Nothing
            btnFilterSortTitle_Movies.Tag = String.Empty
            btnFilterSortTitle_Movies.Image = Nothing
            btnFilterSortYear_Movies.Tag = String.Empty
            btnFilterSortYear_Movies.Image = Nothing
            If btnFilterSortRating_Movies.Tag.ToString = "DESC" Then
                btnFilterSortRating_Movies.Tag = "ASC"
                btnFilterSortRating_Movies.Image = My.Resources.asc
                dgvMovies.Sort(dgvMovies.Columns("Rating"), System.ComponentModel.ListSortDirection.Ascending)
            Else
                btnFilterSortRating_Movies.Tag = "DESC"
                btnFilterSortRating_Movies.Image = My.Resources.desc
                dgvMovies.Sort(dgvMovies.Columns("Rating"), System.ComponentModel.ListSortDirection.Descending)
            End If

            SaveSorting_Movies()
        End If
    End Sub
    ''' <summary>
    ''' sorts the movielist by release date
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>this filter is inverted (DESC first) to get the highest year title on the top of the list</remarks>
    Private Sub btnFilterSortReleaseDate_Movies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterSortReleaseDate_Movies.Click
        If dgvMovies.RowCount > 0 Then
            btnFilterSortDateAdded_Movies.Tag = String.Empty
            btnFilterSortDateAdded_Movies.Image = Nothing
            btnFilterSortDateModified_Movies.Tag = String.Empty
            btnFilterSortDateModified_Movies.Image = Nothing
            btnFilterSortRating_Movies.Tag = String.Empty
            btnFilterSortRating_Movies.Image = Nothing
            btnFilterSortTitle_Movies.Tag = String.Empty
            btnFilterSortTitle_Movies.Image = Nothing
            btnFilterSortYear_Movies.Tag = String.Empty
            btnFilterSortYear_Movies.Image = Nothing
            If btnFilterSortReleaseDate_Movies.Tag.ToString = "DESC" Then
                btnFilterSortReleaseDate_Movies.Tag = "ASC"
                btnFilterSortReleaseDate_Movies.Image = My.Resources.asc
                dgvMovies.Sort(dgvMovies.Columns("ReleaseDate"), System.ComponentModel.ListSortDirection.Ascending)
            Else
                btnFilterSortReleaseDate_Movies.Tag = "DESC"
                btnFilterSortReleaseDate_Movies.Image = My.Resources.desc
                dgvMovies.Sort(dgvMovies.Columns("ReleaseDate"), System.ComponentModel.ListSortDirection.Descending)
            End If

            SaveSorting_Movies()
        End If
    End Sub
    ''' <summary>
    ''' sorts the movielist by year
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>this filter is inverted (DESC first) to get the highest year title on the top of the list</remarks>
    Private Sub btnFilterSortYear_Movies_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilterSortYear_Movies.Click
        If dgvMovies.RowCount > 0 Then
            btnFilterSortDateAdded_Movies.Tag = String.Empty
            btnFilterSortDateAdded_Movies.Image = Nothing
            btnFilterSortDateModified_Movies.Tag = String.Empty
            btnFilterSortDateModified_Movies.Image = Nothing
            btnFilterSortRating_Movies.Tag = String.Empty
            btnFilterSortRating_Movies.Image = Nothing
            btnFilterSortReleaseDate_Movies.Tag = String.Empty
            btnFilterSortReleaseDate_Movies.Image = Nothing
            btnFilterSortTitle_Movies.Tag = String.Empty
            btnFilterSortTitle_Movies.Image = Nothing
            If btnFilterSortYear_Movies.Tag.ToString = "DESC" Then
                btnFilterSortYear_Movies.Tag = "ASC"
                btnFilterSortYear_Movies.Image = My.Resources.asc
                dgvMovies.Sort(dgvMovies.Columns("Year"), System.ComponentModel.ListSortDirection.Ascending)
            Else
                btnFilterSortYear_Movies.Tag = "DESC"
                btnFilterSortYear_Movies.Image = My.Resources.desc
                dgvMovies.Sort(dgvMovies.Columns("Year"), System.ComponentModel.ListSortDirection.Descending)
            End If

            SaveSorting_Movies()
        End If
    End Sub

    Private Sub btnUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp.Click
        Dim currMainTabTag As Structures.MainTabType = DirectCast(tcMain.SelectedTab.Tag, Structures.MainTabType)
        tcMain.Focus()
        If currMainTabTag.ContentType = Enums.ContentType.Movie Then
            InfoPanelState_Movie = 2
        ElseIf currMainTabTag.ContentType = Enums.ContentType.MovieSet Then
            InfoPanelState_MovieSet = 2
        ElseIf currMainTabTag.ContentType = Enums.ContentType.TV Then
            InfoPanelState_TVShow = 2
        End If
        MoveInfoPanel()
    End Sub

    Private Sub BuildStars(ByVal sinRating As Single)
        Try
            pbStar1.Image = Nothing
            pbStar2.Image = Nothing
            pbStar3.Image = Nothing
            pbStar4.Image = Nothing
            pbStar5.Image = Nothing
            pbStar6.Image = Nothing
            pbStar7.Image = Nothing
            pbStar8.Image = Nothing
            pbStar9.Image = Nothing
            pbStar10.Image = Nothing

            Dim tTip As String = String.Concat(Master.eLang.GetString(245, "Rating:"), String.Format(" {0:N}", sinRating))
            ToolTips.SetToolTip(pbStar1, tTip)
            ToolTips.SetToolTip(pbStar2, tTip)
            ToolTips.SetToolTip(pbStar3, tTip)
            ToolTips.SetToolTip(pbStar4, tTip)
            ToolTips.SetToolTip(pbStar5, tTip)
            ToolTips.SetToolTip(pbStar6, tTip)
            ToolTips.SetToolTip(pbStar7, tTip)
            ToolTips.SetToolTip(pbStar8, tTip)
            ToolTips.SetToolTip(pbStar9, tTip)
            ToolTips.SetToolTip(pbStar10, tTip)

            If sinRating >= 0.5 Then ' if rating is less than .5 out of ten, consider it a 0
                Select Case (sinRating)
                    Case Is <= 0.5
                        pbStar1.Image = My.Resources.starhalf
                        pbStar2.Image = My.Resources.starempty
                        pbStar3.Image = My.Resources.starempty
                        pbStar4.Image = My.Resources.starempty
                        pbStar5.Image = My.Resources.starempty
                        pbStar6.Image = My.Resources.starempty
                        pbStar7.Image = My.Resources.starempty
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 1
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.starempty
                        pbStar3.Image = My.Resources.starempty
                        pbStar4.Image = My.Resources.starempty
                        pbStar5.Image = My.Resources.starempty
                        pbStar6.Image = My.Resources.starempty
                        pbStar7.Image = My.Resources.starempty
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 1.5
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.starhalf
                        pbStar3.Image = My.Resources.starempty
                        pbStar4.Image = My.Resources.starempty
                        pbStar5.Image = My.Resources.starempty
                        pbStar6.Image = My.Resources.starempty
                        pbStar7.Image = My.Resources.starempty
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 2
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.starempty
                        pbStar4.Image = My.Resources.starempty
                        pbStar5.Image = My.Resources.starempty
                        pbStar6.Image = My.Resources.starempty
                        pbStar7.Image = My.Resources.starempty
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 2.5
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.starhalf
                        pbStar4.Image = My.Resources.starempty
                        pbStar5.Image = My.Resources.starempty
                        pbStar6.Image = My.Resources.starempty
                        pbStar7.Image = My.Resources.starempty
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 3
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.starempty
                        pbStar5.Image = My.Resources.starempty
                        pbStar6.Image = My.Resources.starempty
                        pbStar7.Image = My.Resources.starempty
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 3.5
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.starhalf
                        pbStar5.Image = My.Resources.starempty
                        pbStar6.Image = My.Resources.starempty
                        pbStar7.Image = My.Resources.starempty
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 4
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.starempty
                        pbStar6.Image = My.Resources.starempty
                        pbStar7.Image = My.Resources.starempty
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 4.5
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.starhalf
                        pbStar6.Image = My.Resources.starempty
                        pbStar7.Image = My.Resources.starempty
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 5
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.star
                        pbStar6.Image = My.Resources.starempty
                        pbStar7.Image = My.Resources.starempty
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 5.5
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.star
                        pbStar6.Image = My.Resources.starhalf
                        pbStar7.Image = My.Resources.starempty
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 6
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.star
                        pbStar6.Image = My.Resources.star
                        pbStar7.Image = My.Resources.starempty
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 6.5
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.star
                        pbStar6.Image = My.Resources.star
                        pbStar7.Image = My.Resources.starhalf
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 7
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.star
                        pbStar6.Image = My.Resources.star
                        pbStar7.Image = My.Resources.star
                        pbStar8.Image = My.Resources.starempty
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 7.5
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.star
                        pbStar6.Image = My.Resources.star
                        pbStar7.Image = My.Resources.star
                        pbStar8.Image = My.Resources.starhalf
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 8
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.star
                        pbStar6.Image = My.Resources.star
                        pbStar7.Image = My.Resources.star
                        pbStar8.Image = My.Resources.star
                        pbStar9.Image = My.Resources.starempty
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 8.5
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.star
                        pbStar6.Image = My.Resources.star
                        pbStar7.Image = My.Resources.star
                        pbStar8.Image = My.Resources.star
                        pbStar9.Image = My.Resources.starhalf
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 9
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.star
                        pbStar6.Image = My.Resources.star
                        pbStar7.Image = My.Resources.star
                        pbStar8.Image = My.Resources.star
                        pbStar9.Image = My.Resources.star
                        pbStar10.Image = My.Resources.starempty
                    Case Is <= 9.5
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.star
                        pbStar6.Image = My.Resources.star
                        pbStar7.Image = My.Resources.star
                        pbStar8.Image = My.Resources.star
                        pbStar9.Image = My.Resources.star
                        pbStar10.Image = My.Resources.starhalf
                    Case Else
                        pbStar1.Image = My.Resources.star
                        pbStar2.Image = My.Resources.star
                        pbStar3.Image = My.Resources.star
                        pbStar4.Image = My.Resources.star
                        pbStar5.Image = My.Resources.star
                        pbStar6.Image = My.Resources.star
                        pbStar7.Image = My.Resources.star
                        pbStar8.Image = My.Resources.star
                        pbStar9.Image = My.Resources.star
                        pbStar10.Image = My.Resources.star
                End Select
            End If
        Catch ex As Exception
            logger.Error(ex, New StackFrame().GetMethod().Name)
        End Try
    End Sub

    Private Sub bwCleanDB_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bwCleanDB.DoWork
        Dim Args As Structures.ScanOrClean = DirectCast(e.Argument, Structures.ScanOrClean)
        Master.DB.Clean(Args.Movies, Args.MovieSets, Args.TV)
    End Sub

    Private Sub bwCleanDB_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bwCleanDB.RunWorkerCompleted
        SetStatus(String.Empty)
        tspbLoading.Visible = False

        FillList(True, True, True)
    End Sub

    Private Sub bwDownloadPic_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bwDownloadPic.DoWork
        Dim Args As Arguments = DirectCast(e.Argument, Arguments)
        Try

            sHTTP.StartDownloadImage(Args.pURL)

            While sHTTP.IsDownloading
                DoEvents()
                If bwDownloadPic.CancellationPending Then
                    e.Cancel = True
                    sHTTP.Cancel()
                    Return
                End If
                Threading.Thread.Sleep(50)
            End While

            e.Result = New Results With {.Result = sHTTP.Image}
        Catch ex As Exception
            e.Result = New Results With {.Result = Nothing}
            e.Cancel = True
        End Try
    End Sub

    Private Sub bwDownloadPic_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bwDownloadPic.RunWorkerCompleted
        '//
        ' Thread finished: display pic if it was able to get one
        '\\

        pbActLoad.Visible = False

        If e.Cancelled Then
            pbActors.Image = My.Resources.actor_silhouette
        Else
            Dim Res As Results = DirectCast(e.Result, Results)

            If Res.Result IsNot Nothing Then
                pbActors.Image = Res.Result
            Else
                pbActors.Image = My.Resources.actor_silhouette
            End If
        End If
    End Sub

    Private Sub bwLoadImages_Movie_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bwLoadImages_Movie.DoWork
        MainActors.Clear()
        MainBanner.Clear()
        MainCharacterArt.Clear()
        MainClearArt.Clear()
        MainClearLogo.Clear()
        MainDiscArt.Clear()
        MainFanart.Clear()
        MainFanartSmall.Clear()
        MainLandscape.Clear()
        MainPoster.Clear()

        If bwLoadImages_Movie.CancellationPending Then
            e.Cancel = True
            Return
        End If

        currMovie.LoadAllImages(True, False)

        If bwLoadImages_Movie.CancellationPending Then
            e.Cancel = True
            Return
        End If

        If Master.eSettings.GeneralDisplayBanner Then MainBanner = currMovie.ImagesContainer.Banner.ImageOriginal
        If Master.eSettings.GeneralDisplayClearArt Then MainClearArt = currMovie.ImagesContainer.ClearArt.ImageOriginal
        If Master.eSettings.GeneralDisplayClearLogo Then MainClearLogo = currMovie.ImagesContainer.ClearLogo.ImageOriginal
        If Master.eSettings.GeneralDisplayDiscArt Then MainDiscArt = currMovie.ImagesContainer.DiscArt.ImageOriginal
        If Master.eSettings.GeneralDisplayFanart Then MainFanart = currMovie.ImagesContainer.Fanart.ImageOriginal
        If Master.eSettings.GeneralDisplayFanartSmall Then MainFanartSmall = currMovie.ImagesContainer.Fanart.ImageOriginal
        If Master.eSettings.GeneralDisplayLandscape Then MainLandscape = currMovie.ImagesContainer.Landscape.ImageOriginal
        If Master.eSettings.GeneralDisplayPoster Then MainPoster = currMovie.ImagesContainer.Poster.ImageOriginal

        If bwLoadImages_Movie.CancellationPending Then
            e.Cancel = True
            Return
        End If
    End Sub

    Private Sub bwLoadImages_Movie_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bwLoadImages_Movie.RunWorkerCompleted
        If Not e.Cancelled Then
            FillScreenInfoWithImages()
        End If
    End Sub

    Private Sub bwLoadImages_MovieSet_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bwLoadImages_MovieSet.DoWork
        MainActors.Clear()
        MainBanner.Clear()
        MainCharacterArt.Clear()
        MainClearArt.Clear()
        MainClearLogo.Clear()
        MainDiscArt.Clear()
        MainFanart.Clear()
        MainFanartSmall.Clear()
        MainLandscape.Clear()
        MainPoster.Clear()

        If bwLoadImages_MovieSet.CancellationPending Then
            e.Cancel = True
            Return
        End If

        currMovieSet.LoadAllImages(True, False)

        If bwLoadImages_MovieSet.CancellationPending Then
            e.Cancel = True
            Return
        End If

        If Master.eSettings.GeneralDisplayBanner Then MainBanner = currMovieSet.ImagesContainer.Banner.ImageOriginal
        If Master.eSettings.GeneralDisplayClearArt Then MainClearArt = currMovieSet.ImagesContainer.ClearArt.ImageOriginal
        If Master.eSettings.GeneralDisplayClearLogo Then MainClearLogo = currMovieSet.ImagesContainer.ClearLogo.ImageOriginal
        If Master.eSettings.GeneralDisplayDiscArt Then MainDiscArt = currMovieSet.ImagesContainer.DiscArt.ImageOriginal
        If Master.eSettings.GeneralDisplayFanart Then MainFanart = currMovieSet.ImagesContainer.Fanart.ImageOriginal
        If Master.eSettings.GeneralDisplayFanartSmall Then MainFanartSmall = currMovieSet.ImagesContainer.Fanart.ImageOriginal
        If Master.eSettings.GeneralDisplayLandscape Then MainLandscape = currMovieSet.ImagesContainer.Landscape.ImageOriginal
        If Master.eSettings.GeneralDisplayPoster Then MainPoster = currMovieSet.ImagesContainer.Poster.ImageOriginal

        If bwLoadImages_MovieSet.CancellationPending Then
            e.Cancel = True
            Return
        End If
    End Sub

    Private Sub bwLoadImages_MovieSet_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bwLoadImages_MovieSet.RunWorkerCompleted
        If Not e.Cancelled Then
            FillScreenInfoWithImages()
        End If
    End Sub

    Private Sub bwLoadImages_MovieSetMoviePosters_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles bwLoadImages_MovieSetMoviePosters.DoWork
        Dim Posters As New List(Of MovieInSetPoster)

        Try
            If currMovieSet.MoviesInSet IsNot Nothing AndAlso currMovieSet.MoviesInSet.Count > 0 Then
                Try
                    For Each tMovieInSet As MediaContainers.MovieInSet In currMovieSet.MoviesInSet
                        If bwLoadImages_MovieSetMoviePosters.CancellationPending Then
                            e.Cancel = True
                            Return
                        End If

                        Dim ResImg As Image
                        If tMovieInSet.DBMovie.ImagesContainer.Poster.LoadAndCache(Enums.ContentType.Movie, True, True) Then
                            ResImg = tMovieInSet.DBMovie.ImagesContainer.Poster.ImageOriginal.Image
                            ImageUtils.ResizeImage(ResImg, 59, 88, True, Drawing.Color.White.ToArgb())
                            Posters.Add(New MovieInSetPoster With {.MoviePoster = ResImg, .MovieTitle = tMovieInSet.DBMovie.Movie.Title, .MovieYear = tMovieInSet.DBMovie.Movie.Year})
                        Else
                            Posters.Add(New MovieInSetPoster With {.MoviePoster = My.Resources.noposter, .MovieTitle = tMovieInSet.DBMovie.Movie.Title, .MovieYear = tMovieInSet.DBMovie.Movie.Year})
                        End If
                    Next
                Catch ex As Exception
                    logger.Error(ex, New StackFrame().GetMethod().Name)
                    e.Result = New Results With {.MovieInSetPosters = Nothing}
                    e.Cancel = True
                End Try
            End If

            e.Result = New Results With {.MovieInSetPosters = Posters}
        Catch ex As Exception
            logger.Error(ex, New StackFrame().GetMethod().Name)
            e.Result = New Results With {.MovieInSetPosters = Nothing}
            e.Cancel = True
        End Try
    End Sub

    Private Sub bwLoadImages_MovieSetMoviePosters_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bwLoadImages_MovieSetMoviePosters.RunWorkerCompleted
        lvMoviesInSet.Clear()
        ilMoviesInSet.Images.Clear()
        ilMoviesInSet.ImageSize = New Drawing.Size(59, 88)
        ilMoviesInSet.ColorDepth = WinForms.ColorDepth.Depth32Bit
        lvMoviesInSet.Visible = False

        If Not e.Cancelled Then
            Try
                Dim Res As Results = DirectCast(e.Result, Results)

                If Res.MovieInSetPosters IsNot Nothing AndAlso Res.MovieInSetPosters.Count > 0 Then
                    lvMoviesInSet.BeginUpdate()
                    For Each tPoster As Ember_Media_Manager.frmMain.MovieInSetPoster In Res.MovieInSetPosters
                        If tPoster IsNot Nothing Then
                            ilMoviesInSet.Images.Add(tPoster.MoviePoster)
                            lvMoviesInSet.Items.Add(String.Concat(tPoster.MovieTitle, Environment.NewLine, "(", tPoster.MovieYear, ")"), ilMoviesInSet.Images.Count - 1)
                        End If
                    Next
                    lvMoviesInSet.EndUpdate()
                    lvMoviesInSet.Visible = True
                End If
            Catch ex As Exception
                logger.Error(ex, New StackFrame().GetMethod().Name)
            End Try
        End If
    End Sub
#End Region
End Class
