Imports System.Windows.Forms

Class frmMainWindow
    Public Property tmrLoad_TVEpisode As New Timer With {.Interval = 100}
    Public Property tmrLoad_TVSeason As New Timer With {.Interval = 100}
    Public Property tmrWait_TVSeason As New Timer With {.Interval = 100}
    Public Property tmrLoad_TVShow As New Timer With {.Interval = 100}
    Public Property tmrWait_MovieSet As New Timer With {.Interval = 250}
    Public Property tmrLoad_MovieSet As New Timer With {.Interval = 100}
    Public Property tmrWait_TVShow As New Timer With {.Interval = 250}
    Public Property tmrWait_TVEpisode As New Timer With {.Interval = 250}
    Public Property tmrWait_Movie As New Timer With {.Interval = 250}
    Public Property tmrLoad_Movie As New Timer With {.Interval = 100}
    Public Property tmrSearch_MovieSets As New Timer With {.Interval = 250}
    Public Property tmrSearchWait_MovieSets As New Timer With {.Interval = 250}
    Public Property tmrSearch_Movies As New Timer With {.Interval = 250}
    Public Property tmrSearchWait_Movies As New Timer With {.Interval = 250}
    Public Property tmrKeyBuffer As New Timer With {.Interval = 2000}
    Public Property tmrAppExit As New Timer With {.Interval = 1000}
    Public Property tmrSearchWait_Shows As New Timer With {.Interval = 250}
    Public Property tmrSearch_Shows As New Timer With {.Interval = 250}
    Public Property tmrRunTasks As New Timer With {.Interval = 100}



    Public Property ilColumnIcons As New ImageList
    Public Property ilMoviesInSet As New ImageList With {.ColorDepth = ColorDepth.Depth8Bit, .ImageSize = New System.Drawing.Size(16, 16)}
End Class
