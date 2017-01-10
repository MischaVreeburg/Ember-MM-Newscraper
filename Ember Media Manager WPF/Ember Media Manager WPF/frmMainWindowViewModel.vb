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
Imports System.Windows.Forms

Public Class frmMainWindowViewModel
    Implements INotifyPropertyChanged

    Private WithEvents _FrmMainWIndowModel1 As frmMainWindowModel
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged


    Public Property FrmMainWIndowModel As frmMainWindowModel
        Get
            Return _FrmMainWIndowModel1
        End Get
        Set
            _FrmMainWIndowModel1 = Value
        End Set
    End Property

    Protected Sub RaisePropertyChangedEvent(ByVal propertyName As String)
        If Me.PropertyChangedEvent IsNot Nothing Then
            RaiseEvent PropertyChanged(Me,
                                       New PropertyChangedEventArgs(propertyName))
        End If
    End Sub

    Public Sub New()
        If FrmMainWIndowModel Is Nothing Then
            FrmMainWIndowModel = New frmMainWindowModel
        End If
    End Sub

    Public Sub CallRaisePropertyChangedEvent(ByVal propName As String) Handles _FrmMainWIndowModel1.CallRaisePropertyChangedEvent
        RaisePropertyChangedEvent(propName)
    End Sub

    Public Property pbFanArtImage As Bitmap
        Get
            Return FrmMainWIndowModel.pbFanArtImage
        End Get
        Set(value As Bitmap)
            FrmMainWIndowModel.pbFanArtImage = value
            RaisePropertyChangedEvent("pbFanArtImage")
        End Set
    End Property
End Class
