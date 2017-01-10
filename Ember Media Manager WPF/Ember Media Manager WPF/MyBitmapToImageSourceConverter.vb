Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Resources
Imports System.Runtime.InteropServices
Imports System.Windows.Interop

Public Class MyBitmapToImageSourceConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim MyBitMap As System.Drawing.Bitmap = TryCast(value, System.Drawing.Bitmap)
        If MyBitMap IsNot Nothing Then
            Dim memoryStream = New MemoryStream()
            MyBitMap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Bmp)
            memoryStream.Position = 0
            Dim bitmapImage = New BitmapImage()
            bitmapImage.BeginInit()
            bitmapImage.StreamSource = memoryStream
            bitmapImage.CacheOption = BitmapCacheOption.OnLoad
            bitmapImage.EndInit()

            Return bitmapImage
        End If
        Throw New ArgumentException("value is not of type System.Drawing.Bitmap")
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function

End Class

Public Class MyIconToImageSourceConverter
    Implements IValueConverter

    <DllImport("gdi32.dll", SetLastError:=True)>
    Private Shared Function DeleteObject(hObject As IntPtr) As Boolean
    End Function
    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim MyIcon As System.Drawing.Icon = TryCast(value, System.Drawing.Icon)
        If MyIcon IsNot Nothing Then
            Dim bitmap As Bitmap = MyIcon.ToBitmap()
            Dim hBitmap As IntPtr = bitmap.GetHbitmap()

            Dim wpfBitmap As ImageSource = Imaging.CreateBitmapSourceFromHBitmap(hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())

            If Not DeleteObject(hBitmap) Then
                Throw New Win32Exception()
            End If

            Return wpfBitmap
        End If
        Throw New ArgumentException("value is not of type System.Drawing.Icon")

    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function

End Class

Friend NotInheritable Class IconUtilities
    Private Sub New()
    End Sub
    <DllImport("gdi32.dll", SetLastError:=True)>
    Private Shared Function DeleteObject(hObject As IntPtr) As Boolean
    End Function

    ' <System.Runtime.CompilerServices.Extension>
    Public Shared Function ToImageSource(icon As Icon) As ImageSource
        Dim bitmap As Bitmap = icon.ToBitmap()
        Dim hBitmap As IntPtr = bitmap.GetHbitmap()

        Dim wpfBitmap As ImageSource = Imaging.CreateBitmapSourceFromHBitmap(hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())

        If Not DeleteObject(hBitmap) Then
            Throw New Win32Exception()
        End If

        Return wpfBitmap
    End Function
End Class

