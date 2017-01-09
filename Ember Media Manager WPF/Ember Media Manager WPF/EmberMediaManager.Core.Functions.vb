Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Windows.Controls.Primitives
Imports System.Math
Imports System.Threading

Public Class EmberMediaManagerCoreFunctions

End Class

Public Module DataGridExtension

    <Extension()>
    Public Function GetRowIndex(ByRef dg As DataGrid, ByVal dgci As DataGridCellInfo) As Integer
        Dim Dgrow As DataGridRow = DirectCast(dg.ItemContainerGenerator.ContainerFromItem(dgci.Item), DataGridRow)
        Return Dgrow.GetIndex()
    End Function

    <Extension()>
    Public Function GetColIndex(ByRef dg As DataGrid, ByVal dgci As DataGridCellInfo) As Integer
        Return dgci.Column.DisplayIndex
    End Function

    <Extension()>
    Public Function GetDataGridCell(ByRef dg As DataGrid, ByVal row As Integer, ByVal column As Integer) As DataGridCell
        Dim RowContainer As DataGridRow = dg.GetDataGridRow(row)

        If RowContainer IsNot Nothing Then
            Dim Presenter As DataGridCellsPresenter = RowContainer.GetVisualChild(Of DataGridCellsPresenter)()

            ' try to get the cell but it may possibly be virtualized
            Dim Cell As DataGridCell = DirectCast(Presenter.ItemContainerGenerator.ContainerFromIndex(column), DataGridCell)
            If Cell Is Nothing Then
                ' now try to bring into view and retreive the cell
                dg.UpdateLayout()
                dg.ScrollIntoView(RowContainer, dg.Columns(column))
                Cell = DirectCast(Presenter.ItemContainerGenerator.ContainerFromIndex(column), DataGridCell)
            End If
            Return Cell
        End If
        Return Nothing
    End Function

    <Extension()>
    Public Function GetDataGridRow(ByRef dg As DataGrid, ByVal index As Integer) As DataGridRow
        Dim Row As DataGridRow = DirectCast(dg.ItemContainerGenerator.ContainerFromIndex(index), DataGridRow)
        If Row Is Nothing Then
            ' may be virtualized, bring into view and try again
            dg.UpdateLayout()
            dg.ScrollIntoView(dg.Items(index))
            Row = DirectCast(dg.ItemContainerGenerator.ContainerFromIndex(index), DataGridRow)
        End If
        Return Row
    End Function

    <Extension()>
    Public Function AsSystemDrawingColor(ByRef color As System.Windows.Media.Color) As System.Drawing.Color
        Return System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B)
    End Function

    <Extension()>
    Public Function GetColumnHeaderFromColumn(ByRef column As DataGridColumn) As DataGridColumnHeader

        Dim DefaultColumnHeader As New DataGridColumnHeader

        Dim DataGridParent As DataGrid = column.GetDataGridParent
        Dim ColumnHeaders As List(Of DataGridColumnHeader) = DataGridParent.GetVisualChildCollection(Of DataGridColumnHeader)()

        'Dim returnValue As DataGridColumnHeader
        For Each ColumnHeader As DataGridColumnHeader In ColumnHeaders
            If (ColumnHeader.Column Is column) Then
                Return ColumnHeader
            End If
        Next
        Return DefaultColumnHeader
    End Function

    <Extension()>
    Public Function GetDataGridParent(ByRef column As DataGridColumn) As DataGrid

        Dim PropertyInfo As Reflection.PropertyInfo = column.GetType().GetProperty("DataGridOwner", BindingFlags.Instance Or BindingFlags.NonPublic)
        Return CType(PropertyInfo.GetValue(column, Nothing), DataGrid)

    End Function

    <Extension()>
    Public Function PrintWidth(ByRef column As DataGridColumn, ByVal stringFont As System.Drawing.Font) As Single

        Dim HeaderString As String = column.Header.ToString
        Dim ColumnString As String = String.Empty
        Dim MeasureString As String

        For I As Integer = 0 To column.GetDataGridParent.Items.Count - 1
            If column.GetDataGridParent.GetDataGridCell(I, column.GetDataGridParent.Columns.IndexOf(column)).Content.GetType Is GetType(String) Then
                If ColumnString.Length < column.GetDataGridParent.GetDataGridCell(I, column.GetDataGridParent.Columns.IndexOf(column)).Content.ToString.Length Then
                    ColumnString = column.GetDataGridParent.GetDataGridCell(I, column.GetDataGridParent.Columns.IndexOf(column)).Content.ToString
                End If
            ElseIf column.GetDataGridParent.GetDataGridCell(I, column.GetDataGridParent.Columns.IndexOf(column)).Content.GetType Is GetType(TextBlock) Then
                If ColumnString.Length < CType(column.GetDataGridParent.GetDataGridCell(I, column.GetDataGridParent.Columns.IndexOf(column)).Content, TextBlock).Text.ToString.Length Then
                    ColumnString = CType(column.GetDataGridParent.GetDataGridCell(I, column.GetDataGridParent.Columns.IndexOf(column)).Content, TextBlock).Text.ToString
                End If
            End If

        Next

        MeasureString = If(HeaderString.Length >= ColumnString.Length, HeaderString, ColumnString)

        ' Measure string.
        Dim StringSize As SizeF
        Dim Window As New System.Windows.Forms.Form
        Dim G As Graphics = Window.CreateGraphics()

        StringSize = G.MeasureString(MeasureString, stringFont)

        Return StringSize.Width

    End Function

    <Extension()>
    Public Function Item(ByRef dg As DataGrid, ByVal columnIndex As Integer, ByVal rowIndex As Integer) As DataGridCell
        Return dg.GetDataGridCell(rowIndex, columnIndex)
    End Function

    <Extension()>
    Public Function Item(ByRef dg As DataGrid, ByVal columnName As String, ByVal rowIndex As Integer) As DataGridCell
        Dim columnindex As Integer = dg.Columns.IndexOf(dg.Columns.FirstOrDefault(Function(c) CStr(c.Header) = columnName))
        Return dg.GetDataGridCell(rowIndex, columnindex)
    End Function


    <Extension()>
    Public Function SelectedRows(ByRef dg As DataGrid) As IList(Of DataGridRow)
        Dim retList = New List(Of DataGridRow)
        If dg.SelectionMode = DataGridSelectionMode.Extended Then
            For Each r In dg.SelectedItems
                retList.Add(TryCast(r, DataGridRow))
            Next
        Else
            retList.Add(TryCast(dg.SelectedItem, DataGridRow))
        End If
        Return retList
    End Function

    <Extension()>
    Public Function Index(ByRef dgr As DataGridRow) As Integer
        Dim parent As DataGrid = TryCast(dgr.Parent, DataGrid)
        Return If(parent IsNot Nothing, parent.Items.IndexOf(dgr), -1)
    End Function



    <Extension()>
    Public Function SelectedTab(ByRef tc As TabControl) As Controls.TabItem
        Return TryCast(tc.SelectedItem, TabItem)
    End Function

    <Extension()>
    Public Function Image(ByRef im As Controls.Image) As ImageSource
        Return im.Source
    End Function
End Module
Public Module ControlExtension
    <Extension()>
    Public Function GetVisualChild(Of T As Visual)(ByVal parent As Visual) As T
        Dim Child As T = Nothing
        Dim NumVisuals As Integer = VisualTreeHelper.GetChildrenCount(parent)
        For I As Integer = 0 To NumVisuals - 1
            Dim V As Visual = DirectCast(VisualTreeHelper.GetChild(parent,
                                                                   I), Visual)
            Child = TryCast(V, T)
            If Child Is Nothing Then
                Child = V.GetVisualChild(Of T)()
            End If
            If Child IsNot Nothing Then
                Exit For
            End If
        Next
        Return Child
    End Function

    ''' <summary>
    ''' Finds a parent of a given control/item on the visual tree.
    ''' </summary>
    ''' <typeparam name="T">Type of Parent</typeparam>
    ''' <param name="child">Child whose parent is queried</param>
    ''' <returns>Returns the first parent item that matched the type (T), if no match found then it will return null</returns>
    <Extension>
    Public Function TryFindParent(Of T As DependencyObject)(child As DependencyObject) As T
        Dim ParentObject As DependencyObject = VisualTreeHelper.GetParent(child)
        If ParentObject Is Nothing Then
            Return Nothing
        End If
        Dim Parent As T = TryCast(ParentObject, T)
        Return If(Not (Parent Is Nothing), Parent, ParentObject.TryFindParent(Of T)())
    End Function

    <Extension>
    Public Function GetVisualParent(Of T As Visual)(element As Visual) As T
        Dim Parent As Visual = element
        While Not (Parent Is Nothing)
            Dim CorrectlyTyped As T = TryCast(Parent, T)
            If Not (CorrectlyTyped Is Nothing) Then
                Return CorrectlyTyped
            End If
            Parent = TryCast(VisualTreeHelper.GetParent(Parent), Visual)
        End While
        Return Nothing
    End Function

    <Extension()>
    Public Function GetVisualChildCollection(Of T As Visual)(ByVal parent As Visual) As List(Of T)

        Dim VisualCollection As New List(Of T)
        parent.GetVisualChildCollection(VisualCollection)
        Return VisualCollection

    End Function

    <Extension()>
    Public Function GetVisualChildByName(Of T As Visual)(ByVal parent As Visual, ByVal nameToFind As String) As T
        Dim Count As Integer = VisualTreeHelper.GetChildrenCount(parent)
        For I As Integer = 0 To Count - 1

            Dim Child As DependencyObject = VisualTreeHelper.GetChild(parent, I)
            If (Child.GetType Is GetType(T)) Then
                If CType(Child, FrameworkElement).Name = nameToFind Then
                    Return CType(Child, T)
                End If
            Else
                Return Nothing
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Finds a Child of a given item in the visual tree.
    ''' </summary>
    ''' <param name="parent">A direct parent of the queried item.</param>
    ''' <typeparam name="T">The type of the queried item.</typeparam>
    ''' <param name="nameToFind">x:Name or Name of child. </param>
    ''' <returns>The first parent item that matches the submitted type parameter.
    ''' If not matching item can be found,
    ''' a null parent is being returned.</returns>
    <Extension()>
    Public Function FindChild(Of T As DependencyObject)(ByVal parent As DependencyObject, ByVal nameToFind As String) As T
        ' Confirm parent and childName are valid.
        If parent Is Nothing Then
            Return Nothing
        End If

        Dim FoundChild As T = Nothing

        Dim ChildrenCount As Integer = VisualTreeHelper.GetChildrenCount(parent)
        Dim I As Integer = 0
        While I < ChildrenCount
            Dim Child = VisualTreeHelper.GetChild(parent, I)
            ' If the child is not of the request child type child
            Dim ChildType As T = TryCast(Child, T)
            If ChildType Is Nothing Then
                ' recursively drill down the tree
                FoundChild = Child.FindChild(Of T)(nameToFind)

                ' If the child is found, break so we do not overwrite the found child.
                If Not (FoundChild Is Nothing) Then
                    Exit While
                End If
            ElseIf Not String.IsNullOrEmpty(nameToFind) Then
                Dim FrameworkElement = TryCast(Child, FrameworkElement)
                ' If the child's name is set for search
                If Not (FrameworkElement Is Nothing) AndAlso FrameworkElement.Name = nameToFind Then
                    ' if the child's name is of the request name
                    FoundChild = DirectCast(Child, T)
                    Exit While
                End If
            Else
                ' child element found.
                FoundChild = DirectCast(Child, T)
                Exit While
            End If
            Max(Interlocked.Increment(I), I - 1)
        End While

        Return FoundChild
    End Function

    <Extension()>
    Private Sub GetVisualChildCollection(Of T As Visual)(ByVal parent As DependencyObject, ByVal visualCollection As List(Of T))

        Dim Count As Integer = VisualTreeHelper.GetChildrenCount(parent)
        For I As Integer = 0 To Count - 1

            Dim Child As DependencyObject = VisualTreeHelper.GetChild(parent, I)
            If (Child.GetType Is GetType(T)) Then

                visualCollection.Add(CType(Child, T))

            ElseIf (Child IsNot Nothing) Then

                Child.GetVisualChildCollection(visualCollection)
            End If
        Next

    End Sub

    <Extension()>
    Public Function Focused(ByVal control As FrameworkElement) As Boolean
        Return control.IsFocused
    End Function

    <Extension()>
    Public Sub Visible(ByVal control As FrameworkElement, ByVal SetVisibile As Boolean)
        If SetVisibile Then
            control.Visibility = Visibility.Visible
        Else
            control.Visibility = Visibility.Collapsed
        End If
    End Sub

End Module
