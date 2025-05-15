Imports System.Media

' OP2SheetsReader
' https://github.com/leviathan400/OP2SheetsReader
'
' Outpost 2 sheets reader. Open and edit sheets .txt files.
'
'
' Outpost 2: Divided Destiny is a real-time strategy video game released in 1997.

Public Class fMain

    'building.txt, mines.txt, morale.txt, space.txt, vehicles.txt and weapons.txt

    Public ApplicationName As String = "OP2SheetsReader"
    Public Version As String = "0.6"

    Private FormTitle As String = "Outpost 2 Sheets Reader"
    Private CurrentSheetsFile As String = Nothing

    Private OriginalHeaderLine As String
    Private AdditionalHeaderLines As New List(Of String)
    Private CommentLines As New Dictionary(Of Integer, String)  ' Line position -> comment text
    Private EmptyLinePositions As New List(Of Integer)
    Private FooterLine As String
    Private OriginalFilePath As String
    Private FileEncoding As System.Text.Encoding

    Private Sub fMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Debug.WriteLine("--- " & ApplicationName & " Started ---")
        Me.Icon = My.Resources.compatibility
        Me.Text = FormTitle

        btnOpen.Text = "&Open File"
        btnSave.Text = "Save File"
        btnSave.Enabled = False

        Dim player As New SoundPlayer(My.Resources.mine_1)
        player.Play()
    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        ' Open a sheets file
        OpenSheetsFile()

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        ' Save edits to sheets file
        SaveSheetsFile()

    End Sub

    Private Sub OpenSheetsFile()
        ' Create and configure OpenFileDialog
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Title = "Select TXT file to open"
        openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        openFileDialog.DefaultExt = "txt"
        openFileDialog.CheckFileExists = True

        ' Show dialog and process file if user didn't cancel
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Try
                ' Clear existing data
                DataGridView.Columns.Clear()
                DataGridView.Rows.Clear()

                ' Reset tracking variables
                OriginalHeaderLine = ""
                AdditionalHeaderLines.Clear()
                CommentLines.Clear()
                EmptyLinePositions.Clear()
                FooterLine = ""

                ' Store the full file path
                OriginalFilePath = openFileDialog.FileName

                ' Detect file encoding by reading the first few bytes
                Dim fileInfo As New IO.FileInfo(OriginalFilePath)
                Dim encodingBytes(4) As Byte

                Using fs As New IO.FileStream(OriginalFilePath, IO.FileMode.Open, IO.FileAccess.Read)
                    fs.Read(encodingBytes, 0, 4)
                End Using

                ' Determine encoding based on BOM
                If encodingBytes(0) = &HEF AndAlso encodingBytes(1) = &HBB AndAlso encodingBytes(2) = &HBF Then
                    FileEncoding = System.Text.Encoding.UTF8
                ElseIf encodingBytes(0) = &HFF AndAlso encodingBytes(1) = &HFE Then
                    FileEncoding = System.Text.Encoding.Unicode
                ElseIf encodingBytes(0) = &HFE AndAlso encodingBytes(1) = &HFF Then
                    FileEncoding = System.Text.Encoding.BigEndianUnicode
                ElseIf encodingBytes(0) = &H0 AndAlso encodingBytes(1) = &H0 AndAlso
                  encodingBytes(2) = &HFE AndAlso encodingBytes(3) = &HFF Then
                    FileEncoding = System.Text.Encoding.UTF32
                Else
                    ' Default to ANSI/ASCII
                    FileEncoding = System.Text.Encoding.Default
                End If

                ' Read the file line by line to preserve structure
                Dim allLines As String() = IO.File.ReadAllLines(OriginalFilePath, FileEncoding)

                ' VALIDATION: Check if file is empty
                If allLines.Length = 0 Then
                    MessageBox.Show("The file is empty and cannot be loaded.", "Invalid File Format", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                ' Process all lines to identify structure
                Dim headerProcessed As Boolean = False
                Dim dataStartLine As Integer = -1
                Dim dataEndLine As Integer = -1
                Dim lineIndex As Integer = 0
                Dim hasTabDelimitedData As Boolean = False

                ' Process all lines to identify structure
                For i As Integer = 0 To allLines.Length - 1
                    Dim line As String = allLines(i)

                    ' Check for empty lines
                    If String.IsNullOrEmpty(line) Then
                        EmptyLinePositions.Add(i)
                        Continue For
                    End If

                    ' Check for comment lines
                    If line.TrimStart().StartsWith(";") Then
                        CommentLines.Add(i, line)
                        Continue For
                    End If

                    ' Process header (first non-empty, non-comment line)
                    If Not headerProcessed Then
                        OriginalHeaderLine = line
                        headerProcessed = True
                        dataStartLine = i + 1

                        ' VALIDATION: Check if header has tabs
                        If line.Contains(ControlChars.Tab) Then
                            hasTabDelimitedData = True
                        End If

                        Continue For
                    End If

                    ' Process additional header lines (tab-prefixed)
                    If line.StartsWith(ControlChars.Tab) Then
                        AdditionalHeaderLines.Add(line)
                        dataStartLine = i + 1
                        Continue For
                    End If

                    ' Check for potential footer line (like "Garbage so checksum matches...")
                    If i = allLines.Length - 1 AndAlso line.Contains("checksum") Then
                        FooterLine = line
                        dataEndLine = i - 1
                    End If

                    ' VALIDATION: Check additional lines for tab-delimited format
                    If Not hasTabDelimitedData AndAlso line.Contains(ControlChars.Tab) Then
                        hasTabDelimitedData = True
                    End If
                Next

                ' VALIDATION: Check if this appears to be a valid sheets file
                If Not headerProcessed Then
                    MessageBox.Show("This file does not appear to be a valid Outpost 2 sheets file. No header row found.", "Invalid File Format", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                If Not hasTabDelimitedData Then
                    MessageBox.Show("This file does not appear to be a valid Outpost 2 sheets file. No tab-delimited data found.", "Invalid File Format", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                ' If no explicit footer was found, data ends at the last line
                If dataEndLine = -1 Then
                    dataEndLine = allLines.Length - 1
                End If

                ' Now process the actual data for the DataGridView
                ' First, process header
                If Not String.IsNullOrEmpty(OriginalHeaderLine) Then
                    ' Split by tabs
                    Dim headers As String() = OriginalHeaderLine.Split(ControlChars.Tab)

                    ' VALIDATION: Check if there are enough columns
                    If headers.Length < 2 Then
                        MessageBox.Show("This file does not have enough columns to be a valid Outpost 2 sheets file.", "Invalid File Format", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If

                    ' Add columns to DataGridView
                    For Each header As String In headers
                        DataGridView.Columns.Add(header, header)
                    Next

                    ' Add hidden column for original row order
                    Dim originalOrderColumn As New DataGridViewTextBoxColumn()
                    originalOrderColumn.Name = "OriginalRowIndex"
                    originalOrderColumn.Visible = False
                    DataGridView.Columns.Add(originalOrderColumn)

                    ' Set column widths
                    For Each col As DataGridViewColumn In DataGridView.Columns
                        If col.Name <> "OriginalRowIndex" Then
                            col.Width = 120
                        End If
                    Next

                    ' Process additional header lines
                    For Each additionalHeader In AdditionalHeaderLines
                        Dim additionalHeaders As String() = additionalHeader.Split(ControlChars.Tab)

                        For i As Integer = 0 To Math.Min(additionalHeaders.Length - 1, DataGridView.Columns.Count - 1)
                            If i > 0 AndAlso Not String.IsNullOrEmpty(additionalHeaders(i)) Then
                                ' Append to existing header
                                If DataGridView.Columns(i).Name <> "OriginalRowIndex" Then
                                    DataGridView.Columns(i).HeaderText &= " " & additionalHeaders(i)
                                End If
                            End If
                        Next
                    Next

                    ' Process data rows
                    Dim rowIndex As Integer = 0
                    Dim rowsAdded As Integer = 0

                    For i As Integer = dataStartLine To dataEndLine
                        ' Skip tracked empty or comment lines
                        If EmptyLinePositions.Contains(i) OrElse CommentLines.ContainsKey(i) Then
                            Continue For
                        End If

                        Dim line As String = allLines(i)

                        ' Skip additional headers we might encounter
                        If line.StartsWith(ControlChars.Tab) Then
                            Continue For
                        End If

                        ' Process data row
                        Dim values As String() = line.Split(ControlChars.Tab)

                        ' Add row
                        DataGridView.Rows.Add()

                        ' Fill in data
                        For j As Integer = 0 To Math.Min(values.Length - 1, DataGridView.Columns.Count - 1)
                            If DataGridView.Columns(j).Name <> "OriginalRowIndex" Then
                                DataGridView.Rows(rowIndex).Cells(j).Value = values(j)
                            End If
                        Next

                        ' Store original index in the hidden column
                        DataGridView.Rows(rowIndex).Cells("OriginalRowIndex").Value = rowIndex

                        rowIndex += 1
                        rowsAdded += 1
                    Next

                    ' VALIDATION: Check if any rows were added
                    If rowsAdded = 0 Then
                        MessageBox.Show("No data rows found in the file.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        ' Still continue, as this might be user intention
                    End If
                End If

                ' Make sure column headers are visible
                DataGridView.ColumnHeadersVisible = True

                ' Store current file info and update title
                CurrentSheetsFile = IO.Path.GetFileName(OriginalFilePath)
                Me.Text = FormTitle & " - " & CurrentSheetsFile

                btnSave.Enabled = True

                Debug.WriteLine("File Opened: " & CurrentSheetsFile)

            Catch ex As Exception
                MessageBox.Show("Error reading file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub SaveSheetsFile()
        Try
            ' Create backup of original file first
            Dim backupFilePath As String = OriginalFilePath & ".bak"

            ' Remove existing backup if it exists
            If IO.File.Exists(backupFilePath) Then
                IO.File.Delete(backupFilePath)
            End If

            ' Create the backup
            IO.File.Copy(OriginalFilePath, backupFilePath)
            'Debug.WriteLine("Backup created: " & backupFilePath)

            ' Create a temporary list of lines to write
            Dim linesToWrite As New List(Of String)

            ' Add the original header
            linesToWrite.Add(OriginalHeaderLine)

            ' Add additional header lines
            For Each line In AdditionalHeaderLines
                linesToWrite.Add(line)
            Next

            ' Create a list to hold rows in their original order
            Dim sortedRows As New List(Of DataGridViewRow)

            ' Add all non-empty rows to the list
            For i As Integer = 0 To DataGridView.Rows.Count - 1
                If Not DataGridView.Rows(i).IsNewRow Then
                    sortedRows.Add(DataGridView.Rows(i))
                End If
            Next

            ' Sort the rows based on the original row index 
            sortedRows.Sort(Function(a, b)
                                Dim indexA As Integer = -1
                                Dim indexB As Integer = -1

                                If a.Cells("OriginalRowIndex").Value IsNot Nothing Then
                                    Integer.TryParse(a.Cells("OriginalRowIndex").Value.ToString(), indexA)
                                End If

                                If b.Cells("OriginalRowIndex").Value IsNot Nothing Then
                                    Integer.TryParse(b.Cells("OriginalRowIndex").Value.ToString(), indexB)
                                End If

                                Return indexA.CompareTo(indexB)
                            End Function)

            ' Process rows in their original order
            For Each row As DataGridViewRow In sortedRows
                Dim rowData As New List(Of String)

                ' Get values from each cell (skip the hidden OriginalRowIndex column)
                For j As Integer = 0 To DataGridView.Columns.Count - 1
                    If DataGridView.Columns(j).Name <> "OriginalRowIndex" Then
                        Dim cellValue As Object = row.Cells(j).Value
                        rowData.Add(If(cellValue Is Nothing, "", cellValue.ToString()))
                    End If
                Next

                ' Join with tabs and add to lines
                linesToWrite.Add(String.Join(ControlChars.Tab, rowData.ToArray()))
            Next

            ' Add footer if exists
            If Not String.IsNullOrEmpty(FooterLine) Then
                linesToWrite.Add(FooterLine)
            End If

            ' Insert empty lines at their original positions
            ' We need to be careful to adjust positions based on our current line count
            Dim currentLineCount As Integer = linesToWrite.Count

            For Each emptyLinePos In EmptyLinePositions.OrderBy(Function(pos) pos)
                If emptyLinePos < currentLineCount Then
                    linesToWrite.Insert(emptyLinePos, "")
                    currentLineCount += 1
                Else
                    linesToWrite.Add("")
                End If
            Next

            ' Insert comment lines at their original positions
            For Each commentEntry In CommentLines.OrderBy(Function(entry) entry.Key)
                Dim commentPos As Integer = commentEntry.Key
                Dim commentText As String = commentEntry.Value

                If commentPos < currentLineCount Then
                    linesToWrite.Insert(commentPos, commentText)
                    currentLineCount += 1
                Else
                    linesToWrite.Add(commentText)
                End If
            Next

            ' Write to the file with original encoding
            IO.File.WriteAllLines(OriginalFilePath, linesToWrite.ToArray(), FileEncoding)

            MessageBox.Show("File saved successfully." & vbNewLine & "A backup was created as " &
                      IO.Path.GetFileName(backupFilePath), "Save Complete",
                      MessageBoxButtons.OK, MessageBoxIcon.Information)

            Debug.WriteLine("File Saved: " & CurrentSheetsFile)

        Catch ex As Exception
            MessageBox.Show("Error saving file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

End Class
