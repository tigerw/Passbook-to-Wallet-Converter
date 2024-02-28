Imports System.IO.Compression
Imports Windows.Data.Json
Imports Windows.Storage.Streams

Module PassConverterCommon

	Async Function GetPassFromPassBundle(PassBundle As ZipArchive) As Task(Of JsonObject)
		Return JsonObject.Parse(Await ReadZipArchiveEntryIntoString(PassBundle.GetEntry("pass.json")))
	End Function

	''' <summary>
	''' Synchronously read a zip entry fully into memory (since the zip stream is forward only) and converts it to a random accesss stream.
	''' </summary>
	''' <remarks>Don't call this from the UI thread.</remarks>
	''' <param name="ZipArchiveEntry">The archive entry.</param>
	''' <returns>A Windows Runtime random access stream.</returns>
	Function ReadZipArchiveEntryIntoRandomAccessStream(ZipArchiveEntry As ZipArchiveEntry) As IRandomAccessStreamReference
		If ZipArchiveEntry Is Nothing Then
			Return Nothing
		End If

		Using ArchiveEntryStream = ZipArchiveEntry.Open()
			Dim MemoryStream = New InMemoryRandomAccessStream()
			RandomAccessStream.CopyAsync(ArchiveEntryStream.AsInputStream(), MemoryStream.GetOutputStreamAt(0)).AsTask().Wait()
			Return RandomAccessStreamReference.CreateFromStream(MemoryStream)
		End Using
	End Function

	Private Async Function ReadZipArchiveEntryIntoString(ZipArchiveEntry As ZipArchiveEntry) As Task(Of String)
		Using ArchiveEntryReader = New StreamReader(ZipArchiveEntry.Open())
			Return Await ArchiveEntryReader.ReadToEndAsync()
		End Using
	End Function

End Module
