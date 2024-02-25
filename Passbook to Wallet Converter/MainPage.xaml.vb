Imports System.IO.Compression
Imports Windows.ApplicationModel.Wallet
Imports Windows.Data.Json
Imports Windows.Storage

Public NotInheritable Class MainPage
	Inherits Page

	Private OpenPicker As Pickers.FileOpenPicker
	Private SavePicker As Pickers.FileSavePicker

	Public Sub New()
		' This call is required by the designer.
		InitializeComponent()

		OpenPicker = New Pickers.FileOpenPicker()
		OpenPicker.FileTypeFilter.Add(".pkpass")

		SavePicker = New Pickers.FileSavePicker()
		SavePicker.FileTypeChoices.Add("Microsoft Wallet", {".xml"})
		SavePicker.SuggestedStartLocation = Pickers.PickerLocationId.DocumentsLibrary
	End Sub

	Private Async Sub ConvertToFile_Clicked(sender As Object, e As RoutedEventArgs)
		Try
			Await PromptForAndOpenPassbookThenConvertIntoFile()
		Catch Oops As Exception
			ShowErrorDialogue(Oops)
		End Try
	End Sub

	Private Async Sub ImportIntoWallet_Clicked(sender As Object, e As RoutedEventArgs)
		Try
			Await PromptForAndOpenPassbookThenImportIntoWallet()
		Catch Oops As Exception
			ShowErrorDialogue(Oops)
		End Try
	End Sub

	Private Async Function PromptForAndOpenPassbookThen(Functor As Func(Of JsonObject, Task)) As Task
		Dim PickedFile = Await OpenPicker.PickSingleFileAsync()
		If Not PickedFile Is Nothing Then
			Using PickedFileStream = Await PickedFile.OpenStreamForReadAsync()
				Using PkPassArchive = New ZipArchive(PickedFileStream)
					Using ArchiveEntryReader = New StreamReader(PkPassArchive.GetEntry("pass.json").Open())
						Await Functor(JsonObject.Parse(Await ArchiveEntryReader.ReadToEndAsync()))
					End Using
				End Using
			End Using
		End If
	End Function

	Private Async Function PromptForAndOpenPassbookThenConvertIntoFile() As Task
		Await PromptForAndOpenPassbookThen(
			Async Function(PassStructure As JsonObject) As Task
				Dim SavePickedFile = Await SavePicker.PickSaveFileAsync()
				If Not SavePickedFile Is Nothing Then
					Dim WalletItem = PassToWalletItemXmlConverter.PassToWalletItem(PassStructure)
					Using SavePickedFileStream = Await SavePickedFile.OpenStreamForWriteAsync()
						WalletItem.Save(SavePickedFileStream)
					End Using
				End If
			End Function
		)
	End Function

	Private Async Function PromptForAndOpenPassbookThenImportIntoWallet() As Task
		Dim Store = Await WalletManager.RequestStoreAsync()

		Await PromptForAndOpenPassbookThen(
			Async Function(PassStructure As JsonObject) As Task
				Dim PassId = PassToWalletItemConverter.GetPassId(PassStructure)
				Dim WalletItem = Await Store.GetWalletItemAsync(PassId)

				If WalletItem Is Nothing Then
					Await Store.AddAsync(PassId, PassToWalletItemConverter.PassToWalletItem(PassStructure))
				Else
					PassToWalletItemConverter.UpdateWalletItemFromPass(PassStructure, WalletItem)
					Await Store.UpdateAsync(WalletItem)
				End If
			End Function
		)

		Await Store.ShowAsync()
	End Function

	Private Sub ShowErrorDialogue(Oops As Exception)
		Dim Dialog As New Windows.UI.Popups.MessageDialog(
			String.Format("The import failed due to an exception. This could be because the Passbook file was invalid or corrupt, a coding error, or a temporary system error. The exception details are listed below.{0}{0}Message: {1}{0}HRESULT: {2}{0}Help Link: {3}", vbCrLf, Oops.Message, Oops.HResult, Oops.HelpLink),
			"Import failed"
		)
		Dialog.ShowAsync().GetResults()
	End Sub
End Class
