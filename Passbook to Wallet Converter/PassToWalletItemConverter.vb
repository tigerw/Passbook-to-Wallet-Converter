Imports System.IO.Compression
Imports Windows.ApplicationModel.Wallet
Imports Windows.Data.Json

Public Module PassToWalletItemConverter
	Function GetPassId(PassStructure As JsonObject) As String
		Return PassStructure.GetNamedString("serialNumber")
	End Function

	Function ConvertPassWithResourcesToWalletItem(Pass As JsonObject, PassBundle As ZipArchive) As WalletItem
		Dim WalletItem As New WalletItem(GetPassKind(Pass), Pass.GetNamedString("description"))
		UpdateWalletItemFromPass(Pass, PassBundle, WalletItem)
		Return WalletItem
	End Function

	Sub UpdateWalletItemFromPassWithResources(Pass As JsonObject, PassBundle As ZipArchive, WalletItem As WalletItem)
		UpdateWalletItemFromPass(Pass, PassBundle, WalletItem)
	End Sub

	Private Function GetPassKind(PassStructure As JsonObject) As WalletItemKind
		If PassStructure.ContainsKey("boardingPass") Then
			Return WalletItemKind.BoardingPass
		End If

		If PassStructure.ContainsKey("coupon") Then
			Return WalletItemKind.Deal
		End If

		If PassStructure.ContainsKey("eventTicket") Then
			Return WalletItemKind.Ticket
		End If

		If PassStructure.ContainsKey("storeCard") Then
			Return WalletItemKind.MembershipCard
		End If

		Return WalletItemKind.General
	End Function

	''' <summary>Retrieves a named string from the <paramref name="PassStructure"/>.</summary>
	''' <returns>A <see cref="Windows.UI.Color"/> colour.</returns>
	Private Function GetPassColour(PassStructure As JsonObject, Key As String) As Windows.UI.Color
		Dim PassColour = PassStructure.GetNamedString(Key)
		Dim HexColours = PassColour.Substring(4, PassColour.Length - 5).Split(","c).Select(AddressOf Byte.Parse)

		Return Windows.UI.Color.FromArgb(255, HexColours.ElementAt(0), HexColours.ElementAt(1), HexColours.ElementAt(2))
	End Function

	Private Sub TryAddPassBarcodeToWalletItem(WalletItem As WalletItem, PassStructure As JsonObject)
		WalletItem.Barcode = Nothing

		If PassStructure.ContainsKey("barcodes") Then
			For Each Value In PassStructure.GetNamedArray("barcodes")
				Dim Barcode = Value.GetObject()
				Dim Symbology = GetPassBarcodeSymbology(Barcode)

				If Symbology Is Nothing Then
					Continue For
				End If

				WalletItem.Barcode = New WalletBarcode(Symbology.Value, Barcode.GetNamedString("message"))
				Return
			Next
		End If

		If PassStructure.ContainsKey("barcode") Then
			Dim Barcode = PassStructure.GetNamedObject("barcode")
			Dim Symbology = GetPassBarcodeSymbology(Barcode)

			If Symbology Is Nothing Then
				Return
			End If

			WalletItem.Barcode = New WalletBarcode(Symbology.Value, Barcode.GetNamedString("message"))
		End If

	End Sub

	Private Function GetPassBarcodeSymbology(Barcode As JsonObject) As WalletBarcodeSymbology?
		Select Case Barcode.GetNamedString("format")
			Case "PKBarcodeFormatQR"
				Return WalletBarcodeSymbology.Qr
			Case "PKBarcodeFormatPDF417"
				Return WalletBarcodeSymbology.Pdf417
			Case "PKBarcodeFormatAztec"
				Return WalletBarcodeSymbology.Aztec
			Case "PKBarcodeFormatCode128"
				Return WalletBarcodeSymbology.Code128
			Case Else
				Return Nothing
		End Select
	End Function

	Private Sub TryAddPassContentsToWalletItem(WalletItem As WalletItem, PassStructure As JsonObject)
		WalletItem.DisplayProperties.Clear()

		If PassStructure.ContainsKey("boardingPass") Then
			TryAddPassDisplayPropertiesToWalletItem(WalletItem, PassStructure.GetNamedObject("boardingPass"))
			Return
		End If

		If PassStructure.ContainsKey("coupon") Then
			TryAddPassDisplayPropertiesToWalletItem(WalletItem, PassStructure.GetNamedObject("coupon"))
			Return
		End If

		If PassStructure.ContainsKey("eventTicket") Then
			TryAddPassDisplayPropertiesToWalletItem(WalletItem, PassStructure.GetNamedObject("eventTicket"))
			Return
		End If

		If PassStructure.ContainsKey("storeCard") Then
			TryAddPassDisplayPropertiesToWalletItem(WalletItem, PassStructure.GetNamedObject("storeCard"))
			Return
		End If

		If PassStructure.ContainsKey("generic") Then
			TryAddPassDisplayPropertiesToWalletItem(WalletItem, PassStructure.GetNamedObject("generic"))
			Return
		End If
	End Sub

	Private Sub TryAddImageResourcesToWalletItem(WalletItem As WalletItem, PassBundle As ZipArchive)
		Dim SmallIcon = PassBundle.GetEntry("logo.png")
		Dim MediumIcon = If(PassBundle.GetEntry("logo@2x.png"), SmallIcon)
		Dim LargeIcon = If(PassBundle.GetEntry("logo@3x.png"), MediumIcon)
		Dim Background = If(PassBundle.GetEntry("background@2x.png"), PassBundle.GetEntry("background.png"))
		Dim Footer = If(PassBundle.GetEntry("footer@2x.png"), PassBundle.GetEntry("footer.png"))

		WalletItem.Logo336x336 = PassConverterCommon.ReadZipArchiveEntryIntoRandomAccessStream(LargeIcon)
		WalletItem.Logo159x159 = PassConverterCommon.ReadZipArchiveEntryIntoRandomAccessStream(MediumIcon)
		WalletItem.Logo99x99 = PassConverterCommon.ReadZipArchiveEntryIntoRandomAccessStream(SmallIcon)
		WalletItem.LogoImage = PassConverterCommon.ReadZipArchiveEntryIntoRandomAccessStream(LargeIcon)
		WalletItem.BodyBackgroundImage = PassConverterCommon.ReadZipArchiveEntryIntoRandomAccessStream(Background)
		WalletItem.PromotionalImage = PassConverterCommon.ReadZipArchiveEntryIntoRandomAccessStream(Footer)
	End Sub

	Private Sub TryAddPassDisplayPropertiesToWalletItem(WalletItem As WalletItem, PassStructure As JsonObject)
		Dim FirstHeaderFieldSummaryViewPosition As WalletSummaryViewPosition

		If WalletItem.LogoText <> String.Empty Then
			WalletItem.DisplayProperties.Add(
				WalletItem.LogoText.GetHashCode().ToString(),
				New WalletItemCustomProperty("Important Information", WalletItem.LogoText) With
				{
					.SummaryViewPosition = WalletSummaryViewPosition.Field1
				}
			)

			FirstHeaderFieldSummaryViewPosition = WalletSummaryViewPosition.Field2
		Else
			FirstHeaderFieldSummaryViewPosition = WalletSummaryViewPosition.Field1
		End If

		TryAddPassDisplayPropertiesToWalletItem(WalletItem, {WalletDetailViewPosition.FooterField1, WalletDetailViewPosition.FooterField2, WalletDetailViewPosition.FooterField3, WalletDetailViewPosition.FooterField4}, {}, PassStructure, "auxiliaryFields")
		TryAddPassDisplayPropertiesToWalletItem(WalletItem, {}, {}, PassStructure, "backFields")
		TryAddPassDisplayPropertiesToWalletItem(WalletItem, {WalletDetailViewPosition.HeaderField1, WalletDetailViewPosition.HeaderField2}, {FirstHeaderFieldSummaryViewPosition, WalletSummaryViewPosition.Field2}, PassStructure, "headerFields")
		TryAddPassDisplayPropertiesToWalletItem(WalletItem, {WalletDetailViewPosition.PrimaryField1, WalletDetailViewPosition.PrimaryField2}, {}, PassStructure, "primaryFields")
		TryAddPassDisplayPropertiesToWalletItem(WalletItem, {WalletDetailViewPosition.SecondaryField1, WalletDetailViewPosition.SecondaryField2, WalletDetailViewPosition.SecondaryField3, WalletDetailViewPosition.SecondaryField4, WalletDetailViewPosition.SecondaryField5}, {}, PassStructure, "secondaryFields")
	End Sub

	Private Sub TryAddPassDisplayPropertiesToWalletItem(WalletItem As WalletItem, DetailViewPositions As WalletDetailViewPosition(), SummaryViewPositions As WalletSummaryViewPosition(), PassStructure As JsonObject, Key As String)
		If Not PassStructure.ContainsKey(Key) Then
			Return
		End If

		Dim Fields = PassStructure.GetNamedArray(Key)
		Dim MaxDetailViewPosition = DetailViewPositions.GetUpperBound(0)
		Dim MaxSummaryViewPosition = SummaryViewPositions.GetUpperBound(0)

		For Index = 0 To Fields.Count - 1
			Dim DetailViewPosition = If(Index > MaxDetailViewPosition, WalletDetailViewPosition.Hidden, DetailViewPositions(Index))
			Dim SummaryViewPosition = If(Index > MaxSummaryViewPosition, WalletSummaryViewPosition.Hidden, SummaryViewPositions(Index))
			Dim Field = Fields(Index).GetObject()
			Dim FieldValue = Function()
								 Dim Value = Field.GetNamedValue("value")
								 Select Case Value.ValueType
									 Case JsonValueType.Number
										 Return CStr(Value.GetNumber())
									 Case JsonValueType.String
										 Return ProcessEmbeddedHTML(Value.GetString())
									 Case Else
										 Return String.Empty
								 End Select
							 End Function()

			WalletItem.DisplayProperties.Add(
				Field.GetNamedString("key"),
				New WalletItemCustomProperty(Field.GetNamedString("label", ""), FieldValue) With
				{
					.AutoDetectLinks = Not Field.ContainsKey("dataDetectorTypes") OrElse (Field.GetNamedArray("dataDetectorTypes").Count <> 0),
					.DetailViewPosition = DetailViewPosition,
					.SummaryViewPosition = SummaryViewPosition
				}
			)
		Next
	End Sub

	Private Function ProcessEmbeddedHTML(Text As String) As String
		Try
			Dim XML = XElement.Parse("<root>" & Text & "</root>")
			Dim Processed = String.Empty

			Dim CurrentNode = XML.FirstNode
			While CurrentNode IsNot Nothing
				Select Case CurrentNode.NodeType
					Case System.Xml.XmlNodeType.Element
						Dim CurrentElement = CType(CurrentNode, XElement)
						If CurrentElement.Name = "a" Then
							Dim HrefAttribute = CurrentElement.Attribute("href")
							If HrefAttribute IsNot Nothing Then
								Processed += HrefAttribute.Value & " · "
							End If
						End If
						Processed += CurrentElement.Value
					Case System.Xml.XmlNodeType.Text
						Processed += CType(CurrentNode, XText).Value
				End Select

				CurrentNode = CurrentNode.NextNode
			End While

			Return Processed
		Catch
			Return Text
		End Try
	End Function

	Private Sub TryAddPassRelevantLocationsToWalletItem(WalletItem As WalletItem, PassStructure As JsonObject)
		WalletItem.RelevantLocations.Clear()

		If Not PassStructure.ContainsKey("locations") Then
			Return
		End If

		For Each LocationEntry In PassStructure.GetNamedArray("locations")
			Dim Location = LocationEntry.GetObject()
			Dim BasicGeoposition As New Windows.Devices.Geolocation.BasicGeoposition() With
			{
				.Latitude = Location.GetNamedNumber("latitude"),
				.Longitude = Location.GetNamedNumber("longitude")
			}

			If Location.ContainsKey("altitude") Then
				BasicGeoposition.Altitude = Location.GetNamedNumber("altitude")
			End If

			Dim RelevantLocation As New WalletRelevantLocation() With
			{
				.Position = BasicGeoposition
			}

			If Location.ContainsKey("relevantText") Then
				RelevantLocation.DisplayMessage = Location.GetNamedString("relevantText")
			End If

			WalletItem.RelevantLocations.Add(Location.GetHashCode().ToString(), RelevantLocation)
		Next
	End Sub

	Private Sub UpdateWalletItemFromPass(Pass As JsonObject, PassBundle As ZipArchive, WalletItem As WalletItem)
		WalletItem.DisplayMessage = "Provided by the Passbook Converter application."
		WalletItem.IssuerDisplayName = Pass.GetNamedString("organizationName")

		Dim DefaultColour = New Windows.UI.Color()
		Dim BackgroundColour = If(Pass.ContainsKey("backgroundColor"), GetPassColour(Pass, "backgroundColor"), DefaultColour)
		Dim ForegroundColour = If(Pass.ContainsKey("foregroundColor"), GetPassColour(Pass, "foregroundColor"), DefaultColour)

		WalletItem.HeaderColor = BackgroundColour
		WalletItem.BodyColor = BackgroundColour

		WalletItem.BodyFontColor = ForegroundColour
		WalletItem.HeaderFontColor = ForegroundColour
		WalletItem.LogoText = If(Pass.ContainsKey("logoText"), Pass.GetNamedString("logoText"), String.Empty)
		WalletItem.ExpirationDate = If(Pass.ContainsKey("expirationDate"), DateTimeOffset.Parse(Pass.GetNamedString("expirationDate")), CType(Nothing, DateTimeOffset?))

		TryAddPassBarcodeToWalletItem(WalletItem, Pass)
		TryAddPassContentsToWalletItem(WalletItem, Pass)
		TryAddImageResourcesToWalletItem(WalletItem, PassBundle)
		TryAddPassRelevantLocationsToWalletItem(WalletItem, Pass)

		WalletItem.RelevantDate = If(Pass.ContainsKey("relevantDate"), DateTimeOffset.Parse(Pass.GetNamedString("relevantDate")), CType(Nothing, DateTimeOffset?))
	End Sub
End Module