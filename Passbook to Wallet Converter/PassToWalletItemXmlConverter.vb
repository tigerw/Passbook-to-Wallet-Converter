Imports System.IO.Compression
Imports Windows.Data.Json

Public Module PassToWalletItemXmlConverter
	Async Function ConvertPassBundleToWalletPackage(PassBundle As ZipArchive, OutputStream As Stream) As Task
		Using WalletPackage = New ZipArchive(OutputStream, ZipArchiveMode.Create, True)
			Dim WalletItem = PassToWalletItem(Await PassConverterCommon.GetPassFromPassBundle(PassBundle))
			Dim WalletItemEntry = WalletPackage.CreateEntry("WalletItem.xml")
			Using WalletItemWriter = WalletItemEntry.Open()
				WalletItem.Save(WalletItemWriter)
			End Using
		End Using
	End Function

	Private Function PassToWalletItem(Pass As JsonObject) As XElement
		Dim WalletItem =
			<WalletItem>
				<Id><%= Pass.GetNamedString("serialNumber") %></Id>
				<Kind><%= GetPassKind(Pass) %></Kind>
				<DisplayName><%= Pass.GetNamedString("description") %></DisplayName>
				<IssuerDisplayName><%= Pass.GetNamedString("organizationName") %></IssuerDisplayName>
				<DisplayMessage>Provided by the Passbook Converter application.</DisplayMessage>
			</WalletItem>

		Dim BackgroundColour = GetPassColour(Pass, "backgroundColor")
		Dim ForegroundColour = GetPassColour(Pass, "foregroundColor")

		If BackgroundColour IsNot Nothing Then
			WalletItem.Add(<HeaderColor><%= BackgroundColour %></HeaderColor>)
			WalletItem.Add(<BodyColor><%= BackgroundColour %></BodyColor>)
		End If

		If ForegroundColour IsNot Nothing Then
			WalletItem.Add(<HeaderFontColor><%= ForegroundColour %></HeaderFontColor>)
			WalletItem.Add(<BodyFontColor><%= ForegroundColour %></BodyFontColor>)
		End If

		If Pass.ContainsKey("logoText") Then
			WalletItem.Add(<LogoText><%= Pass.GetNamedString("logoText") %></LogoText>)
		End If

		If Pass.ContainsKey("expirationDate") Then
			WalletItem.Add(<ExpirationDate><%= Pass.GetNamedString("expirationDate") %></ExpirationDate>)
		End If

		TryAddPassBarcodeToWalletItem(WalletItem, Pass)
		TryAddPassContentsToWalletItem(WalletItem, Pass)
		TryAddPassRelevantLocationsToWalletItem(WalletItem, Pass)

		If Pass.ContainsKey("relevantDate") Then
			WalletItem.Add(<RelevantDate><Date><%= Pass.GetNamedString("relevantDate") %></Date></RelevantDate>)
		End If

		Return WalletItem
	End Function

	Private Function GetPassKind(PassStructure As JsonObject) As String
		If PassStructure.ContainsKey("boardingPass") Then
			Return "BoardingPass"
		End If

		If PassStructure.ContainsKey("coupon") Then
			Return "Deal"
		End If

		If PassStructure.ContainsKey("eventTicket") Then
			Return "Ticket"
		End If

		If PassStructure.ContainsKey("storeCard") Then
			Return "MembershipCard"
		End If

		Return "General"
	End Function

	''' <summary>Retrieves a named string from the <paramref name="PassStructure"/>, defaulting to white if the key doesn't exist.</summary>
	''' <returns>A string in the form #rrggbb.</returns>
	Private Function GetPassColour(PassStructure As JsonObject, Key As String) As String
		If Not PassStructure.ContainsKey(Key) Then
			Return Nothing
		End If

		Dim PassColour = PassStructure.GetNamedString(Key, "rgb(255, 255,255)")
		Dim HexColours = PassColour.Substring(4, PassColour.Length - 5).Split(","c).Select(
			Function(ColourComponent As String)
				Return Integer.Parse(ColourComponent).ToString("X2")
			End Function
		)

		Return "#"c + String.Join("", HexColours)
	End Function

	Private Sub TryAddPassBarcodeToWalletItem(WalletItem As XElement, PassStructure As JsonObject)
		If PassStructure.ContainsKey("barcodes") Then
			For Each Value In PassStructure.GetNamedArray("barcodes")
				Dim Barcode = Value.GetObject()
				Dim Symbology = GetPassBarcodeSymbology(Barcode)

				If Symbology Is Nothing Then
					Continue For
				End If

				WalletItem.Add(
					<Barcode>
						<Symbology><%= Symbology %></Symbology>
						<Value><%= Barcode.GetNamedString("message") %></Value>
					</Barcode>
				)
				Return
			Next
		End If

		If Not PassStructure.ContainsKey("barcode") Then
			Dim Barcode = PassStructure.GetNamedObject("barcode")
			Dim Symbology = GetPassBarcodeSymbology(Barcode)

			WalletItem.Add(
				<Barcode>
					<Symbology><%= Symbology %></Symbology>
					<Value><%= Barcode.GetNamedString("message") %></Value>
				</Barcode>
			)
		End If
	End Sub

	Private Function GetPassBarcodeSymbology(Barcode As JsonObject) As String
		Select Case Barcode.GetNamedString("format")
			Case "PKBarcodeFormatQR"
				Return "QR"
			Case "PKBarcodeFormatPDF417"
				Return "PDF417"
			Case "PKBarcodeFormatAztec"
				Return "AZTEC"
			Case "PKBarcodeFormatCode128"
				Return "CODE128"
			Case Else
				Return Nothing
		End Select
	End Function

	Private Sub TryAddPassContentsToWalletItem(WalletItem As XElement, PassStructure As JsonObject)
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

	Private Sub TryAddPassDisplayPropertiesToWalletItem(WalletItem As XElement, PassStructure As JsonObject)
		Dim DisplayPropertiesElement = <DisplayProperties/>

		TryAddPassDisplayPropertiesToElement(DisplayPropertiesElement, Nothing, PassStructure, "backFields")
		TryAddPassDisplayPropertiesToElement(DisplayPropertiesElement, <Footer/>, PassStructure, "auxiliaryFields")
		TryAddPassDisplayPropertiesToElement(DisplayPropertiesElement, <Header/>, PassStructure, "headerFields")
		TryAddPassDisplayPropertiesToElement(DisplayPropertiesElement, <Primary/>, PassStructure, "primaryFields")
		TryAddPassDisplayPropertiesToElement(DisplayPropertiesElement, <Secondary/>, PassStructure, "secondaryFields")

		WalletItem.Add(DisplayPropertiesElement)
	End Sub

	Private Sub TryAddPassDisplayPropertiesToElement(DisplayPropertiesElement As XElement, LocationElement As XElement, PassStructure As JsonObject, Key As String)
		If Not PassStructure.ContainsKey(Key) Then
			Return
		End If

		Dim DisplayProperties = PassStructure.GetNamedArray(Key)

		If LocationElement Is Nothing Then
			TryAddPassDisplayPropertiesToElement(DisplayPropertiesElement, DisplayProperties)
		Else
			If TryAddPassDisplayPropertiesToElement(LocationElement, DisplayProperties) Then
				DisplayPropertiesElement.Add(LocationElement)
			End If
		End If
	End Sub

	Private Function TryAddPassDisplayPropertiesToElement(Element As XElement, DisplayProperties As JsonArray) As Boolean
		Dim Added = False

		For Each FieldEntry In DisplayProperties
			Dim Field = FieldEntry.GetObject()
			Dim PropertyElement =
				<Property>
					<Key><%= Field.GetNamedString("key") %></Key>
					<Value><%= Field.GetNamedString("value") %></Value>
				</Property>

			If Not Field.ContainsKey("dataDetectorTypes") OrElse (Field.GetNamedArray("dataDetectorTypes").Count <> 0) Then
				PropertyElement.Add(<AutoDetectLinks/>)
			End If

			If Field.ContainsKey("label") Then
				PropertyElement.Add(<Name><%= Field.GetNamedString("label") %></Name>)
			End If

			If Field.ContainsKey("dateStyle") Then
				Dim Format = GetPassDisplayPropertyDateTimeFormat(Field)
				If Format IsNot Nothing Then
					PropertyElement.Add(<DateTimeFormat><%= Format %></DateTimeFormat>)
				End If
			ElseIf Field.ContainsKey("currencyCode") Then
				PropertyElement.Add(<CurrencyCode><%= Field.GetNamedString("currencyCode") %></CurrencyCode>)
			ElseIf Field.ContainsKey("numberStyle") Then
				PropertyElement.Add(<NumberFormat><%= GetPassDisplayPropertyCurrencyCode(Field) %></NumberFormat>)
			End If

			Added = True
			Element.Add(PropertyElement)
		Next

		Return Added
	End Function

	Private Function GetPassDisplayPropertyDateTimeFormat(PassStructure As JsonObject) As String
		Dim DateStyle = PassStructure.GetNamedString("dateStyle")
		Dim TimeStyle = PassStructure.GetNamedString("timeStyle")

		Select Case DateStyle
			Case "PKDateStyleNone"

				Select Case TimeStyle
					Case "PKDateStyleNone"
						Return Nothing
					Case "PKDateStyleShort"
					Case "PKDateStyleMedium"
						Return "ShortTime"
					Case "PKDateStyleLong"
					Case "PKDateStyleFull"
						Return "LongTime"
				End Select

			Case "PKDateStyleShort"
			Case "PKDateStyleMedium"

				Select Case TimeStyle
					Case "PKDateStyleNone"
						Return "ShortDate"
					Case "PKDateStyleShort"
					Case "PKDateStyleMedium"
						Return "ShortDateTime"
					Case "PKDateStyleLong"
					Case "PKDateStyleFull"
						Return "FullDateTime"
				End Select

			Case "PKDateStyleLong"
			Case "PKDateStyleFull"

				Select Case TimeStyle
					Case "PKDateStyleNone"
						Return "LongDate"
					Case "PKDateStyleShort"
					Case "PKDateStyleMedium"
						Return "FullDateTime"
					Case "PKDateStyleLong"
					Case "PKDateStyleFull"
						Return "FullDateTime"
				End Select

		End Select

		Return Nothing
	End Function

	Private Function GetPassDisplayPropertyCurrencyCode(PassStructure As JsonObject) As String
		Select Case PassStructure.GetNamedString("numberStyle")
			Case "PKNumberStyleDecimal"
				Return "Decimal"
			Case "PKNumberStylePercent"
				Return "Percent"
			Case "PKNumberStylePercent"
				Return "Exponential"
			Case "PKNumberStyleSpellOut"
				Return "Currency"
			Case Else
				Return Nothing
		End Select
	End Function

	Private Sub TryAddPassRelevantLocationsToWalletItem(WalletItem As XElement, PassStructure As JsonObject)
		If Not PassStructure.ContainsKey("locations") Then
			Return
		End If

		Dim LocationsElement = <Locations></Locations>

		For Each LocationEntry In PassStructure.GetNamedArray("locations")
			Dim Location = LocationEntry.GetObject()
			Dim LocationElement =
				<Location>
					<Latitude><%= Location.GetNamedNumber("latitude") %></Latitude>
					<Longitude><%= Location.GetNamedNumber("longitude") %></Longitude>
				</Location>

			If Location.ContainsKey("altitude") Then
				LocationElement.Add(<Altitude><%= Location.GetNamedNumber("altitude") %></Altitude>)
			End If

			If Location.ContainsKey("relevantText") Then
				LocationElement.Add(<DisplayMessage><%= Location.GetNamedString("relevantText") %></DisplayMessage>)
			End If

			LocationsElement.Add(LocationElement)
		Next

		WalletItem.Add(LocationsElement)
	End Sub
End Module