Imports Windows.Data.Json

Public Module PassToWalletItemXmlConverter
	Public Function PassToWalletItem(PassStructure As JsonObject) As XElement
		Dim WalletItem =
			<WalletItem>
				<Version>1</Version>
				<Id><%= PassStructure.GetNamedString("serialNumber") %></Id>
				<Kind><%= GetPassKind(PassStructure) %></Kind>
				<DisplayName><%= PassStructure.GetNamedString("description") %></DisplayName>
				<IssuerDisplayName><%= PassStructure.GetNamedString("organizationName") %></IssuerDisplayName>
				<HeaderColor><%= GetPassColour(PassStructure, "foregroundColor") %></HeaderColor>
				<BodyColor><%= GetPassColour(PassStructure, "backgroundColor") %></BodyColor>
				<DisplayMessage><%= PassStructure.GetNamedString("description") %></DisplayMessage>
			</WalletItem>

		If PassStructure.ContainsKey("labelColor") Then
			Dim Colour = GetPassColour(PassStructure, "labelColor")
			WalletItem.Add(<HeaderFontColor><%= Colour %></HeaderFontColor>)
			WalletItem.Add(<BodyFontColor><%= Colour %></BodyFontColor>)
		End If

		If PassStructure.ContainsKey("logoText") Then
			WalletItem.Add(<LogoText><%= PassStructure.GetNamedString("logoText") %></LogoText>)
		End If

		If PassStructure.ContainsKey("expirationDate") Then
			WalletItem.Add(<ExpirationDate><%= PassStructure.GetNamedString("expirationDate") %></ExpirationDate>)
		End If

		TryAddPassBarcodeToWalletItem(WalletItem, PassStructure)
		TryAddPassContentsToWalletItem(WalletItem, PassStructure)
		TryAddPassRelevantLocationsToWalletItem(WalletItem, PassStructure)

		If PassStructure.ContainsKey("relevantDate") Then
			WalletItem.Add(<RelevantDate><Date><%= PassStructure.GetNamedString("relevantDate") %></Date></RelevantDate>)
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
			Return
		End If

		Dim Barkode = PassStructure.GetNamedObject("barcode")
		WalletItem.Add(
			<Barcode>
				<Symbology><%= GetPassBarcodeSymbology(Barkode) %></Symbology>
				<Value><%= Barkode.GetNamedString("message") %></Value>
			</Barcode>
		)
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

		REM TryAddPassDisplayProperties(DisplayPropertiesElement, <Center/>, PassStructure, "auxiliaryFields")
		TryAddPassDisplayPropertiesToElement(DisplayPropertiesElement, <Footer/>, PassStructure, "backFields")
		TryAddPassDisplayPropertiesToElement(DisplayPropertiesElement, <Header/>, PassStructure, "headerFields")
		TryAddPassDisplayPropertiesToElement(DisplayPropertiesElement, <Primary/>, PassStructure, "primaryFields")
		TryAddPassDisplayPropertiesToElement(DisplayPropertiesElement, <Secondary/>, PassStructure, "secondaryFields")

		WalletItem.Add(DisplayPropertiesElement)
	End Sub

	Private Sub TryAddPassDisplayPropertiesToElement(DisplayPropertiesElement As XElement, LocationElement As XElement, PassStructure As JsonObject, Key As String)
		If Not PassStructure.ContainsKey(Key) Then
			Return
		End If

		Dim Added = False

		For Each FieldEntry In PassStructure.GetNamedArray(Key)
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
				Dim AbortProperty = False
				Dim Format = GetPassDisplayPropertyDateTimeFormat(Field, AbortProperty)
				If AbortProperty Then
					Continue For
				End If
				If Format IsNot Nothing Then
					PropertyElement.Add(<DateTimeFormat><%= Format %></DateTimeFormat>)
				End If
			ElseIf Field.ContainsKey("currencyCode") Then
				PropertyElement.Add(<CurrencyCode><%= Field.GetNamedString("currencyCode") %></CurrencyCode>)
			ElseIf Field.ContainsKey("numberStyle") Then
				PropertyElement.Add(<NumberFormat><%= GetPassDisplayPropertyCurrencyCode(Field) %></NumberFormat>)
			End If

			Added = True
			LocationElement.Add(PropertyElement)
		Next

		If Added Then
			DisplayPropertiesElement.Add(LocationElement)
		End If
	End Sub

	Private Function GetPassDisplayPropertyDateTimeFormat(PassStructure As JsonObject, ByRef AbortProperty As Boolean) As String
		Dim DateStyle = PassStructure.GetNamedString("dateStyle")
		Dim TimeStyle = PassStructure.GetNamedString("timeStyle")

		Select Case DateStyle
			Case "PKDateStyleNone"
				Select Case TimeStyle
					Case "PKDateStyleNone"
						AbortProperty = True
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
				End Select
			Case "PKDateStyleLong"
			Case "PKDateStyleFull"
				Select Case TimeStyle
					Case "PKDateStyleNone"
						Return "LongDate"
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