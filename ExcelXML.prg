*-- This PRG file was extracted to a PRG file from the original VCX file by Matt Slay 2017-09-03.
*-- This allows better updates of the source code by the VFP community on GitHib.
*-- GitHub: https://github.com/VFPX/ExcelXML
*---------------------------------------------------------------------------------------
*-- Change Log:
*---------------------------------------------------------------------------------------
*- 2017-09-10:	Ver 1.10
*--				1. Added new method:  ConvertXmlToXlsx(tcFilename, tnFileFormat, tlOpenAfterExporting)
*--				2. Fixed bug in Bottom Border logic if cursor/grid only has 1 row of data.
*--				3. Added new properties for ColumnHeaderBackgroundColor and ColumnHeaderForeColor 
*--				4. Added new property GridClass to use when creating a temporary form to create grid to host the cursor/alias during the export.
*- 
*- 2017-09-03:  Ver 1.09
*--				Added Try/Catch to handle Dynamic properties that do evaluate properly.
*-				[By Matt Slay]
*---------------------------------------------------------------------------------------

Define Class ExcelXml As Custom

	* Array with information about the structure of the table in a ;
	* specified work area, specified by a table alias, or in the currently ;
	* selected work area in an array and returns the number of fields in ;
	* the table.
	* Name of the table/cursor defined in the Grid or name of current ;
	* table/cursor opened.
	Alias           = ''
	* Returns the number of columns included in the Excel file.
	ColumnCount     = 0
	crlf            = ''
	* Specifies the date format.
	DateFormat      = ''
	* Inform the name of Excel file. If you don't inform the name with the ;
	* extension, the XML extension will be included. The default file name ;
	* is "Book1"
	File            = ''
	* Inform the grid control object to convert a grid control in an Excel ;
	* XML file.
	GridObject      = ''
	* .T. Includes the option Filter in all columns in the generated file.
	HasFilter       = .F.
	Height          = 16
	* .T. locks the header in the generated file. This option in Excel is ;
	* called by Freeze Top Row.
	LockHeader      = .F.
	* .T. to open the file after saving it.
	OpenAfterSaving = .F.
	* Returns the number of rows included in the Excel file.
	RowCount        = 0
	* Defines if the Excel file will have all the grid graphical attributes ;
	* transported.
	SetStyles       = .T.
	* Excel sheet name. The default name is "Sheet1"
	SheetName       = 'Sheet1'
	stylecodenumber = 0
	* Object that contain the information about this class.
	Version         = ''
	Width           = 70
	* XML encoding type used to set the code that defines special ;
	* characters. Default code is "iso-8859-1".
	xmlEncoding     = 'iso-8859-1'
	cErrorMessage = ""
	* The grid class name to use when creating a temporary form to create grid to host the cursor
	* during the export.
	GridClass = "grid"
	* Colmn Header Background color. Can override grid header backcolor. Set to a string with Hex value, like "#CCCCCC" for light gray.
	ColumnHeaderBackgroundColor = .null.
	* Colmn Header ForegColor. Can override grid header forecolor. Set to a string with Hex value, like "#000000" for black.
	ColumnHeaderForeColor = .null.
	

	*|================================================================================ 
	*| ExcelXml::
	Procedure About
	
		Messagebox("ExcelXml " + This.Version.Number + " " + This.Version.Datetime + This.crlf + ;
					"Converts a Grid control into a Microsoft Excel XML file" + This.crlf + ;
					"" + This.crlf + ;
					"Created by " + This.Version.Author + This.crlf + ;
					This.Version.CountryAndCity + This.crlf + ;
					This.Version.url + This.crlf + ;
					This.Version.Email, 64, "About ExcelXml")
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure AddNewStyle
		Lparameters plcType, plnRow, plnCol, ;
			plcAlignH, plcAlignV, plcFontName, plcFontFamily, ;
			plcFontSize, plcForeColor, plcFontBold, plcFontItalic, ;
			plcFontUnderline, plcFontStrikeThru, plcBackColor, plcPattern, ;
			plcFormat

		Local lcStyleCode, lcXmlStyle
		lcXmlStyle = ""

		*- Defini��o de bordas entre as linhas/colunas (c�lulas)
		lcXmlBorderStyle = ""
		lcTop = "0"
		lcBottom = "0"

		If This.GridObject.GridLines >= 1 And This.SetStyles
			lcGridLineWidth = Iif(plcType = "c", Alltrim(Str(Iif(This.GridObject.GridLineWidth >= 4, 3, This.GridObject.GridLineWidth))), "1")
			lcGridLineColor = Iif(plcType = "c", This.ColorToStrHexa(This.GridObject.GridLineColor), This.ColorToStrHexa(Rgb(100, 100, 100)))
			lcXmlBorderStyle = [   <Borders>] + This.crlf

		*- Linhas na horizontal   
			If Inlist(This.GridObject.GridLines, 1, 3)
				lcXmlBorderStyle = lcXmlBorderStyle + [    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="] + lcGridLineWidth + [" ss:Color="] + lcGridLineColor + ["/>] + This.crlf
				lcXmlBorderStyle = lcXmlBorderStyle + [    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="] + lcGridLineWidth + [" ss:Color="] + lcGridLineColor + ["/>] + This.crlf
			EndIf
			
		*- Linhas na vertical
			If Inlist(This.GridObject.GridLines, 2, 3)
				If This.GridObject.GridLines = 2
					If plnRow = 1					&&- Se for a primeira linha
						lcTop = "1"
						lcXmlBorderStyle = lcXmlBorderStyle + [    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="] + lcGridLineWidth + [" ss:Color="] + lcGridLineColor + ["/>] + This.crlf
					Endif
					If plnRow = (This.RowCount - 1)	&&- Se for a ultima linha
						lcBottom = "1"
						lcXmlBorderStyle = lcXmlBorderStyle + [    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="] + lcGridLineWidth + [" ss:Color="] + lcGridLineColor + ["/>] + This.crlf
					Endif
				Endif

				lcXmlBorderStyle = lcXmlBorderStyle + [    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="] + lcGridLineWidth + [" ss:Color="] + lcGridLineColor + ["/>] + This.crlf
				lcXmlBorderStyle = lcXmlBorderStyle + [    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="] + lcGridLineWidth + [" ss:Color="] + lcGridLineColor + ["/>] + This.crlf
			Endif

			lcXmlBorderStyle = lcXmlBorderStyle + [   </Borders>] + This.crlf
		Else
			lcXmlBorderStyle = [   <Borders></Borders>] + This.crlf
		Endif


		*- Adiciono no cursor caso n�o ache um registro com os mesmos dados
		If Not Seek(plcAlignH + plcAlignV + plcFontName + plcFontFamily + ;
						plcFontSize + plcForeColor + plcFontBold + plcFontItalic + ;
						plcFontUnderline + plcFontStrikeThru + plcBackColor + plcPattern + ;
						plcFormat + lcTop + lcBottom, ;
						"xxxStylesProperties", "idxStyle")

			This.stylecodenumber = This.stylecodenumber + 1
			lcStyleCode = Alltrim(Lower(plcType)) + Transform(This.stylecodenumber, "@L 99999")

			*- xml de estilo da celula
			lcXmlStyle = [  <Style ss:ID="] + lcStyleCode + [">] + This.crlf + ;
				[   <Alignment ss:Horizontal="] + Alltrim(plcAlignH) + [" ss:Vertical="] + Alltrim(plcAlignV) + ["/>] + This.crlf + ;
				[   <Font ss:FontName="] + Alltrim(plcFontName) + [" x:Family="] + Alltrim(plcFontFamily) + [" ss:Size="] + Alltrim(Str(Val(plcFontSize))) + [" ss:Color="] + Alltrim(plcForeColor) + ["] + This.crlf + ;
				[    ss:Bold="] + plcFontBold + [" ss:Italic="] + plcFontItalic + ["] + Iif(!Empty(plcFontUnderline), [ ss:Underline="] + Alltrim(plcFontUnderline) + ["], "") + [ ss:StrikeThrough="] + Alltrim(plcFontStrikeThru) + ["/>] + This.crlf + ;
				Iif(This.SetStyles, [   <Interior ss:Color="] + Alltrim(plcBackColor) + [" ss:Pattern="] + Alltrim(plcPattern) + ["/>], [   <Interior/>] ) + This.crlf + ;
				Iif(!Empty(plcFormat), [   <NumberFormat ss:Format="] + Alltrim(plcFormat) + ["/>] + This.crlf, "") + ;
				lcXmlBorderStyle + ;
				[  </Style>]

			Insert Into xxxStylesProperties ;
				Values ( 	lcStyleCode, ;
						plcAlignH, ;
						plcAlignV, ;
						plcFontName, ;
						plcFontFamily, ;
						plcFontSize, ;
						plcForeColor, ;
						plcFontBold, ;
						plcFontItalic, ;
						plcFontUnderline, ;
						plcFontStrikeThru, ;
						plcBackColor, ;
						plcPattern, ;
						plcFormat, ;
						lcTop, ;
						lcBottom, ;
						lcXmlStyle )
		Endif

		Insert Into xxxStylesRowCol ;
			Values ( Transform(plnRow, "@L 999999"), ;
					Transform(plnCol, "@L 999"), ;
					xxxStylesProperties.ssCode )

		Return lcXmlStyle
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure BuildColumnsStyles
	
		Local lcAlignH, lcAlignV, lcFontName, lcFontFamily, ;
			lcFontSize, lcForeColor, lcFontBold, lcFontItalic, ;
			lcFontUnderline, lcFontStrikeThru, lcBackColor, lcPattern, ;
			lcFormat, lcXmlBorderStyle, lcXmlStyles, lnRow, lnCol, lnRowFound

		This.stylecodenumber = 0
		lnRow = 0
		lnCol = 0
		lcXmlStyles = ""


		*- Verifico os estilos de todas as linhas/colunas do grid
		Select (This.Alias)
		Go Top
		Scan
			lnRow = lnRow + 1

			If Not This.SetStyles And lnRow >= 2		&&- N�o aplica os estilos ao grid.
				Exit
			Endif

			For lnCol = 1 To This.GridObject.ColumnCount
				loColumn = This.GetColumn(lnCol)
				If Not loColumn.Visible 	&&-Considero somente as colunas visiveis
					Loop
				Endif

				*- Formato dos dados da linha/coluna (c�lula)
				lcDataColumn = Evaluate(loColumn.ControlSource)
				loCurrentControl = This.GetCurrentControlObject(loColumn)
				lcFormat = ""

				If Not Isnull(loCurrentControl)
					Do Case
					Case Inlist(Vartype(lcDataColumn), "N", "Y")
						If Lower(loCurrentControl.BaseClass) $ "textbox//spinner"
							If Not Empty(loColumn.InputMask)
								lcInputMask = loColumn.InputMask
								If Occurs(".", lcInputMask) > 0
									lcFormat = "#,##0." + Replicate("0", Len(Subs(lcInputMask, Rat(".", lcInputMask) + 1)))
								Else
									lcFormat = "#,##0"
								Endif
							Else
								lnRowFound = Ascan(This._Fields, Iif("." $ loColumn.ControlSource, Substr(loColumn.ControlSource, At(".", loColumn.ControlSource) + 1), loColumn.ControlSource), -1, -1, 1, 15)
								If lnRowFound > 0 And This._Fields[lnRowFound, 4] > 0
									lcFormat = "#,##0." + Replicate("0", This._Fields[lnRowFound, 4])
								Else
									lcFormat = ""
								Endif
							Endif
						Endif

						If Lower(loCurrentControl.BaseClass) $ "checkbox//optiongroup"
							lcFormat = ""
						Endif

					Case Vartype(lcDataColumn) = "D"
						lcFormat = This.DateFormat + ";@"

					Case Vartype(lcDataColumn) = "T"
						If Lower(loCurrentControl.BaseClass) = "textbox"
							lnHasSeconds = loCurrentControl.Seconds
						Else
							lnHasSeconds = 2
						Endif

						If lnHasSeconds = 0
							*- Data e hora sem segundos
							lcFormat = This.DateFormat + "\ h:mm" + Iif(Set("hours") = 12, " AM/PM", "")
						Else
							*- Data e hora com segundos
							lcFormat = This.DateFormat + "\ h:mm:ss" + Iif(Set("hours") = 12, " AM/PM", "")
						Endif

					Case Vartype(lcDataColumn) = "L"
						lcFormat = "True/False"

					Otherwise
						lcFormat = ""
					Endcase
				Endif

				lcFormat = Padr(lcFormat, Len(xxxStylesProperties.ssFormat))

				*- Requisitos fixos para o estilo
				lcFontFamily = Padr("Swiss", Len(xxxStylesProperties.ssFontFamily))
				lcPattern = Padr("Solid", Len(xxxStylesProperties.ssPattern))

				*- Alinhamento Horizontal do texto da coluna/linha
				If Not Isnull(loCurrentControl) And Lower(loCurrentControl.BaseClass) = "combobox"
					lcAlignH = This.GetColumnAlign("H", loCurrentControl.Alignment, Vartype(lcDataColumn))
					lcAlignV = This.GetColumnAlign("V", loCurrentControl.Alignment, Vartype(lcDataColumn))
				Else
					lcAlignH = This.GetColumnAlign("H", loColumn.Alignment, Vartype(lcDataColumn))
					lcAlignV = This.GetColumnAlign("V", loColumn.Alignment, Vartype(lcDataColumn))
				Endif

				*- cor de fundo da coluna/linha
				If Not Empty(loColumn.DynamicBackColor)
					Try
						lcBackColor = This.ColorToStrHexa(Evaluate(loColumn.DynamicBackColor) )
					Catch
						lcBackColor = This.ColorToStrHexa( loColumn.BackColor )
					Endtry
				Else
					lcBackColor = This.ColorToStrHexa( loColumn.BackColor )
				Endif


				*- cor da fonte da coluna/linha
				If Not Empty(loColumn.DynamicForeColor)
					Try
						lcForeColor = This.ColorToStrHexa( Evaluate(loColumn.DynamicForeColor) )
					Catch
						lcForeColor = This.ColorToStrHexa( loColumn.ForeColor )
					Endtry
				Else
					lcForeColor = This.ColorToStrHexa( loColumn.ForeColor )
				Endif


				*- fonte usada na coluna/linha
				If Not Empty(loColumn.DynamicFontName)
					Try
						lcFontName = Evaluate(loColumn.DynamicFontName)
					Catch
						lcFontName = Padr(lcFontName, Len(xxxStylesProperties.ssFontName))
					Endtry
				Else
					lcFontName = loColumn.FontName
				Endif
				lcFontName = Padr(lcFontName, Len(xxxStylesProperties.ssFontName))


				*- tamanho da fonte da coluna/linha			
				If Not Empty(loColumn.DynamicFontSize)
					Try
						lcFontSize = Transform(Evaluate(loColumn.DynamicFontSize), "@L 999")
					Catch
						lcFontSize = Transform(loColumn.FontSize, "@L 999")
					Endtry
				Else
					lcFontSize = Transform(loColumn.FontSize, "@L 999")
				Endif


				*- Fonte Italica da coluna/linha
				If Not Empty(loColumn.DynamicFontItalic)
					Try
						lcFontItalic = Iif(Evaluate(loColumn.DynamicFontItalic), "1", "0")
					Catch
						lcFontItalic = Iif(loColumn.FontItalic, "1", "0")
					Endtry
				Else
					lcFontItalic = Iif(loColumn.FontItalic, "1", "0")
				Endif


				*- Fonte Negrito da coluna/linha
				If Not Empty(loColumn.DynamicFontBold)
					Try
						lcFontBold = Iif(Evaluate(loColumn.DynamicFontBold), "1", "0")
					Catch
						lcFontBold = Iif(loColumn.FontBold, "1", "0")
					Endtry
				Else
					lcFontBold = Iif(loColumn.FontBold, "1", "0")
				Endif


				*- Fonte Underline da coluna/linha
				If Not Empty(loColumn.DynamicFontUnderline)
					Try
						lcFontUnderline = Iif(Evaluate(loColumn.DynamicFontUnderline), "Single", "")
					Catch
						lcFontUnderline = Iif(loColumn.FontUnderline, "Single", "")
					Endtry
				Else
					lcFontUnderline = Iif(loColumn.FontUnderline, "Single", "")
				Endif
				lcFontUnderline = Padr(lcFontUnderline, Len(xxxStylesProperties.ssFontUnderline))


				*- Fonte Underline da coluna/linha
				If Not Empty(loColumn.DynamicFontStrikethru)
					Try
						lcFontStrikeThru = Iif(Evaluate(loColumn.DynamicFontStrikethru), "1", "0")
					Catch
						lcFontStrikeThru = Iif(loColumn.FontStrikethru, "1", "0")
					Endtry
				Else
					lcFontStrikeThru = Iif(loColumn.FontStrikethru, "1", "0")
				Endif
				lcFontStrikeThru = Padr(lcFontStrikeThru, Len(xxxStylesProperties.ssFontStrikeThru))


				*- se o estilo j� existir "lcXmlStyle" retorna ""
				lcXmlStyle = This.AddNewStyle( "c", lnRow, lnCol, ;
							lcAlignH, lcAlignV, lcFontName, lcFontFamily, ;
							lcFontSize, lcForeColor, lcFontBold, lcFontItalic, ;
							lcFontUnderline, lcFontStrikeThru, lcBackColor, lcPattern, ;
							lcFormat )

				If Not Empty(lcXmlStyle)
					lcXmlStyles = lcXmlStyles + This.crlf + lcXmlStyle
				Endif
			Endfor
		Endscan

		Return lcXmlStyles
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure BuildColumnsWidth
	
		Local lcXmlColumnsWidth, lnCol, lnColumnWidth
		
		lcXmlColumnsWidth = This.crlf

		For lnCol = 1 To This.GridObject.ColumnCount
			loColumn = This.GetColumn(lnCol)
			If loColumn.Visible = .T.
				lnColumnWidth = Iif(loColumn.Width > 700, 700, loColumn.Width)		&&- Avoiding error in Excel
				lcXmlColumnsWidth = lcXmlColumnsWidth + [   <Column ss:AutoFitWidth="0" ss:Width="] + Alltrim(Str(lnColumnWidth)) + ["/>] + This.crlf
			Endif
		Endfor

		Return lcXmlColumnsWidth
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure BuildHeadersStyles
	
		Local loColumn, loColumnHeader, lnCol, lcXmlStyles, lcXmlStyle, ;
			lcBackColor, lcForeColor, lcFontName, lcFontSize, lcFontItalic, ;
			lcFontBold,	lcFontUnderline, lcFontStrikeThru, lcFormat, ;
			lcFontFamily, lcPattern, lcAlignH, lcAlignV, lcCollate

		lcXmlStyle = ""
		lcXmlStyles = ""
		This.stylecodenumber = 0
		lcCollate = Set("Collate")

		Set Collate To "MACHINE"

		*- Crio cursor para armazenar todos os estilos encontrados
		Create Cursor xxxStylesProperties ( ssCode c(6), ;
					ssAlignH c(6), ;
					ssAlignV c(6), ;
					ssFontName c(40), ;
					ssFontFamily c(5), ;
					ssFontSize c(3), ;
					ssFontColor c(7), ;
					ssFontBold c(1), ;
					ssFontItalic c(1), ;
					ssFontUnderline c(6), ;
					ssFontStrikeThru c(1), ;
					ssBackColor c(7), ;
					ssPattern c(5), ;
					ssFormat c(40), ;
					ssTop c(1), ;
					ssBottom c(1), ;
					ssStyle m )



		Select xxxStylesProperties
		Index On ssAlignH + ssAlignV + ssFontName +	;
			ssFontFamily + ssFontSize + ssFontColor + ;
			ssFontBold + ssFontItalic + ssFontUnderline + ssFontStrikeThru + ;
			ssBackColor + ssPattern + ssFormat	+ ssTop + ssBottom Tag idxStyle

		Index On ssCode Tag idxCode

		*- Crio cursor para gravar o estilo que sera usado pela linha/coluna (c�lula)
		Create Cursor xxxStylesRowCol ( ssRow c(6), ;
					ssCol c(3), ;
					ssCode c(6) )

		Select xxxStylesRowCol
		Index On ssRow + ssCol Tag idxRowCol

		Set Collate To lcCollate


		*- Verifico os estilos dos headers de cada coluna
		If This.GridObject.HeaderHeight > 0
			For lnCol = 1 To This.GridObject.ColumnCount
				loColumn = This.GetColumn(lnCol)
				loColumnHeader = This.GetColumnHeader(loColumn)

				If IsNull(This.ColumnHeaderBackgroundColor)
					lcBackColor = This.ColorToStrHexa( Iif(This.SetStyles, loColumnHeader.BackColor, Rgb(255, 255, 255)) )
				Else
					lcBackColor = This.ColumnHeaderBackgroundColor
				EndIf
				
				If IsNull(This.ColumnHeaderForeColor)
					lcForeColor = This.ColorToStrHexa( Iif(This.SetStyles, loColumnHeader.ForeColor, Rgb(0, 0, 0)) )
				Else
					lcForeColor = This.ColumnHeaderForeColor
				EndIf
				
				lcFontName = Padr(loColumnHeader.FontName, Len(xxxStylesProperties.ssFontName))
				lcFontSize = Transform(loColumnHeader.FontSize, "@L 999")
				lcFontItalic = Iif(loColumnHeader.FontItalic, "1", "0")
				lcFontBold = Iif(loColumnHeader.FontBold Or This.SetStyles = .F., "1", "0")
				lcFontUnderline = Padr(Iif(loColumnHeader.FontUnderline Or This.SetStyles = .F., "Single", ""), Len(xxxStylesProperties.ssFontUnderline))
				lcFontStrikeThru = Iif(loColumnHeader.FontStrikethru, "1", "0")
				lcFormat = Padr("", Len(xxxStylesProperties.ssFormat))
				lcFontFamily = Padr("Swiss", Len(xxxStylesProperties.ssFontFamily))
				lcPattern = Padr("Solid", Len(xxxStylesProperties.ssPattern))
				lcAlignH = Iif(This.SetStyles, This.GetColumnAlign("H", loColumnHeader.Alignment), "Left")
				lcAlignV = Iif(This.SetStyles, This.GetColumnAlign("V", loColumnHeader.Alignment), "Center")

				*- se o estilo j� existir "lcXmlStyle" retorna ""
				lcXmlStyle = This.addnewstyle( "h", 0, lnCol, ;
							lcAlignH, lcAlignV, lcFontName, lcFontFamily, ;
							lcFontSize, lcForeColor, lcFontBold, lcFontItalic, ;
							lcFontUnderline, lcFontStrikeThru, lcBackColor, lcPattern, ;
							lcFormat )

				If Not Empty(lcXmlStyle)
					lcXmlStyles = lcXmlStyles + This.crlf + lcXmlStyle
				Endif
			Endfor
		Endif

		Return lcXmlStyles
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure BuildRows
		Local lcXmlRows, lcDataType, lcDataColumn, lcAuxDataColumn, lnRow, lnCol, loColumn, loColumnHeader, loCurrentControl, ;
			lnPercent, lnCountRowSource, lcCountOption, laArrayTmp, lcComboOption, lcToolTipText, lnBytes, llHasDecimals, ;
			lnSetDecimals, lnRowFound, lnYear

		lcXmlRows = This.crlf
		lnRow = 0
		lnCol = 0
		lnBytes = 0
		lnSetDecimals = Set("Decimals")

		*- Adiciono a linha do Header no arquivo excel
		If This.GridObject.HeaderHeight > 0
			lcXmlRows = lcXmlRows + [   <Row ss:AutoFitHeight="0" ss:Height="] + Alltrim(Str(This.GridObject.HeaderHeight)) + [">] + This.crlf

			For lnCol = 1 To This.GridObject.ColumnCount
				loColumn = This.GetColumn(lnCol)
				loColumnHeader = This.GetColumnHeader(loColumn)

				If loColumn.Visible = .T.
					*- caso tenha tooltiptext
					lcToolTipText = ""
					If Not Empty(loColumnHeader.ToolTipText)
						lcToolTipText = [<Comment ss:Author="Rodrigo_Bruscain">] + ;
							[<ss:Data xmlns="http://www.w3.org/TR/REC-html40">] + ;
							[<Font html:Face="Tahoma" x:Family="Swiss" html:Color="#000000">] + Alltrim(loColumnHeader.ToolTipText) + [</Font>] + ;
							[</ss:Data>] + ;
							[</Comment>]
					Endif

					*- linha do header
					lcXmlRows = lcXmlRows + [     <Cell ss:StyleID="] + This.SeekStyle("000000", Transform(lnCol, "@L 999")) + ["><Data ss:Type="String">] + loColumnHeader.Caption + [</Data>] + lcToolTipText + [</Cell>] + This.crlf
				Endif
			Endfor

			lcXmlRows = lcXmlRows + [   </Row>] + This.crlf
		Endif

		lcXmlRows = lcXmlRows + This.crlf

		*- Adiciono a linha do Registro no arquivo excel
		Select (This.Alias)
		Go Top
		
		Scan
			lnRow = lnRow + 1
			lcXmlRows = lcXmlRows + [   <Row ss:AutoFitHeight="0">] + This.crlf

			*- percentual processado
			lnPercent = Int((lnRow / (This.RowCount - (Iif(This.GridObject.HeaderHeight > 0, 1, 0))) ) * 100)
			This.Progress(lnPercent)

			*- fa�o a varredura em todas as colunas
			For lnCol = 1 To This.GridObject.ColumnCount
				loColumn = This.GetColumn(lnCol)
				If Not loColumn.Visible
					Loop
				Endif

				*- Verifico o tipo de dado da coluna
				lcDataColumn = Evaluate(loColumn.ControlSource)
				loCurrentControl = This.GetCurrentControlObject(loColumn)

				*- se n�o tem objeto de controle na linha da coluna n�o levo a informa��o da tabela ao excel
				If Isnull(loCurrentControl)
					lcDataType = "String"
					lcDataColumn = ""
				Else
					Do Case
					Case Vartype(lcDataColumn) $ "N//Y"
						lcDataType = "Number"

						*- Se o currentcontrol da coluna for um combobox mostro o seu conteudo ao inves da posi��o numerica
						If Lower(loCurrentControl.BaseClass) = "combobox"
							Try
								Do Case
								*- Mostro o texto do value
								Case loCurrentControl.RowSourceType = 1
									lcDataType = "String"

									If Not Empty(loCurrentControl.RowSource)
										lcAuxDataColumn = Alltrim(loCurrentControl.RowSource)
										lcAuxDataColumn = Strtran(Strtran(Strtran(lcAuxDataColumn, " ,", ","), ", ", ","), " , ", ",")
										lcCountOption = Occurs(",", lcAuxDataColumn) + 1

										Dimension laArrayTmp[lcCountOption]
										For lnCountRowSource = 1 To lcCountOption
											lcComboOption = Substr(lcAuxDataColumn, 1, Iif(lnCountRowSource < lcCountOption, At(",", lcAuxDataColumn) - 1, Len(lcAuxDataColumn)) )
											lcAuxDataColumn = Strtran(lcAuxDataColumn, lcComboOption + Iif(lcCountOption >= 2, ",", ""), "")
											laArrayTmp[lnCountRowSource] =  lcComboOption
										Endfor

										lcDataColumn = Evaluate("laArrayTmp[" + Alltrim(Str(lcDataColumn)) + "]")
									Endif

								*- Mostro o texto do array do combo	
								Case loCurrentControl.RowSourceType = 5
									lcDataType = "String"

									*- Se for um array objeto ex: thisform.ArrayName ou MyObj.ArrayName
									If Occurs(".", loCurrentControl.RowSource) > 0
										lcObjArrayName = Substr(loCurrentControl.RowSource, 1, Rat(".", loCurrentControl.RowSource) - 1)

										*- Se for um array objeto publico
										If Type(lcObjArrayName) = "O"
											lcAuxDataColumn = loCurrentControl.RowSource + "[" + Alltrim(Str(lcDataColumn)) + "]"
										Else
											lcArrayName = Substr(loCurrentControl.RowSource, Rat(".", loCurrentControl.RowSource) + 1)
											lnCountObjectHierarchy = Occurs(".", Sys(1272, This.GridObject))
											lcAuxDataColumn = "This.GridObject" + Replicate(".Parent", lnCountObjectHierarchy) + "." + lcArrayName + "[" + Alltrim(Str(lcDataColumn)) + "]"
										Endif

									*- Array comum	
									Else
										lcAuxDataColumn = loCurrentControl.RowSource + "[" + Alltrim(Str(lcDataColumn)) + "]"
									Endif

									lcDataColumn = Evaluate(lcAuxDataColumn)


								*- Qualquer outro mostro o conteudo do campo e n�o o conteudo do array
								Otherwise
									lcDataColumn = lcDataColumn
								Endcase

							Catch To loError
							Endtry

							If Vartype(loError) = "O"
								Messagebox( "Combo array '" + loCurrentControl.RowSource + "' in column '" + loColumn.Name + "' not is valid", 48)
								Select (This.Alias)
								Go Top
								Return .F.
							Endif
						Else

							lnRowFound = Ascan(This._Fields, Iif("." $ loColumn.ControlSource, Substr(loColumn.ControlSource, At(".", loColumn.ControlSource) + 1), loColumn.ControlSource), -1, -1, 1, 15)
							If lnRowFound > 0 And This._Fields[lnRowFound, 4] > 0
								llHasDecimals = .T.
								Set Decimals To This._Fields[lnRowFound, 4]
							Else
								llHasDecimals = .F.
							Endif

						Endif

					Case Vartype(lcDataColumn) = "D"
						lcDataType = "DateTime"
						If Not Empty(Nvl(lcDataColumn, ""))
							lnYear = Iif(Year(lcDataColumn) < 1900, 1900, Year(lcDataColumn))
							lcAuxDataColumn = Str(lnYear, 4) + "-" + Transform(Month(lcDataColumn), "@L 99") + "-" + Transform(Day(lcDataColumn), "@L 99") + "T00:00:00.000"
							lcDataColumn = lcAuxDataColumn
						Else
							lcDataType = "String"
							lcDataColumn = ""
						Endif

					Case Vartype(lcDataColumn) = "T"
						lcDataType = "DateTime"
						If Not Empty(Nvl(lcDataColumn, ""))
							lnYear = Iif(Year(lcDataColumn) < 1900, 1900, Year(lcDataColumn))
							lcAuxDataColumn = Str(lnYear, 4) + "-" + Transform(Month(lcDataColumn), "@L 99") + "-" + Transform(Day(lcDataColumn), "@L 99") + ;
								"T" + Transform(Hour(lcDataColumn), "@L 99") + ":" + Transform(Minute(lcDataColumn), "@L 99") + ":" + Transform(Sec(lcDataColumn), "@L 99") + ".000"
							lcDataColumn = lcAuxDataColumn
						Else
							lcDataType = "String"
							lcDataColumn = ""
						Endif

					Case Vartype(lcDataColumn) = "L"
						lcDataType = "Number"
						lcDataColumn = Iif(lcDataColumn, 1, 0)

					Otherwise
						lcDataType = "String"
						If Isnull(lcDataColumn)
							lcDataColumn = ""
						Endif
					Endcase
				Endif

				*- removing invalid characters
				If lcDataType = "String" And ("<" $ lcDataColumn Or ">" $ lcDataColumn)
					lcDataColumn = Strtran(Strtran(lcDataColumn, "<", "["), ">", "]")
				Endif

				*- incluo a linha de dados
				lcXmlRows = lcXmlRows + [     <Cell ss:StyleID="] + This.SeekStyle(Transform(lnRow, "@L 999999"), Transform(lnCol, "@L 999")) + ["><Data ss:Type="] + lcDataType + [">] + Alltrim(Transform(lcDataColumn, "")) + [</Data></Cell>] + This.crlf

				*- devolvo o atributo original
				If llHasDecimals
					Set Decimals To lnSetDecimals
				Endif
			Endfor

			lcXmlRows = lcXmlRows + [   </Row>] + This.crlf
			lnBytes = lnBytes + Strtofile( lcXmlRows + This.crlf, This.File, 1)
			lcXmlRows = ""

		Endscan

		Return lnBytes
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure ColorToStrHexa(plnColor)
	
		Local lnDecimalColor
	
		lnDecimalColor = Substr(Transform(plnColor, '@0'), 5)
		Return "#" + Right(lnDecimalColor, 2) + Substr(lnDecimalColor, 3, 2) + Left(lnDecimalColor, 2)
	
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure GetColumn(plcColumnNumber)

		Local lnCol

		For lnCol = 1 To This.GridObject.ColumnCount
			If This.GridObject.Columns(lnCol).ColumnOrder = plcColumnNumber
				Return This.GridObject.Columns(lnCol)
			Endif
		EndFor
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure GetColumnAlign(plcWhat, plnAlignment, plcVartype)
	
		Local lcAlignment, lcAlignH, lcAlignV
		
		plcVartype = Evl(plcVartype, "")
		lcAlignment = Alltrim(Str(plnAlignment))

		*- Alinhamento Horizontal do texto da coluna/linha
		If plcWhat = "H"
			Do Case
			Case lcAlignment $ "0//4//7"
				lcAlignH = "Left"
			Case lcAlignment $ "1//5//8"
				lcAlignH = "Right"
			Case lcAlignment $ "2//6//9"
				lcAlignH = "Center"
			Otherwise
				lcAlignH = Iif(plcVartype $ "N//Y", "Right", "Left")
			Endcase

			lcAlignH = Padr(lcAlignH, Len(xxxStylesProperties.ssAlignH))
			Return lcAlignH
		Endif

		*- Alinhamento vertical do texto da coluna/linha
		If plcWhat = "V"
			Do Case
			Case lcAlignment $ "4//5//6"
				lcAlignV = "Top"
			Case lcAlignment $ "7//8//9"
				lcAlignV = "Bottom"
			Case lcAlignment $ "0//1//2"
				lcAlignV = "Center"
			Otherwise
				lcAlignV = "Center"
			Endcase

			lcAlignV = Padr(lcAlignV, Len(xxxStylesProperties.ssAlignV))
			Return lcAlignV
		EndIf
	
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure GetColumnHeader(ploColumn)
		
		Local loReturn, lnX
		loReturn = ""

		If ploColumn.ControlCount > 0
			For lnX = 1 To ploColumn.ControlCount
				If Lower(ploColumn.Controls(lnX).BaseClass) = "header"
					loReturn = ploColumn.Controls(lnX)
					Exit
				Endif
			Endfor
		Endif

		Return loReturn
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure GetCurrentControlObject(ploGridColumn)
		
		Local lcCurrentControl

		If Not Empty(ploGridColumn.DynamicCurrentControl)
			Try
				lcCurrentControl = Evaluate(ploGridColumn.DynamicCurrentControl)
			Catch
				lcCurrentControl = ploGridColumn.CurrentControl
			Endtry
		Else
			lcCurrentControl = ploGridColumn.CurrentControl
		Endif

		If Not Empty(lcCurrentControl)
			Return Evaluate("ploGridColumn." + lcCurrentControl)
		Else
			Return Null
		EndIf
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure HasColumnVisible
	
		Local lnCol, llReturn
		llReturn = .F.

		For lnCol = 1 To This.GridObject.ColumnCount
			If This.GridObject.Columns(lnCol).Visible
				llReturn = .T.
				Exit
			Endif
		Endfor

		Return llReturn
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	*/---------------------------------------------------------------------------------------------------/*
	*/ Descripton..: - Classe para converter o grid do vfp em um arquivo xml para o Excel.               /*
	*/				 - A grande vantagem na utiliza��o � que N�O NECESSITA DO EXCEL INSTALADO            /*
	*/                 pois em nenhum momento o Excel � instanciado para automa��o.                      /*
	*/                 Apesar de ser um arquivo xml, se encontra no padr�o Microsoft onde � reconhecido  /*
	*/                 pelo Excel como "Planilha XML 2003 (*.xml)". Dessa forma fica restrito o uso      /*
	*/                 para Excel 2003 ou superior.                                                      /*
	*/                                                                                                   /*
	*/				 - Se o Excel estiver instalado o icone do arquivo gerado ser� reconhecido		     /*
	*/				   pelo Excel e abrindo o arquivo ser� reconhecido como se fosse um XLS ou XLSX,     /*		
	*/                 ou seja, tudo ser� transparente para o Excel.                                     /*
	*/                                                                                                   /*
	*/               - Praticamente todos os recursos visuais do grid, headers, colunas e linhas         /*
	*/                 s�o tratados na exporta��o. Segue abaixo as propriedades reconhecidas:            /*
	*/                                                                                                   /*
	*/                 Header Properties                                                                 /*
	*/                 ---------------------------------                                                 /*  
	*/                 ToolTipText / HeaderHeight / Alignment / FontBold / FontItalic / FontUnderline /  /*
	*/                 FontStrikeThru  / FontName / FontSize / ForeColor / BackColor / Caption /         /*
	*/                                                                                                   /*
	*/                 Columns Properties                                                                /*
	*/                 ---------------------------------                                                 /*  
	*/				   ControlSource / BaseClass / InputMask / Seconds / RowHeight / Alignment / 		 /*
	*/				   FontBold / FontItalic / FontUnderline / FontStrikeThru  / FontName / FontSize /   /*
	*/				   ForeColor / FontBackColor / CurrentControl / DynamicFontBold / DynamicFontItalic  /*
	*/				   DynamicFontUnderline / DynamicFontStrikeThru / DynamicCurrentControl /         	 /*
	*/				   DynamicFontName / DynamicFontSize / DynamicForeColor / DynamicBackColor / 		 /*	
	*/				   ColumnCount / ColumnOrder / Width / Visible / Combobox.Alignment /                /*
	*/                 Combobox.RowSource / Combobox.RowSourceType					 					 /*
	*/                																					 /*
	*/                 Environment																		 /*			
	*/                 ---------------------------------                                                 /*  
	*/				   set date / set century / set hours                                                /*
	*/                                                                                                   /*
	*/
	*/				   Goals
	*/                 ------
	*/                 a) Possibilidade de gerar planilhas com mais de 65,535 linhas superando 
	*/                    a limita�ao nativa do VFP
	*/                 b) Converte um grid em planilha Excel assumindo 99% do visual do grid
	*/				   c) Easy to implement and it is not necessary to change your code
	*/                 d) Compativel com Excel 2003 ou superior
	*/                 e) Pode ser aberto pelo OpenOffice reduzindo erros de convers�o
	*/                 f) Ao abrir o arquivo pelo Excel � possivel salvar em outros formatos
	*/                 g) Nao precisa ter o Excel instalado
	*/
	*/                                                                                                   /*
	*/ Original Author......: Rodrigo Bruscain                                                           /*
	*/ Original Date........: 25/05/2013 (Original)                                                      /*
	*/ Country.....: Brazil - S�o Paulo - SP                                                             /*
	*/---------------------------------------------------------------------------------------------------/*
	Procedure Init

		This.crlf = Chr(13) + Chr(10)
		
		Local lcDateFormat, lcCentury

		AddProperty(This, "_Fields[1]")
		Dimension This._Fields[1,18]

		lcDateFormat = Set("Date")
		lcCentury = Iif(Set("century") = "ON", "yyyy", "yy")

		Do Case
		Case Inlist(lcDateFormat, "AMERICAN", "MDY")				&& month/day/year
			This.DateFormat = "mm/dd/" + lcCentury

		Case lcDateFormat = "ANSI"									&& year.month.day
			This.DateFormat = lcCentury + ".mm.dd"

		Case Inlist(lcDateFormat, "BRITISH", "DMY", "FRENCH") 		&& day/month/year
			This.DateFormat = "dd/mm/" + lcCentury

		Case lcDateFormat = "GERMAN"								&& day.month.year
			This.DateFormat = "dd.mm." + lcCentury

		Case lcDateFormat = "ITALIAN"								&& day-month-year
			This.DateFormat = "dd-mm-" + lcCentury

		Case Inlist(lcDateFormat, "JAPAN", "YMD")					&& year/month/day
			This.DateFormat = lcCentury + "/mm/dd"

		Case lcDateFormat = "USA"									&& month-day-year
			This.DateFormat = "mm-dd-" + lcCentury

		Otherwise
			This.DateFormat = "dd/mm/" + lcCentury
		Endcase

		*- version object
		This.Version = Createobject("empty")
		AddProperty(This.Version, "Version", "1.10")
		AddProperty(This.Version, "DateTime", "Sep.10.2017 3:59:41 AM")
		AddProperty(This.Version, "Author", "Rodrigo Duarte Bruscain")
		AddProperty(This.Version, "CountryAndCity", "kitchener ON - Canada")
		AddProperty(This.Version, "Url", "https://github.com/ExcelXml")
		AddProperty(This.Version, "Email", "bruscain@hotmail.com")
		AddProperty(This.Version, "Email2", "mattslay@jordanmachine.com")
		
	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure Progress(plnPercent)
	
		*-- Add any code here that you want to execute as processing scans over each row...

	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure Save(plcFile)

		Local lcCreatedDate, lnCol, lcSetPoint, loForm, lcAlias, lnRecNo, ;
			lcXmlStart, lcXmlDocumentProperties, lcXmlExcelWorkbook, lcStringStyles, ;
			lcXmlAllStyles, lcXmlFreezePanes, lcStringFilter, lcStringColumnWidth, ;
			lcXmlWorksheet_part1, lcXmlWorksheet_part2, lnBytes, loError

		plcFile = Evl(plcFile, "Book1")
		This.File = Evl(This.File, plcFile)
		This.File = This.File + Iif(Empty(Justext(This.File)), ".XML", "")

		If Empty(Alias())
			Messagebox("No table is open in the current work area.   ", 48)
			Return .F.
		Endif

		*- crio um grid virtual caso a nao exista um grid para conversao, 
		*- ou seja, estou convertendo somente a tabela
		If VarType(This.GridObject) != "O"
			loForm = CreateObject("form")
			loForm.AddObject("grid1", This.GridClass)
			loForm.Grid1.RecordSource =  Alias()
			loForm.Grid1.Visible = .T.
			loForm.Refresh()
			This.GridObject = loForm.Grid1
			This.SetStyles = .F.
		Endif

		*- environment
		If This.GridObject.RecordSourceType = 1
			This.Alias = This.GridObject.RecordSource
		Else
			This.Alias = Alias()
		Endif

		lnRecNo = Recno()
		Afields(This._Fields, This.Alias)


		*- Data da cria��o do arquivo excel
		lcCreatedDate = Str(Year(Date()), 4) + "-" + Transform(Month(Date()), "@L 99") + "-" + Transform(Day(Date()), "@L 99") + "T" + Time() + "Z"

		*- Numero de colunas v�lidas para o excel
		This.ColumnCount = 0
		For lnCol = 1 To This.GridObject.ColumnCount
			If This.GridObject.Columns(lnCol).Visible = .T.
				This.ColumnCount = This.ColumnCount + 1
			Endif
		Endfor

		*- Numero de linhas dispon�veis para o excel
		This.RowCount = 0
		Select (This.Alias)
		Count To This.RowCount
		Go Top

		If This.GridObject.HeaderHeight > 0
			This.RowCount = This.RowCount + 1
		Endif

		*- verifico se tudo esta ok para prosseguir
		If Isnull(This.GridObject) Or This.GridObject.ColumnCount <= 0 And This.hascolumnvisible()
			Return .F.
		Endif

		*- No Excel casas decimais obrigat�riamente trabalham com ponto "."
		lcSetPoint = Set("Point")
		Set Point To "."

		*- Inicio tratamento dos dados
		Text To lcXmlStart Textmerge Pretext 2 Noshow
			<?xml version="1.0" encoding="<<This.xmlEncoding>>"?>
			<?mso-application progid="Excel.Sheet"?>
			<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
			 xmlns:o="urn:schemas-microsoft-com:office:office"
			 xmlns:x="urn:schemas-microsoft-com:office:excel"
			 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
			 xmlns:html="http://www.w3.org/TR/REC-html40">
		ENDTEXT

		Text To lcXmlDocumentProperties Textmerge Pretext 2 Noshow
			 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
			  <Author><<iif(!empty(getenv("USERNAME")), getenv("USERNAME"), iif(!empty(getenv("COMPUTERNAME")), getenv("COMPUTERNAME"), "RODRIGO_BRUSCAIN"))>></Author>
			  <LastAuthor><<iif(!empty(getenv("USERNAME")), getenv("USERNAME"), iif(!empty(getenv("COMPUTERNAME")), getenv("COMPUTERNAME"), "RODRIGO_BRUSCAIN"))>></LastAuthor>
			  <Created><<lcCreatedDate>></Created>
			  <LastSaved><<lcCreatedDate>></LastSaved>
			  <Version>12.00</Version>
			 </DocumentProperties>
		ENDTEXT

		Text To lcXmlExcelWorkbook Textmerge Pretext 2 Noshow
			 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
			  <WindowHeight>8130</WindowHeight>
			  <WindowWidth>15135</WindowWidth>
			  <WindowTopX>120</WindowTopX>
			  <WindowTopY>45</WindowTopY>
			  <ProtectStructure>False</ProtectStructure>
			  <ProtectWindows>False</ProtectWindows>
			 </ExcelWorkbook>
		ENDTEXT


		*- Crio os estilos de cores/fontes/formato/etc das colunas
		*- Depois junto com o estilo padr�o todos os estilos encontrados
		*- Estilos s�o todas as format�es da c�lulas combinadas onde um estilo pode ser usado
		*- por v�rias c�luas ou por uma �nica c�lula.
		lcStringStyles = ""
		lcStringStyles = This.BuildHeadersStyles()					&&- Estilos do header
		lcStringStyles = lcStringStyles + This.buildcolumnsstyles()	&&- Estilos das linhas/colunas

		Text To lcXmlAllStyles Textmerge Pretext 2 Noshow
			 <Styles>
			  <Style ss:ID="Default" ss:Name="Normal">
			   <Alignment ss:Vertical="Center"/>
			   <Borders/>
			   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#000000"/>
			   <Interior/>
			   <NumberFormat/>
			   <Protection/>
			  </Style>
			  <<lcStringStyles>>
			 </Styles>
		ENDTEXT


		*- Congelando paineis na horizontal e vertical 
		Do Case
		*- Congelo a linha do header
		Case This.GridObject.LockColumns = 0 And (This.GridObject.HeaderHeight > 0 And This.LockHeader)
			Text To lcXmlFreezePanes Textmerge Pretext 2 Noshow
				   <FreezePanes/>
				   <FrozenNoSplit/>
				   <SplitHorizontal>1</SplitHorizontal>
				   <TopRowBottomPane>1</TopRowBottomPane>
				   <ActivePane>2</ActivePane>
				   <Panes>
				    <Pane>
				     <Number>3</Number>
				    </Pane>
				    <Pane>
				     <Number>2</Number>
				    </Pane>
				   </Panes>
			ENDTEXT

		*- congelo a linha do header e a coluna definida 
		Case This.GridObject.LockColumns > 0 And (This.GridObject.HeaderHeight > 0 And This.LockHeader)
			Text To lcXmlFreezePanes Textmerge Pretext 2 Noshow
				   <FreezePanes/>
				   <FrozenNoSplit/>
				   <SplitHorizontal>1</SplitHorizontal>
				   <TopRowBottomPane>1</TopRowBottomPane>
				   <SplitVertical><<alltrim(str(This.GridObject.LockColumns))>></SplitVertical>
				   <LeftColumnRightPane><<alltrim(str(This.GridObject.LockColumns))>></LeftColumnRightPane>
				   <ActivePane>0</ActivePane>
				   <Panes>
				    <Pane>
				     <Number>3</Number>
				    </Pane>
				    <Pane>
				     <Number>1</Number>
				    </Pane>
				    <Pane>
				     <Number>2</Number>
				    </Pane>
				    <Pane>
				     <Number>0</Number>
				    </Pane>
				   </Panes>			 
			ENDTEXT

		*- congelo somente a coluna definida
		Case This.GridObject.LockColumns > 0 And (This.GridObject.HeaderHeight = 0 Or Not This.LockHeader)
			Text To lcXmlFreezePanes Textmerge Pretext 2 Noshow
				   <FreezePanes/>
				   <FrozenNoSplit/>
				   <SplitVertical>2</SplitVertical>
				   <LeftColumnRightPane>2</LeftColumnRightPane>
				   <ActivePane>1</ActivePane>
				   <Panes>
				    <Pane>
				     <Number>3</Number>
				    </Pane>
				    <Pane>
				     <Number>1</Number>
				    </Pane>
				   </Panes>
			ENDTEXT

		Otherwise
			lcXmlFreezePanes = ""
		Endcase


		*- filtros na colunas
		lcStringFilter = ""
		If This.HasFilter And This.GridObject.HeaderHeight > 0
			Text To lcStringFilter Textmerge Pretext 2 Noshow
				<AutoFilter x:Range="R1C1:R<<alltrim(str(This.RowCount))>>C<<alltrim(str(This.ColumnCount))>>"
				 xmlns="urn:schemas-microsoft-com:office:excel">
				</AutoFilter>
			ENDTEXT
		Endif


		*- tratamento do nome da planilha
		This.SheetName = Chrtran(Alltrim(Substr(This.SheetName, 1, 31)), ':?][*/\', '')
		This.SheetName = Iif(Empty(This.SheetName), "Sheet1", This.SheetName)

		*- Monto a tabela
		lcStringColumnWidth = This.buildcolumnswidth()

		Text To lcXmlWorksheet_part1 Textmerge Pretext 2 Noshow
			 <Worksheet ss:Name="<<This.SheetName>>">
			  <Table ss:ExpandedColumnCount="<<This.ColumnCount>>" ss:ExpandedRowCount="<<This.RowCount>>" x:FullColumns="1"
			   x:FullRows="1" ss:DefaultRowHeight="<<alltrim(str(This.GridObject.RowHeight-3))>>">
			   <<lcStringColumnWidth>>
		ENDTEXT

		Text To lcXmlWorksheet_part2 Textmerge Pretext 2 Noshow
			  </Table>
			  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
			   <PageSetup>
			    <Header x:Margin="0.31496062000000002"/>
			    <Footer x:Margin="0.31496062000000002"/>
			    <PageMargins x:Bottom="0.78740157499999996" x:Left="0.511811024"
			     x:Right="0.511811024" x:Top="0.78740157499999996"/>
			   </PageSetup>
			   <Unsynced/>
			   <Print>
			    <ValidPrinterInfo/>
			    <PaperSizeIndex>9</PaperSizeIndex>
			    <HorizontalResolution>300</HorizontalResolution>
			    <VerticalResolution>300</VerticalResolution>
			   </Print>
			   <Selected/>
			   <<lcXmlFreezePanes>>
			   <ProtectObjects>False</ProtectObjects>
			   <ProtectScenarios>False</ProtectScenarios>
			  </WorksheetOptions>
			  <<lcStringFilter>> 	
			 </Worksheet>
			</Workbook>		
		ENDTEXT

		Try
			lnBytes = 0
			lnBytes = lnBytes + Strtofile("", This.File, 0)
			lnBytes = lnBytes + Strtofile( lcXmlStart + This.crlf, This.File, 1)
			lnBytes = lnBytes + Strtofile( lcXmlDocumentProperties + This.crlf, This.File, 1)
			lnBytes = lnBytes + Strtofile( lcXmlExcelWorkbook + This.crlf, This.File, 1)
			lnBytes = lnBytes + Strtofile( lcXmlAllStyles + This.crlf, This.File, 1)
			lnBytes = lnBytes + Strtofile( lcXmlWorksheet_part1 + This.crlf, This.File, 1)

			lnBytes = lnBytes + This.BuildRows()

			lnBytes = lnBytes + Strtofile( lcXmlWorksheet_part2 + This.crlf, This.File, 1)

			llReturn = Iif(lnBytes > 0, .T., .F.)

		Catch To loError
			If File(This.File)
				Erase (This.File)
			Endif

			Messagebox("An error occurred during the data exporting. " + Chr(13) + "Error: " + loError.Message, 16, "Exporting")

			llReturn = .F.
		Endtry

		*select xxxStylesRowCol
		*browse normal 
		*select xxxStylesProperties
		*browse normal 

		Set Point To &lcSetPoint

		If Used("xxxStylesProperties")
			Use In xxxStylesProperties
		Endif

		If Used("xxxStylesRowCol")
			Use In xxxStylesRowCol
		Endif

		If Used(This.Alias)
			Go lnRecNo
		Endif

		If Vartype(This.GridObject) <> "O"
			loForm.Release()
		Endif

		This.GridObject = .Null.

		If Used(This.Alias)
			Select (This.Alias)
		Endif

		*- abre o arquivo apos salva-lo
		If llReturn And This.OpenAfterSaving
			Declare Integer ShellExecute In SHELL32.Dll As WinAPI_OpenAfterSavingExcelXml;
				Integer HndWin, String cAction, String cFileName, ;
				String cParams, String cDir, Integer nShowWin

			WinAPI_OpenAfterSavingExcelXml(0, "OPEN", This.File, "", "", 1)
			Clear Dlls "WinAPI_OpenAfterSavingExcelXml"
		Endif

		Return llReturn

	Endproc


	*|================================================================================ 
	*| ExcelXml::
	Procedure SeekStyle(plcRow, plcCol)

		Local lcReturn
		lcReturn = ""

		*- se nao aplica estilos
		If Not This.SetStyles And plcRow > "000001"
			plcRow = "000001"
		Endif

		*- Procuro um estilo para a celula, caso nao encontre aplico o padr�o.
		*- Teoricamente todas as celulas deve ter um estilo e n�o o padr�o.		 
		If Seek(plcRow + plcCol, "xxxStylesRowCol", "idxRowCol")
			lcReturn = xxxStylesRowCol.ssCode
		Else
			lcReturn = "Default"
		Endif

		Return lcReturn

	EndProc
	
	
	*---------------------------------------------------------------------------------------
	* After creating XML file in the Save() method, you can call this method and pass filename of XML file,
	* to use Excell to open the XML file and convert it to an XLSX file.
	*  Values for lnFileFormat:
	* 	51 = xlOpenXMLWorkbook (without macro's in 2007-2013, xlsx)
	* 	52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2013, xlsm)
	* 	50 = xlExcel12 (Excel Binary Workbook in 2007-2013 with or without macro's, xlsb)
	* 	56 = xlExcel8 (97-2003 format in Excel 2007-2013, xls)
	Procedure ConvertXmlToXlsx(tcFilename, tnFileFormat, tlOpenAfterExporting)

		Local loExcel as "Excel.Application"
		Local lcNewFilename, lnFileFormat, loWorkBook, lcSafety

		loExcel = Createobject("Excel.Application")

		If !IsObject(loExcel)
			This.cErrorMessage = "Error starting Excel."
			Return .F.
		Endif

		If !File(tcFileName)
			This.cErrorMessage = "File not found: " + tcFilename
			Return .F.
		Else
			loWorkBook = loExcel.Application.Workbooks.Open(tcFileName)
		EndIf

		lnFileFormat = Evl(tnFileFormat, 51) && 51 = xlsx as default
		
		If (".xml" $ tcFilename)
			lcNewFilename = Strtran(tcFilename, ".xml", ".xlsx", 1, 99, 1)
			loWorkBook.SaveAs(lcNewFilename, lnFileFormat)
			lcSafety = Set("Safety")
			Set Safety Off
			Delete File (tcFileName)
			Set Safety &lcSafety
		Endif

		If tlOpenAfterExporting
			loExcel.Visible = .T.
		Else
			loExcel.Quit()
		EndIf
		
	Endproc	

Enddefine
