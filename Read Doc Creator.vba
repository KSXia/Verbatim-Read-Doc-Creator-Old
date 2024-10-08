' ---Read Doc Creator v1.0.7---
' Updated on 2024-08-23.
' Basic Edition: This edition of the Read Doc Creator only creates the read doc and does not have any mechanisms regarding the saving of the read doc.
' This macro consists of 6 sub procedures.
' https://github.com/KSXia/Verbatim-Read-Doc-Creator-Old
' Thanks to Truf for creating and providing the original code for activating invisibility mode! You can find Truf's macros on his website at https://debate-decoded.ghost.io/leveling-up-verbatim/

' Sub procedure 1 of 6: Read Doc Creator Core
Sub CreateReadDoc(EnableInvisibilityMode As Boolean, EnableFastInvisibilityMode As Boolean)
	Dim DeleteStyles As Boolean
	Dim StylesToDelete() As Variant
	Dim DeleteLinkedCharacterStyles As Boolean
	Dim LinkedCharacterStylesToDelete() As Variant
	Dim DeleteForReferenceHighlightingInInvisibilityMode As Boolean
	Dim DeleteForReferenceCardHighlightingInNormalMode As Boolean
	Dim ForReferenceHighlightingColor As String
	
	' ---USER CUSTOMIZATION---
	' <<SET THE STYLES TO DELETE HERE!>>
	' Add the names of styles that you want to delete to the list in the StylesToDelete array. Make sure that the name of the style is in quotation marks and that each term is separated by commas!
	' If the list is empty, this macro will still work, but no styles will be deleted.
	StylesToDelete = Array("Undertag")
	
	' If DeleteStyles is set to True, the styles listed in the StylesToDelete array will be deleted. If DeleteStyles is set to False, the styles listed in the StylesToDelete array will not be deleted.
	' If you want to disable the deletion of the styles listed in the StylesToDelete array, set DeleteStyles to False.
	DeleteStyles = True
	
	' <<SET THE LINKED CHARACTER STYLES TO DELETE HERE!>>
	' A linked style will either apply the style to the entire paragraph or a selection of words depending on what you have selected. If you have clicked on a paragraph and have selected no text or have selected the entire paragraph, it will apply the paragraph variant of the style. If you have selected a subset of the paragraph, it will apply the character variant of the style to your selection. The options in this section control whether this macro will delete the instances of character variants of linked styles and which linked styles this macro will operate on.
	
	' If DeleteLinkedCharacterStyles is set to True, the character variants of the linked styles listed in the LinkedCharacterStylesToDelete array will be deleted. If DeleteLinkedCharacterStyles is set to False, they will not be deleted.
	DeleteLinkedCharacterStyles = False
	
	' Add the names of linked styles that you want to delete the character variant of to the list in the LinkedCharacterStylesToDelete array. Make sure that the name of the style is in quotation marks and that each term is separated by commas!
	' If the list is empty, this macro will still work, but no character variants of linked styles will be deleted.
	LinkedCharacterStylesToDelete = Array()
	
	' <<SET WHETHER TO DELETE HIGHLIGHTED TEXT IN "For Reference" CARDS HERE!>>
	' If DeleteForReferenceCardsForInvisibilityMode is set to True, text highlighted in your "For Reference" highlighting color (which is set in the ForReferenceHighlightingColor option below) will be deleted when the read doc is set to have invisibility mode activated.
	DeleteForReferenceHighlightingInInvisibilityMode = False
	' If DeleteForReferenceCardsForNormalMode is set to True, text highlighted in your "For Reference" highlighting color (which is set in the ForReferenceHighlightingColor option below) will be deleted when the read doc is not set to have invisibility mode activated.
	DeleteForReferenceCardHighlightingInNormalMode = False
	
	' <<SET THE COLOR YOU USE FOR "For Reference" CARDS HERE!>>
	' Set ForReferenceHighlightingColor to the name of the highlighting color you use for "For Reference" cards.
	' WARNING: This highlighting color MUST ONLY be used for "For Reference" cards and nothing that you are reading! If this is not the case, DISABLE the function to delete highlighting for "For Reference" cards by setting DeleteForReferenceHighlightingInInvisibilityMode and DeleteForReferenceCardHighlightingInNormalMode to False.
	'
	' These are the names of the highlighting colors in the each row of the highlighting color selection menu, listed from left to right:
	' First row: Yellow, Bright Green, Turquoise, Pink, Blue
	' Second row: Red, Dark Blue, Teal, Green, Violet
	' Third row: Dark Red, Dark Yellow, Dark Gray, Light Gray, Black
	' MAKE SURE TO USE THIS EXACT CAPITALIZATION AND SPELLING!
	ForReferenceHighlightingColor = "Light Gray"
	
	' ---INITIAL VARIABLE SETUP---
	Dim OriginalDoc As Document
	' Assign the original document to a variable
	Set OriginalDoc = ActiveDocument
	
	' Check if the original document has previously been saved
	If OriginalDoc.Path = "" Then
		' If the original document has not been previously saved:
		MsgBox "The current document must be saved at least once. Please save the current document and try again.", Title:="Error in Creating Read Doc"
		Exit Sub
	End If
	
	' Assign the original document name to a variable
	Dim OriginalDocName As String
	OriginalDocName = OriginalDoc.Name
	
	' ---INITIAL GENERAL SETUP---
	' Disable screen updating for faster execution
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False
	
	' ---VARIABLE SETUP---
	Dim ReadDoc As Document
	
	' If the doc has been previously saved, create a copy of it to be the read doc
	Set ReadDoc = Documents.Add(OriginalDoc.FullName)
	
	Dim GreatestStyleIndex As Integer
	GreatestStyleIndex = UBound(StylesToDelete) - LBound(StylesToDelete)
	
	Dim GreatestLinkedCharacterStyleIndex As Integer
	GreatestLinkedCharacterStyleIndex = UBound(LinkedCharacterStylesToDelete) - LBound(LinkedCharacterStylesToDelete)
	
	' ---STYLE DELETION SETUP---
	' Disable error prompts in case one of the styles set to be deleted isn't present
	On Error Resume Next
	
	' ---PRE-PROCESSING FOR STYLE DELETION---
	' Use Find and Replace to replace paragraph marks in the character variants of linked styles set for deletion with paragraph marks in Tag style.
	' This ensures all paragraph marks in lines or paragraphs that have character variants of linked styles set to be delted are in Tag style so they do not get deleted in the style deletion stage of this macro.
	' Otherwise, lines ending in character variants of linked styles set to be delted may have their paragraph mark deleted and have the following line be merged into them, which can mess up the formatting of the line.
	If DeleteLinkedCharacterStyles = True Then
		Dim CurrentLinkedCharacterStyleNameToProcessIndex As Integer
		For CurrentLinkedCharacterStyleNameToProcessIndex = 0 To GreatestLinkedCharacterStyleIndex Step 1
			LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleNameToProcessIndex) = LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleNameToProcessIndex) & " Char"
		Next CurrentLinkedCharacterStyleNameToProcessIndex
		
		Dim CurrentLinkedCharacterStyleToProcessIndex As Integer
		For CurrentLinkedCharacterStyleToProcessIndex = 0 To GreatestLinkedCharacterStyleIndex Step 1
			Dim LinkedCharacterStyleToProcess As Style
			
			Set LinkedCharacterStyleToProcess = ReadDoc.Styles(LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleToProcessIndex))
			
			With ReadDoc.Content.Find
				.ClearFormatting
				.Text = "^p"
				.Style = LinkedCharacterStyleToProcess
				.Replacement.ClearFormatting
				.Replacement.Text = "^p"
				.Replacement.Style = "Tag Char"
				.Format = True
				' Ensure various checks are disabled to have the search properly function
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				' Delete all text with the style to delete
				.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
			End With
		Next CurrentLinkedCharacterStyleToProcessIndex
	End If
	
	' ---STYLE DELETION---
	If DeleteStyles = True Then
		Dim CurrentStyleToDeleteIndex As Integer
		For CurrentStyleToDeleteIndex = 0 To GreatestStyleIndex Step 1
			Dim StyleToDelete As Style
			
		' Specify the style to be deleted and delete it
			Set StyleToDelete = ReadDoc.Styles(StylesToDelete(CurrentStyleToDeleteIndex))
			
			' Use Find and Replace to remove text with the specified style and delete it
			With ReadDoc.Content.Find
				.ClearFormatting
				.Style = StyleToDelete
				.Replacement.ClearFormatting
				.Replacement.Text = ""
				.Format = True
				' Disable checks in the find process for optimization
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				' Delete all text with the style to delete
				.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
			End With
		Next CurrentStyleToDeleteIndex
	End If
	
	If DeleteLinkedCharacterStyles = True Then
		Dim CurrentLinkedCharacterStyleToDeleteIndex As Integer
		For CurrentLinkedCharacterStyleToDeleteIndex = 0 to GreatestLinkedCharacterStyleIndex Step 1
			Dim LinkedCharacterStyleToDelete As Style
			
			' Specify the linked style to delete the character variants of
			Set LinkedCharacterStyleToDelete = ReadDoc.Styles(LinkedCharacterStylesToDelete(CurrentLinkedCharacterStyleToDeleteIndex))
			
			' Use Find and Replace to remove text with the character variants of the specified linked style and delete it
			With ReadDoc.Content.Find
				.ClearFormatting
				.Style = LinkedCharacterStyleToDelete
				.Replacement.ClearFormatting
				.Replacement.Text = ""
				.Format = True
				' Disable checks in the find process for optimization
				.MatchCase = False
				.MatchWholeWord = False
				.MatchWildcards = False
				.MatchSoundsLike = False
				.MatchAllWordForms = False
				' Delete all text with the style to delete
				.Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
			End With
		Next CurrentLinkedCharacterStyleToDeleteIndex
	End If
	
	' ---POST STYLE DELETION PROCESSES---
	' Re-enable error prompts
	On Error GoTo 0
	
	' ---DELETE HIGHLIGHTED WORDS IN "For Reference" CARDS---
	If EnableInvisibilityMode = False And DeleteForReferenceCardHighlightingInNormalMode Then
		Call DeleteForReferenceCardHighlighting(ReadDoc, ForReferenceHighlightingColor)
	ElseIf EnableInvisibilityMode = True And DeleteForReferenceHighlightingInInvisibilityMode Then
		Call DeleteForReferenceCardHighlighting(ReadDoc, ForReferenceHighlightingColor)
	End If
	
	' ---DESTRUCTIVE INVISIBILITY MODE---
	If EnableInvisibilityMode And EnableFastInvisibilityMode Then
		Call EnableDestructiveInvisibilityMode(ReadDoc, True)
	ElseIf EnableInvisibilityMode Then
		Call EnableDestructiveInvisibilityMode(ReadDoc, False)
	End If
	
	' ---FINAL PROCESSES---
	' Re-enable screen updating and alerts
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True
End Sub

' Sub procedure 2 of 6: Invisibility Mode Enabler
' Thanks to Truf for creating and providing the original code for activating invisibility mode! This sub procedure is based on Truf's "InvisibilityOn" and "InvisibilityOnFast" sub procedures. You can find Truf's macros on his website at https://debate-decoded.ghost.io/leveling-up-verbatim/
Sub EnableDestructiveInvisibilityMode(TargetDoc As Document, UseFastMode As Boolean)
	' Move the cursor to the beginning of the document
	TargetDoc.Content.Select
	Selection.HomeKey Unit:=wdStory
	
	' Replace all paragraph marks with highlighted and bolded paragraph marks
	With TargetDoc.Content.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Text = "^p"
		.Replacement.Text = "^p"
		.Replacement.Style = "Underline"
		.Replacement.Highlight = True
		.Replacement.Font.Bold = True
		.MatchWildcards = False
		.Execute Replace:=wdReplaceAll
	End With
	
	' Delete non-highlighted "Normal" text
	With TargetDoc.Content.Find
		.ClearFormatting
		.Style = "Normal"
		.Highlight = False
		.Font.Bold = False
		.Replacement.ClearFormatting
		.Text = ""
		.Replacement.Text = " "
		.Execute Replace:=wdReplaceAll
	End With
	
	' Delete non-highlighted "Underline" text
	With TargetDoc.Content.Find
		.ClearFormatting
		.Style = "Underline"
		.Highlight = False
		.Replacement.ClearFormatting
		.Text = ""
		.Replacement.Text = " "
		.Execute Replace:=wdReplaceAll
	End With
	
	' Delete non-highlighted "Emphasis" text
	With TargetDoc.Content.Find
		.ClearFormatting
		.Style = "Emphasis"
		.Highlight = False
		.Replacement.ClearFormatting
		.Text = ""
		.Replacement.Text = " "
		.Execute Replace:=wdReplaceAll
	End With
	
	' Remove extra spaces between paragraph marks
	With TargetDoc.Content.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Text = "^p ^p"
		.Replacement.Text = ""
		.Replacement.Highlight = False
		.Execute Replace:=wdReplaceAll
	End With
	
	' Remove consecutive spaces in non-highlighted text
	With TargetDoc.Content.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Text = "( ){2,}"
		.Highlight = False
		.MatchWildcards = True
		.Replacement.Text = " "
		.Execute Replace:=wdReplaceAll
	End With
	
	' Remove spaces at the beginning of paragraphs
	With TargetDoc.Content.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Text = "^p "
		.Replacement.Text = "^p"
		.MatchWildcards = False
		.Execute Replace:=wdReplaceAll
	End With
	
	' Remove consecutive paragraph marks in non-highlighted text
	With TargetDoc.Content.Find
		.ClearFormatting
		.Replacement.ClearFormatting
		.Text = "^13{1,}"
		.Replacement.Text = "^p"
		.MatchWildcards = True
		.Execute Replace:=wdReplaceAll
	End With
	
	If Not UseFastMode Then
		Dim i As Long
		
		' Remove line breaks surrounded on both sides by highlighted text
		Dim para As Paragraph
		Dim rng As Range
		Dim highlighted As Boolean
		
		For Each para In TargetDoc.Paragraphs
			Set rng = para.Range
			rng.MoveEnd wdCharacter, -1 ' Ignore the paragraph mark
			
			' Check if the current paragraph contains highlighted text
			highlighted = False
			For i = 1 To rng.Characters.Count
				If rng.Characters(i).HighlightColorIndex <> wdNoHighlight Then
					highlighted = True
					Exit For
				End If
			Next i
			
			' Check if the next paragraph exists and contains highlighted text
			Dim nextHighlighted As Boolean
			nextHighlighted = False
			If Not para.Next Is Nothing Then
				For i = 1 To para.Next.Range.Characters.Count - 1 ' Ignore the paragraph mark
					If para.Next.Range.Characters(i).HighlightColorIndex <> wdNoHighlight Then
						nextHighlighted = True
						Exit For
					End If
				Next i
			End If
			
			' If both paragraphs contain highlighted text, join them
			If highlighted And nextHighlighted Then
				rng.InsertAfter " " ' Insert a space after the current paragraph
				para.Range.Characters.Last.Delete ' Delete the paragraph mark
			End If
		Next para
	End If
	
	' Clean up and suppress errors
	TargetDoc.Content.Find.ClearFormatting
	TargetDoc.Content.Find.MatchWildcards = False
	TargetDoc.Content.Find.Replacement.ClearFormatting
	TargetDoc.ShowGrammaticalErrors = False
	TargetDoc.ShowSpellingErrors = False
End Sub

' Sub procedure 3 of 6: Delete Highlighting in "For Reference" Cards
Sub DeleteForReferenceCardHighlighting(Doc As Document, ForReferenceHighlightingColor As String)
	Dim ForReferenceHighlightingColorEnum As Long
	' This code for converting highlighting color to enum is from Verbatim 6.0.0's "Standardize Highlighting With Exception" functon
	Select Case ForReferenceHighlightingColor
		Case Is = "None"
			ForReferenceHighlightingColorEnum = wdNoHighlight
		Case Is = "Black"
			ForReferenceHighlightingColorEnum = wdBlack
		Case Is = "Blue"
			ForReferenceHighlightingColorEnum = wdBlue
		Case Is = "Bright Green"
			ForReferenceHighlightingColorEnum = wdBrightGreen
		Case Is = "Dark Blue"
			ForReferenceHighlightingColorEnum = wdDarkBlue
		Case Is = "Dark Red"
			ForReferenceHighlightingColorEnum = wdDarkRed
		Case Is = "Dark Yellow"
			ForReferenceHighlightingColorEnum = wdDarkYellow
		Case Is = "Light Gray"
			ForReferenceHighlightingColorEnum = wdGray25
		Case Is = "Dark Gray"
			ForReferenceHighlightingColorEnum = wdGray50
		Case Is = "Green"
			ForReferenceHighlightingColorEnum = wdGreen
		Case Is = "Pink"
			ForReferenceHighlightingColorEnum = wdPink
		Case Is = "Red"
			ForReferenceHighlightingColorEnum = wdRed
		Case Is = "Teal"
			ForReferenceHighlightingColorEnum = wdTeal
		Case Is = "Turquoise"
			ForReferenceHighlightingColorEnum = wdTurquoise
		Case Is = "Violet"
			ForReferenceHighlightingColorEnum = wdViolet
		Case Is = "White"
			ForReferenceHighlightingColorEnum = wdWhite
		Case Is = "Yellow"
			ForReferenceHighlightingColorEnum = wdYellow
		Case Else
			ForReferenceHighlightingColorEnum = wdNoHighlight
	End Select
	' End of code based on Verbatim 6.0.0 functions
	
	With Doc.Content
		With .Find
			.ClearFormatting
			.Highlight = True
			.Text = ""
			.Replacement.ClearFormatting
			.Replacement.Text = ""
			.Format = True
			' Disable checks in the find process for optimization
			.MatchCase = False
			.MatchWholeWord = False
			.MatchWildcards = False
			.MatchSoundsLike = False
			.MatchAllWordForms = False
			' Modify the search process settings
			.Forward = True
			.Wrap = wdFindStop
			End With
			' Delete all text with the "For Reference" highlighting color
			Do While .Find.Execute = True
				If .HighlightColorIndex = ForReferenceHighlightingColorEnum Then .Delete
			Loop
	End With
End Sub

' Sub procedure 4 of 6: Trigger for Read Doc Creator
Sub CreateNormalReadDoc()
	Call CreateReadDoc(False, False)
End Sub

' Sub procedure 5 of 6: Trigger for Read Doc Creator
Sub CreateReadDocWithInvisibilityMode()
	Call CreateReadDoc(True, False)
End Sub

' Sub procedure 6 of 6: Trigger for Read Doc Creator
Sub CreateReadDocWithFastInvisibilityMode()
	Call CreateReadDoc(True, True)
End Sub
' <<END Read Doc Creator>>
