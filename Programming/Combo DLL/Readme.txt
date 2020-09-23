CmbAcmpl
Windows Combo box enhance active X Dll
With this Active X Dll you can enhance a standard windows combo box
--------
Functions
1.Sub Drop(Combo As Object, n As Boolean)
Action:		Drops the the DropDown list of a ComboBox
Parameters:	Combo as object =Reference to the Combo
Returns:	None

2. Sub AutoComplete(Combo As Object)
Action:		Autocomplete function in the typed text
Parameters:	Combo as object =Reference to the Combo
Returns:	None
Notes:		Use it in the Event Change of the combobox

3. Function FindString(Combo As Object, mString As String, SelectIt As Boolean) As Integer
Action:		Finds an item of the combo box that matches the string 'mString'
Parameters:	Combo as object =Reference to the Combo
		mString as string= the string that is used to search the combo box items
		SelectedIt as boolean=If True then the function selects the found item
Returns:	The Index of the item Found

4. Function GetDroppedState(Combo As Object) As Boolean
Action:		Gets the state of the DropDown list of a combo box
Parameters:	Combo as object =Reference to the Combo
Returns:	True or False

5. Function GetDroppedWidth(Combo As Object) As Long
Action: 	Returns the width in pixels of the DropDown List of a Combobox
Parameters:	Combo as object =Reference to the Combo
Returns:	The width in pixels of the DropDown List of a Combobox

6. Function GetEditBoxHeight(Combo As Object) As Long
Action:		Returns the height in pixels of the edit box which resides in a combobox
Parameters:	Combo as object =Reference to the Combo
Returns:	Returns the height in pixels of the edit box which resides in a combobox

7. Sub SetDroppedWidth(Combo As Object, nWidth As Integer)
Actions:	Sets the width in pixels of  the DropDown List of a Combobox. It cannot be 
		smaller from the width of the combobox
Parameters:	Combo as object =Reference to the Combo
		nWidth as long (the new width in pixels)
Returns:	None

8. Function SetEditBoxHeight(Combo As Object, nHeight As Long) As Long
Action:		Not working


9. Sub SetItemHeight(Combo As Object, nItem As Integer)
Action:		Not Working


10. Sub SetMaxLen(Combo As Object, MaxChars As Long)
Action:		Set the maximum number of characters that can be typped into the combo box
Parameters:	Combo as object =Reference to the Combo
		MaxChars as integer, Maximum number of characters. If <1 then the combo box is 
		reset to the windows default
Returns:	None


Notes on the DLL
The auto complete function is not mine, but unfortunately I cannot remember the man who made it and added it on the Planet Source code. I thank him any way.

To use the dll you can add the following:
1st 	add a reference to the the Dll from the project menu/References. If it is not listed in 
	the list then you can Browse to find it.
2nd add the code:
	
	Dim cb as new clsAutoComplete
	Now you can use the functions as this
	
	ret&=cb.GetDroppedWidth(combo1) 'To get the width of the Dropdown
	cb.Drop(Combo1) 'to Drop the dropdown list
To use the Autocomplete do the folowing
In the Combo_Change event add the code
	Dim cb as new clsAutoComplete
	cb.autocomplete(combo1)



Any comments to 
anast@otenet.gr

