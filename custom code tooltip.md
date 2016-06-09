######SQL tooltip expression code (Series Properties > Series Data > Tooltip)

######Websites Used:

LookUpSet - Group By: http://stackoverflow.com/questions/23059364/lookupset-group-by
<br />

```sql
--Returns an Object and String Variable.
=Code.test(LookupSet(Fields!sCompositeOf.Value & Fields!Variable.Value, Fields!sCompositeOf.Value & Fields!Variable.Value, Trim(CStr(Fields!sVariableName.Value)) & "~"& CStr(Fields!fValue.Value), "DataSet1"),Trim(CStr(Fields!sVariableName.Value)) & "~"& CStr(Fields!fValue.Value))
```
<br />

######Update splits between mouseover value and other values within the dataset for stacked barcharts (Report Properties > Code)

######Websites Used:
*Excel VBA Dictionary – A Complete Guide: http://excelmacromastery.com/Blog/index.php/vba-dictionary/#A_Simple_Example_of_using_the_VBA_Dictionary*<br />
*Dictionaries: http://www.snb-vba.eu/VBA_Dictionary_en.html*<br />
*Data Dictionary in VBA - Complete Syntax Documentation: https://sites.google.com/site/beyondexcel/project-updates/datadictionaryinvba-completesyntaxdocumentation*<br />
*Does VBA have Dictionary Structure?: http://stackoverflow.com/questions/915317/does-vba-have-dictionary-structure*

######Breakthrough Websites Used:
*SSRS 2008 R2 custom code clear dictionary in render lifecycle event: http://stackoverflow.com/questions/18927516/ssrs-2008-r2-custom-code-clear-dictionary-in-render-lifecycle-event?rq=1*<br />
*SSRS custom code and variables life: http://stackoverflow.com/questions/16564241/ssrs-custom-code-and-variables-life*<br />
*System.Collections.Generic.Dictionary<TKey,TValue> Class: https://www.gnu.org/projects/dotgnu/pnetlib-doc/System/Collections/Generic/DictionaryTKeyTValue.html#Dictionary%3CTKey%2CTValue%3E.System.Collections.Generic.IDictionary%3CTKey%2CTValue%3E.Values%20Property*<br />
<br />

```vbnet
Public Shared Dim Totals As New System.Collections.Generic.Dictionary(Of String, Decimal)

Public Function WipeKeys() as Decimal 'Clear Data from Dictionary (this will clear the cached object as well)
  Totals.Clear()
  Return 0D
End Function

Public Function AddKeys(key as String,item as Decimal)
Totals.Add(key,item)
End Function

Public Function PrintKeys(val as String)
dim text as string
text =  "Mouseover Value" + vbcrlf+  val & ": " & cstr(Totals(val)) + vbcrlf + vbcrlf+"""Other Values"""+ vbcrlf
 For Each Entry As string In Totals.Keys()
	If val <> Entry
		text = text +Entry +": "+ cstr(Totals(Entry)) + vbcrlf
	End if
 Next
return text
End Function

Public Function PrintItems()
 For Each Entry As string In Totals.Keys()
            Return Totals(Entry)
 Next
End Function

Public Function CheckKey(val as string)
dim val1 as Boolean
val1 = Totals.ContainsKey(val)
return val1
End Function 

Public Function EditKeyVal(key as string,val as Decimal)
Totals(key) = Totals(key) +val
End Function


Function test (ByVal items As Object(),mouseover as String) as String
If items Is Nothing Then
Return Nothing
End If

WipeKeys() 
Dim item as String
Dim check as Boolean
For Each item in items
check = CheckKey(Split(item,"~")(0))
If check = False Then
AddKeys(Split(item,"~")(0),Split(item,"~")(1))
Else
EditKeyVal(Split(item,"~")(0),Split(item,"~")(1))
End If
Next item
return PrintKeys(Split(mouseover,"~")(0))


End Function
```
