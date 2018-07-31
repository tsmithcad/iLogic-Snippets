Sub Main
	Call TraverseAssemblySample()
End Sub

Public Sub TraverseAssemblySample() 
    ' Get the active assembly. 
    Dim oAsmDoc As AssemblyDocument 
    oAsmDoc = ThisApplication.ActiveDocument 
    'Debug.Print oAsmDoc.DisplayName 

    ' Call the function that does the recursion. 
    Call TraverseAssembly(oAsmDoc.ComponentDefinition.Occurrences, 1) 
End Sub 

Private Sub TraverseAssembly(Occurrences As ComponentOccurrences, StackQuantity As Integer) 
    ' Iterate through all of the occurrence in this collection.  This 
    ' represents the occurrences at the top level of an assembly. 
    Dim oOcc As ComponentOccurrence 
    For Each oOcc In Occurrences 
        ' Print the name of the current occurrence. 
			On Error Resume Next
			
			'MsgBox(oOccurrence.Name)
			
			Dim file As System.IO.StreamWriter
			file = My.Computer.FileSystem.OpenTextFileWriter("c:\FileNames.txt", True)
			
			If InStr(1,oOcc.Name,":1") Then
				file.WriteLine(oOcc.Name)
			End If
			
			file.Close()
			
			i += 1 

        ' Check to see if this occurrence represents a subassembly 
        ' and recursively call this function to traverse through it. 
        If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then 
            Call TraverseAssembly(oOcc.SubOccurrences, StackQuantity + 1) 
        End If 
    Next 
End Sub
