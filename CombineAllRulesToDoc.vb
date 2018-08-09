'iLogic
Public Sub Main()

    Dim oAsmDoc As AssemblyDocument = ThisApplication.ActiveDocument
    
    Call rilogic(oAsmDoc, original_text, replacement_text)
    
    Dim item As Document
    For Each item In oAsmDoc.AllReferencedDocuments
        Call rilogic(item, original_text, replacement_text)
    Next item
    
    MsgBox ("Done")
End Sub

Public Sub rilogic(docfile As Document, str1 As String, str2 As String)
		Dim iLogicObject As Object = iLogicVb.Automation
      
    	Dim rules As Object
    	rules = iLogicObject.rules(docfile)

    	If (Not rules Is Nothing) Then
        	Dim rule As Object
        	For Each rule In rules
				'MsgBox(docfile.DisplayName & " open looking in rule " & rule.name)
				
				Dim file As System.IO.StreamWriter
				
				Dim v As String = rule.name
				
				file = My.Computer.FileSystem.OpenTextFileWriter("c:\AllRules.txt", True)

				file.WriteLine("Document: " & docfile.DisplayName & " - Rule: " & rule.name)
				file.WriteLine(" ")
				file.WriteLine(CStr(rule.Text))
				file.WriteLine(" ")
				file.Close()
				
            	rule.Text = Replace(rule.Text, str1, str2)
        	Next rule
    	End If
End Sub
