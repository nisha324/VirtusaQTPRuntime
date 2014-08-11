Dim CommandObj

Function SendObject (ByVal Obj, ByVal identifire)
	
	identifire = resolvedIdentifire(identifire)
	
	Select Case Obj

	
	End Select

End Function

Function resolvedIdentifire(identifire)
		If (identifire<>"") Then
			arr=Split(identifire, "_PARAM:")
		'Conversion of : Id_PARAM:Sign in -> Sign in 
			resolvedIdentifire = arr(1)
		Else
			resolvedIdentifire = identifire
		End If
End Function
