<%
	Function ParseDecimalToSQL(pValue)
		Dim value = pValue
		value = Replace(value, ".", "")
		value = Replace(value, "," , ".")
		ParseDecimalToSQL = value
	End Function
                
	Function ToSqlString(pValue)
		Dim value = pValue
		value = "'" & value & "'"
		ToSqlString = value
	End Function
	
	Function ToSqlBit(pValue)
		If pValue = true Or LCase(pValue) = "true" Then
			ToSqlBit = 1
		Else
			ToSqlBit = 0
		End If 
	End Function 
%>