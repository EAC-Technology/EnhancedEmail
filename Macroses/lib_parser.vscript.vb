'Version 08/10/2014
'Fix problem with DIM to pre-assign a String type with empty value

eTokenType=Dictionary("none",0,"end_of_formula",1,"operator_plus",2,"operator_minus",3,"operator_mul",4,"operator_div",5,"operator_percent",6,"open_parenthesis",7,"comma",8,"dot",9,"close_parenthesis",10,"operator_ne",11,"operator_gt",12,"operator_ge",13,"operator_eq",14,"operator_le",15,"operator_lt",16,"operator_and",17,"operator_or",18,"operator_not",19,"operator_concat",20,"value_identifier",21,"value_true",22,"value_false",23,"value_number",24,"value_string",25,"open_bracket",26,"close_bracket",27,"operator_affectation",28)
ePriority=Dictionary("none",0,"concat",1,"or",2,"and",3,"not",4,"equality",5,"plusminus",6,"muldiv",7,"percent",8,"unaryminus",9)
exceptionType=Dictionary("name","","code",0,"position",0,"description","")
 
Function revDate(inDate)
	if isDate(inDate) then
		jour	= mid(cstr(indate),1,2)
		mois  = mid(cstr(indate),4,2)
		annee   = mid(cstr(indate),7,4)
		revDate = annee & "." & mois & "." & jour
	Else
		revDate = ""
	End If
End Function

Function normDate(inDate)
	jour	= mid(cstr(indate),9,2)
	mois  = mid(cstr(indate),6,2)
	annee   = mid(cstr(indate),1,4)
	normDate = jour & "." & mois & "." & annee
End Function

Dim EvaluatorError

class vException

	dim exp

	Sub Class_Initialize()
		exp = exceptionType
	End Sub

	sub raise()
		throw scripterror
	End Sub

	Sub unspecifyError(msg,pos)
		exp("name")="unspecifyError"
		exp("code")=0
		exp("position")=pos
		exp("description")=msg
		logger "generalError : " & msg
        EvaluatorError = "generalError : " & msg
	End Sub

	Sub incompleteString(missingChar,pos)
		exp("name")="incompleteString"
		exp("code")=1
		exp("position")=pos
		exp("description")="Incomplete string, missing " & missingChar & "; String started"
		logger "incompleteString : " & exp("description")
        EvaluatorError = "incompleteString : " & exp("description")
	End Sub

	Sub unexpectedToken(msg,pos)
		exp("name")="unexpectedToken"
		exp("code")=2
		exp("position")=pos
		exp("description")=msg
		logger "unexpectedToken : " & msg
        EvaluatorError = "unexpectedToken : " & msg
	End Sub

	Sub wrongOperator(msg,pos)
		exp("name")="wrongOperator"
		exp("code")=3
		exp("position")=pos
		exp("description")=msg
		logger "wrongOperator : " & msg
        EvaluatorError = "wrongOperator : " & msg
	End sub

End Class

class variables

	dim varType
	dim value

	sub Class_Initialize()
		value="null"
		varType="system"
	End Sub

	sub setValue(varValue)
        select case typename(varValue)
          case "Integer","Double"
            value = Cdbl(varValue)
            varType = "numeric"
          
          case "String"
            value = varValue
            varType = "string"
          
          case "Date"
            value = varValue
            varType = "date"
          
          case "Boolean"
            value = varValue
            varType = "boolean"
            
          case else
            value="null"
            varType="system"
        end select

	End Sub

	function getValue
		getValue=value
	End Function

	Function isNull
		if varType="system" then
			if value="null" then
				isNull = True
			Else
				isNull = False
			End If
		Else
			isNull=False
		End If
	End Function

	function getType
		getType=varType
	End Function

	function convertTo(newType)
		Select Case lcase(newType)
			case "string"
				Try
					value=cstr(value)
					varType="string"
					convertTo=True
				Catch
					convertTo=False
				End Try

			case "numeric"
				Try
					value=cdbl(value)
					varType="numeric"
					convertTo=True
				Catch
					convertTo=False
				End Try

			case "date"
				Try
					value=cdate(value)
					varType="date"
					convertTo=True
				Catch
					convertTo=False
				end Try

			case "boolean"
				Try
					value=cbool(value)
					varType="boolean"
					convertTo=True
				Catch
					convertTo=False
				End Try

		End Select

	End Function

	function serialize
		serialize=dictionary("type",varType,"value",cstr(value))
	End function

	Sub unserialize(serialvalue)
		value=serialvalue("value")
		convertTo(serialvalue("type"))
	End sub

End Class

Class EvalContext2

	Dim variablesDict

	sub Class_Initialize
		variablesDict = Dictionary()
	end Sub

	sub vDim(variableName)        
		if typename(variablesDict(variableName))="Empty" Then
            set variablesDict(variableName) = New variables
			variablesDict(variableName).setValue("")          
        else
          if variablesDict(variableName).getType="system" then
            variablesDict(variableName).setValue("")
          end if
	    end if
	End Sub

	Function getVariable(variableName)

		if isObject(variablesDict(variableName)) Then
			set getVariable = variablesDict(variableName)
		Else
			set variablesDict(variableName) = New variables
			Set getVariable = variablesDict(variableName)
		End If

	End Function

	Sub setVariable(variableName,variableValue)
		if isObject(variablesDict(variableName)) Then
			variablesDict(variableName).setValue(variableValue)
		Else
			set variablesDict(variableName) = New variables
			variablesDict(variableName).setValue(variableValue)
		End If
	end Sub


	function loadContext(serializedContext)

		try
            if typename(serializedContext)="String" then
              objContext = asJson(serializedContext)
            else
              if typename(serializedContext)="Dictionary" then
                objContext = serializedContext
              else
                objContext = Dictionary
              end if
            end if
			for each var in objContext
				set variablesDict(var) = New variables
				variablesDict(var).unserialize(objContext(var))
			next
			loadContext = true
		catch
			loadContext = false
		end try
	end function

	function saveContext
		outputDict = dictionary
		for each var in variablesDict
			set myVar = variablesDict(var)
			outputDict(var)=myVar.serialize
		next
		saveContext = ToJSON(outputDict)
	end function

End Class

Class Evaluator

	Dim mParser
	Dim EvaluatorFunctions
	Dim EvaluatorContext
	Dim parserexception
    Dim formObject


	Sub Class_Initialize()
		Set mParser = new parser
		Set parserexception = new vException
		Set EvaluatorContext = new EvalContext
	End Sub

	Function Eval(str)
		Try
			set mParser.mContext = EvaluatorContext
            if typename(formObject)="xmldialogbuilder" then
              set mParser.formObject = formObject
            end if
			Eval=mParser.Eval(str)

		catch scripterror as err
			Eval="<ERROR> " & err.message
		End try
	end function

	Function EvalString(str)
		Try
			EvalString=mParser.EvalString(str)
		catch
			EvalString=""
			parserexception = mParser.parserexception
		End try
	End Function

End Class

Class tokenizer
	dim mString
	dim mLen
	dim mPos
	dim mCurChar
	dim startpos

	dim type
	dim value
	dim mParser
	dim parserexception
	dim localContext
    dim formObject

	Sub Class_Initialize()
		mString=""
		mLen=0
		mPos=1
		mCurChar=""
		startpos=0
		type=eTokenType
		value=""
		set parserexception = new vException

	End Sub

	sub newInstance(Parser,str)
		mString=str
		mLen=len(str)
		set mParser = Parser
'		set localContext = pContext
		NextChar
	End Sub

	function getVariableFromContext(varName)
		set tmpVar = localContext.getVariable(varName)
		if tmpVar.getType="system" then
			raiseError("The variable (" & varName & ") is not defined in the local context")
		Else
			getVariableFromContext = tmpVar.getValue
		End If

	End Function

	Sub NextChar()
         If mPos < mLen+1 Then
            mCurChar = mid(mString,mPos,1)
            If mCurChar = Chr(147) Or mCurChar = Chr(148) Then
               mCurChar = """"
            End If
            mPos = mPos + 1
         Else
            mCurChar = ""
         End If
     End Sub

     Function IsOp()
         IsOp = mCurChar = "+" Or mCurChar = "-" Or mCurChar = "%" Or mCurChar = "/" Or mCurChar = "(" Or mCurChar = ")" Or mCurChar = "."
     End Function

	 Sub NextToken()
         startpos = mPos
         value = ""
         type = eTokenType("none")
         Do
            Select Case mCurChar
               Case ""
                  type = eTokenType("end_of_formula")
               Case "0"
                  ParseNumber()
               Case "1"
                  ParseNumber()
               Case "2"
                  ParseNumber()
               Case "3"
                  ParseNumber()
               Case "4"
                  ParseNumber()
               Case "5"
                  ParseNumber()
               Case "6"
                  ParseNumber()
               Case "7"
                  ParseNumber()
               Case "8"
                  ParseNumber()
               Case "9"
                  ParseNumber()
               Case "-"
                  NextChar()
                  type = eTokenType("operator_minus")
               Case "+"
                  NextChar()
                  type = eTokenType("operator_plus")
               Case "*"
                  NextChar()
                  type = eTokenType("operator_mul")
               Case "/"
                  NextChar()
                  type = eTokenType("operator_div")
               Case "%"
                  NextChar()
                  type = eTokenType("operator_percent")
               Case "("
                  NextChar()
                  type = eTokenType("open_parenthesis")
               Case ")"
                  NextChar()
                  type = eTokenType("close_parenthesis")
               Case "<"
                  NextChar()
                  If mCurChar = "=" Then
                     NextChar()
                     type = eTokenType("operator_le")
                  ElseIf mCurChar = ">" Then
                     NextChar()
                     type = eTokenType("operator_ne")
                  Else
                     type = eTokenType("operator_lt")
                  End If
               Case ">"
                  NextChar()
                  If mCurChar = "=" Then
                     NextChar()
                     type = eTokenType("operator_ge")
                  Else
                     type = eTokenType("operator_gt")
                  End If
               Case ","
                  NextChar()
                  type = eTokenType("comma")
               Case "="
                  NextChar()
                  type = eTokenType("operator_eq")

           	Case ":"
            	  NextChar()
            	  If mCurChar = "=" Then
                     NextChar()
                     type = eTokenType("operator_affectation")
                  End If
               Case "."
                  NextChar()
                  type = eTokenType("dot")
               Case "'"
                  ParseString(True)
                  type = eTokenType("value_string")
               Case """"
                  ParseString(True)
                  type = eTokenType("value_string")
               Case "&"
                  NextChar()
                  type = eTokenType("operator_concat")
               Case "["
                  NextChar()
                  type = eTokenType("open_bracket")
               Case "]"
                  NextChar()
                  type = eTokenType("close_bracket")
               Case Else
               	if asc(mCurChar)>= 0 and asc(mCurChar)<= asc(" ") Then
               		' do Nothing
               	Else
               		ParseIdentifier()
               	End If

            End Select

            If type <> eTokenType("none") Then
            	Exit Do
            end If
            NextChar()
         Loop
      End Sub

	  Sub ParseNumber()

         type = eTokenType("value_number")
         Do While mCurChar >= "0" And mCurChar <= "9"
            value = value & mCurChar
            NextChar()
         Loop

         If mCurChar = "." Then
            value = value & mCurChar
            NextChar()
            Do While mCurChar >= "0" And mCurChar <= "9"
               value = value & mCurChar
               NextChar()
           Loop
         End If
      End Sub

	  Sub ParseIdentifier()
         Do While (mCurChar >= "0" And mCurChar <= "9") Or (mCurChar >= "a" And mCurChar <= "z") Or (mCurChar >= "A" And mCurChar <= "Z") Or (mCurChar >= "A" And mCurChar <= "Z") Or (mCurChar >= Chr(128)) Or (mCurChar = "_")
            value = value & mCurChar
            NextChar()
         Loop
         Select Case value
            Case "and"
               type = eTokenType("operator_and")
            Case "or"
               type = eTokenType("operator_or")
            Case "not"
               type = eTokenType("operator_not")
            Case "true", "yes"
               type = eTokenType("value_true")
            Case "yes"            
               type = eTokenType("value_true")
            Case "oui"            
               type = eTokenType("value_true")
            Case "false"
               type = eTokenType("value_false")
            Case "no"
               type = eTokenType("value_false")
            Case "non"
               type = eTokenType("value_false")
            Case Else
               type = eTokenType("value_identifier")
         End Select
      End Sub

	  Sub ParseString(InQuote)
         Dim OriginalChar
         Dim testOrgCur
         If InQuote Then
            OriginalChar = mCurChar
            NextChar()
         End If

         Do While mCurChar <> ""
         	if  mCurChar = OriginalChar Then
         		testOrgCur=True
         	Else
         		testOrgCur=False
         	end If

             If InQuote and testOrgCur Then
             	NextChar()
             	if  mCurChar = OriginalChar Then
         			testOrgCur=True
         		Else
         			testOrgCur=False
         		end If

             	If testOrgCur Then
                 	value = value & mCurChar
                 	NextChar()
                 Else
                  	'End of String
                  	Exit Sub
                End If
            ' ElseIf mCurChar = "%" Then
            '   NextChar()
            '   If mCurChar = "[" Then
            '      NextChar()
            '      Dim SaveValue
            '      SaveValue=value
            '      Dim SaveStartPos
            '      SaveStartPos = startpos
            '      me.value = ""
            '      NextToken() ' restart the tokenizer for the subExpr
            '      Dim subExpr
            '      Try
            '         subExpr = mParser.ParseExpr(0, ePriority("none"))
            '         If isNothing(subExpr) Then
            '            me.value = me.value & "<Nothing>"
            '         Else
            '            me.value = me.value & subExpr
            '         End If
            '      Catch
            '         me.value = me.value & "<error " & "some error !" & ">"
            '      End Try
            '      SaveValue = SaveValue & value
            '      value = SaveValue
            '      startpos = SaveStartPos
            '   Else
            '      value.Append("%")
            '   End If
            Else
               value = value & mCurChar
               NextChar()
            End If
         Loop
         If InQuote Then
         	parserexception.incompleteString(OriginalChar,startpos)
         	parserexception.raise
         End If
      End Sub

      sub raiseError(msg)
      	parserexception.unspecifyError(msg,startpos)
      	parserexception.raise
      End Sub

      sub RaiseUnexpectedToken(msg)
      	if len(msg)<>0 Then
      		msg = msg & ";"
      	End If
      	parserexception.unexpectedToken(msg & "Unexpected " & ToJSON(type) & " : " & value,startpos)
      	parserexception.raise

      End Sub

      sub RaiseWrongOperator(tt,ValueLeft,valueRight,msg)
      	If len(msg)<>0 then
      		msg = replace (msg,"[op]",ToJSON(tt))
      		msg = msg & "."
      	Else
      		 msg = "Cannot apply the operator " & ToJSON(tt)
      	End If
      	If IsNothing(ValueLeft) Then
            msg = msg & " on Nothing"
		  Else
            msg = msg & " on a " & typeName(ValueLeft)
          End If
          If Not IsNothing(valueRight) Then
            msg = msg & " and a " & typeName(ValueLeft)
          End If
      	parserexception.wrongOperator(msg,startpos)
      	parserexception.raise
      End Sub

End Class

Class parser

	Dim parserTokenizer
	Dim mEvalFunctions
	Dim parserexception
	Dim mContext
    Dim formObject

	Sub Class_Initialize()

	End Sub

'	sub newInstance(getContext)
'		set mContext=getContext
'	End Sub

	Function ParseExpr(Acc,priority)

		Dim ValueLeft, valueRight
		Do
			Select Case parserTokenizer.type

				Case eTokenType("operator_minus")
					'unary minus operator
					parserTokenizer.NextToken()
                    ValueLeft = ParseExpr(0, ePriority("unaryminus"))
                    If TypeName(ValueLeft)="Double" Then
                     	ValueLeft = CDbl(-1*ValueLeft)
                    Else
                     	parserTokenizer.RaiseWrongOperator(eTokenType("operator_minus"), ValueLeft, "Nothing", "You can use [op] only with numbers")
                    End If
                    Exit Do

				Case eTokenType("operator_plus")
					'unary plus operator
					parserTokenizer.NextToken()

				Case eTokenType("operator_not")
					parserTokenizer.NextToken()
					ValueLeft = ParseExpr(0, ePriority("not"))
					If TypeName(ValueLeft)="Boolean" Then
                     	ValueLeft = Not ValueLeft
                    Else
                     	parserTokenizer.RaiseWrongOperator(eTokenType("operator_not"), ValueLeft, "Nothing", "You can use [op] only with boolean values")
                  End If

				Case  eTokenType("value_identifier")
					identifier = parserTokenizer.value
					parserTokenizer.NextToken()
					if parserTokenizer.type = eTokenType("open_parenthesis") Then
						'We got a parenthesis, so it's function, run external function call
                        if typename(formObject)="xmldialogbuilder" then
                          set parserTokenizer.formObject = formObject
                        end if
						ValueLeft = InternalGetFunction(identifier)
					Else
						'The next token is not a parenthesis so, it's a variable
						ValueLeft = parserTokenizer.getVariableFromContext(identifier)
					end if
					Exit Do

				Case eTokenType("value_true")
					ValueLeft = True
					parserTokenizer.NextToken()
					Exit Do

				Case eTokenType("value_false")
					ValueLeft = False
					parserTokenizer.NextToken()
					Exit Do

				Case eTokenType("value_string")
					ValueLeft = cstr(parserTokenizer.value)
					parserTokenizer.NextToken()
					Exit Do
				Case eTokenType("value_number")
					ValueLeft = CDbl(parserTokenizer.value)

					parserTokenizer.NextToken()
					Exit Do

				Case eTokenType("open_parenthesis")
				    parserTokenizer.NextToken()
                    ValueLeft = ParseExpr(0, ePriority("none"))
                    If parserTokenizer.type = eTokenType("close_parenthesis") Then
                     	' good we eat the end parenthesis and continue ...
                     	parserTokenizer.NextToken()
                     	Exit Do
                    Else
                     	parserTokenizer.RaiseUnexpectedToken("End parenthesis not found")
                    End If

				Case Else
					Exit Do

			End Select
		Loop

		Do
			dim tt
			tt=parserTokenizer.type
			select Case tt
				Case eTokenType("end_of_formula")
					ParseExpr = ValueLeft
					Exit Do
				Case eTokenType("value_number")
					parserTokenizer.RaiseUnexpectedToken("Unexpected number without previous opterator")
				Case eTokenType("operator_plus")
					if priority < ePriority("plusminus") Then
						parserTokenizer.NextToken()
						valueRight = ParseExpr(ValueLeft, ePriority("plusminus"))
						If TypeName(ValueLeft)="Double" and TypeName(valueRight)="Double" then
							 ValueLeft = CDbl(ValueLeft) + CDbl(valueRight)
						ElseIf TypeName(ValueLeft)="Date" And TypeName(valueRight)="Double" then
							 ValueLeft = Dateadd("d",cint(valueRight),cdate(ValueLeft))
						ElseIf TypeName(ValueLeft)="Double" And TypeName(valueRight)="Date" then
							 ValueLeft = Dateadd("d",cint(ValueLeft),cdate(valueRight))
						Else
							 ValueLeft = ValueLeft & valueRight
						End If
					Else
						Exit Do
					End If

				Case eTokenType("operator_minus")
					if priority < ePriority("plusminus") Then
						parserTokenizer.NextToken()
						valueRight = ParseExpr(ValueLeft, ePriority("plusminus"))
						If TypeName(ValueLeft)="Double" and TypeName(valueRight)="Double" then
							ValueLeft = CDbl(ValueLeft) - CDbl(valueRight)
						ElseIf TypeName(ValueLeft)="Date" And TypeName(valueRight)="Double" then
							ValueLeft = Dateadd("d",cint(-1*valueRight),cdate(ValueLeft))
						ElseIf TypeName(ValueLeft)="Date" And TypeName(valueRight)="Date" then
							ValueLeft = datediff("d",ValueRight,ValueLeft)
						Else
							parserTokenizer.RaiseWrongOperator(tt,ValueLeft,valueRight,"You can use [op] only with numbers or dates")
						End If
					Else
						Exit Do
					End If

				Case eTokenType("operator_concat")
				  If priority < ePriority("concat") Then
                     parserTokenizer.NextToken()
                     valueRight = ParseExpr(ValueLeft, ePriority("concat"))
                     ValueLeft = cstr(ValueLeft) & cstr("valueRight")
                  Else
                     Exit Do
                  End If

				Case eTokenType("operator_mul")

					if priority < ePriority("muldiv") Then
						parserTokenizer.NextToken()
						valueRight = ParseExpr(ValueLeft, ePriority("muldiv"))
						if TypeName(ValueLeft)="Double" and TypeName(Valueright)="Double"  Then
							ValueLeft = CDbl(ValueLeft) * CDbl(valueRight)
						Else
							parserTokenizer.raiseError("Cannot apply the operator * with a " & TypeName(ValueLeft) & " and " & TypeName(valueRight) & " You can use - only with numbers")
							parserexception.raise
						End If
					Else
						Exit Do
					End If

				Case eTokenType("operator_div")
				    if priority < ePriority("muldiv") Then
						parserTokenizer.NextToken()
						valueRight = ParseExpr(ValueLeft, ePriority("muldiv"))
						if TypeName(ValueLeft)="Double" and TypeName(Valueright)="Double"  Then
							Try
								ValueLeft = CDbl(ValueLeft) / CDbl(valueRight)
							catch divisionbyzero
								parserTokenizer.raiseError("Math error you can not divide by 0")
								parserexception.raise
							end Try
						Else
							parserTokenizer.raiseError("Cannot apply the operator / with a " & TypeName(ValueLeft) & " and " & TypeName(valueRight) & " You can use - only with numbers")
							parserexception.raise
						End If
				    Else
				    	Exit Do
					End If

				Case eTokenType("operator_percent")
					if priority < ePriority("percent") Then
						parserTokenizer.NextToken()
						if typeName(ValueLeft)="Double" and TypeName(Acc)="Double" Then
							ValueLeft = CDbl(Acc) * CDbl(ValueLeft) / 100.0
						Else
							Dim ValueLeftString
							if isNothng(ValueLeft) Then
								ValueLeftString = "Nothing"
							Else
								ValueLeftString = typeName(ValueLeft)
							End If
							parserTokenizer.raiseError("Cannot apply the operator + or - on a " & ValueLeftString & " and " & TypeName(valueRight) & ",You can use % only with numbers. For example 150 + 20.5% ")
							parserexception.raise
						End If
					Else
						Exit Do
					End If

				Case eTokenType("operator_or")
					if priority < ePriority("or") Then
						parserTokenizer.NextToken()
						valueRight = ParseExpr(ValueLeft, ePriority("or"))
						if typename(ValueLeft) = "Boolean" and typename(valueRight) = "Boolean" then
							ValueLeft = CBool(ValueLeft) Or CBool(valueRight)
						Else
							parserTokenizer.raiseError("Cannot apply the operator OR on a " & TypeName(valueLeft) & " and " & TypeName(valueRight) & "You can use OR only with boolean values")
							parserexception.raise
						End If
					End If
				Case eTokenType("operator_and")
					if priority < ePriority("and") Then
						parserTokenizer.NextToken()
						valueRight = ParseExpr(ValueLeft, ePriority("or"))
						if typename(ValueLeft) = "Boolean" and typename(valueRight) = "Boolean" then
							ValueLeft = CBool(ValueLeft) And CBool(valueRight)
						Else
							parserTokenizer.raiseError("Cannot apply the operator AND on a " & TypeName(valueLeft) & " and " & TypeName(valueRight) & "You can use OR only with boolean values")
							parserexception.raise
						End If
					End If

				Case  eTokenType("operator_ne")
					If priority < ePriority("equality") Then
						tt=parserTokenizer.type
						parserTokenizer.NextToken()
						valueRight = ParseExpr(ValueLeft, ePriority("equality"))
						if ValueLeft <> valueRight Then
							ValueLeft = True
						Else
							ValueLeft = False
						End If

					Else
						parserTokenizer.RaiseWrongOperator(tt, ValueLeft, valueRight,"Wrong priority order in your operation with 'not'")
					End If

				Case eTokenType("operator_gt")
					If priority < ePriority("equality") Then
						tt=parserTokenizer.type
						parserTokenizer.NextToken()
						valueRight = ParseExpr(ValueLeft, ePriority("equality"))
						if ValueLeft > valueRight Then
							ValueLeft = True
						Else
							ValueLeft = False
						end if
					Else
						parserTokenizer.RaiseWrongOperator(tt, ValueLeft, valueRight,"Wrong priority order in your operation with '>'")
					End If

				Case eTokenType("operator_ge")
					If priority < ePriority("equality") Then
						tt=parserTokenizer.type
						parserTokenizer.NextToken()
						valueRight = ParseExpr(ValueLeft, ePriority("equality"))
						if ValueLeft >= valueRight Then
							ValueLeft = True
						Else
							ValueLeft = False
						End If
					Else
						parserTokenizer.RaiseWrongOperator(tt, ValueLeft, valueRight,"Wrong priority order in your operation with '>='")
					End If

				Case eTokenType("operator_eq")
					If priority < ePriority("equality") Then
						tt=parserTokenizer.type
						parserTokenizer.NextToken()
						valueRight = ParseExpr(ValueLeft, ePriority("equality"))
						if ValueLeft = valueRight Then
							ValueLeft = True
						Else
							ValueLeft = False
						end If
					Else
						parserTokenizer.RaiseWrongOperator(tt, ValueLeft, valueRight,"Wrong priority order in your operation with '='")
					End If

				Case eTokenType("operator_le")
					If priority < ePriority("equality") Then
						tt=parserTokenizer.type
						parserTokenizer.NextToken()
						valueRight = ParseExpr(ValueLeft, ePriority("equality"))
						if ValueLeft <= valueRight Then
							ValueLeft = True
						Else
							ValueLeft = False
						end If
					Else
						parserTokenizer.RaiseWrongOperator(tt, ValueLeft, valueRight,"Wrong priority order in your operation with '<='")
					End If

				Case eTokenType("operator_lt")
					If priority < ePriority("equality") Then
						tt=parserTokenizer.type
						parserTokenizer.NextToken()
						valueRight = ParseExpr(ValueLeft, ePriority("equality"))
						if ValueLeft < valueRight Then
							ValueLeft = True
						Else
							ValueLeft = False
						End If

					Else
						parserTokenizer.RaiseWrongOperator(tt, ValueLeft, valueRight,"Wrong priority order in your operation with '>='")
					End If

				Case Else
					Exit Do
			End Select
		Loop

		ParseExpr = ValueLeft

	End Function

	Function InternalGetFunction(functionName)

		Dim valueleft
		parameterList = Array(0)
		paramCount=0
		parserTokenizer.NextToken
		'Let's get the parameters of the function
		Do
			if parserTokenizer.type = eTokenType("close_parenthesis") Then
				'End of the function parmeters
				parserTokenizer.NextToken
				Exit Do
			End If
			valueleft = ParseExpr(0, ePriority("none"))
			if paramCount>0 Then
				ReDim Preserve parameterList(paramCount)
				'ReDim Preserve parameterList(paramCount+1)  ### seems redim work now
			End If
			parameterList(paramCount)=valueleft
			paramCount = paramCount + 1
			if parserTokenizer.type = eTokenType("close_parenthesis") Then
				parserTokenizer.NextToken
				Exit Do
			ElseIf eTokenType("comma") Then
				parserTokenizer.NextToken()
			Else
				parserTokenizer.RaiseUnexpectedToken("End parenthesis not found")
			End If
		Loop
		'Now check if we have this function in out lib
		'logger "here the function found : " & functionname & " " & ToJSON(parameterList)
		set myFuncLib = New functionsLib
		set myFuncLib.varContext = parserTokenizer.localContext
        
        if typename(parserTokenizer.formObject)="xmldialogbuilder" then
          set myFuncLib.formObject = parserTokenizer.formObject
        end if
        
		myFuncLib.eval(functionName,parameterList)

		if  myFuncLib.executionState Then
			InternalGetFunction = myFuncLib.result
		Else
			parserTokenizer.raiseError(myFuncLib.errMsg)
		End If


	End Function

	Function Eval(str)
		'logger "Parser - Eval - in : (" & str & ")"
		if Len(str) > 0 Then
			set parserTokenizer = New tokenizer
			parserTokenizer.newInstance(Me, str )
			set parserTokenizer.localContext = mContext
			parserTokenizer.NextToken
			Dim res
			res = ParseExpr(Nothing,ePriority("none"))
			'logger "Parser - Eval - out : (" & cstr(res) & ")"
			if parserTokenizer.type = eTokenType("end_of_formula") Then
				Eval=res
			Else
				parserTokenizer.RaiseUnexpectedToken("")
			End If
		End If
	End Function

	Function EvalString(str)
			if Len(str) > 0 Then
				set parserTokenizer = New tokenizer
				parserTokenizer.newInstance(Me, str, mContext)
				parserTokenizer.ParseString(False)
				EvalString = parserTokenizer.value
			Else
				EvalString = ""
			End If
	End Function
    
End Class