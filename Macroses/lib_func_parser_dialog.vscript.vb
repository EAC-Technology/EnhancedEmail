Class functionsLib

    dim executionState
    dim errMsg
    dim result
    dim DBConn
    dim varContext
    dim formObject

    Sub Class_Initialize()
    End Sub

    sub eval(funcName,parameters)

        '####################### Hack because of bug in redim preserve !!!!
        if len(parameters)<=1 Then
            nbparameters=len(parameters)
        Else
            nbparameters=len(parameters)-1
        End If
        executionState = True
        select case lcase(funcName)      
            '##################################################
            '#              VARIABLES        (1)              #
            '##################################################
            case "dim"
                'Create a variable in the local context
                If len(parameters) = 1 Then
                    if typename(parameters(0))="String" then
                        varContext.vDim(parameters(0))
                        result= ""
                    else
                        result= ""
                        executionState = False
                        errMsg = "The function '"+funcName+"' needs first the name of the variable and it must be a string."
                    end if
                elseif len(parameters) = 2 Then
                    if typename(parameters(0))="String" then                        
                        varContext.vDim(parameters(0))
                        if cstr(varContext.getVariable(parameters(0)).value)="" then
                            varContext.setVariable(parameters(0),parameters(1))
                        end if
                        result= varContext.getVariable(parameters(0)).value
                    else
                        result= ""
                        executionState = False
                        errMsg = "The function '"+funcName+"' needs first the name of the variable and it must be a string."
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcName+"' needs the name of the variable as parameter."
                end if
                
            case "affect"
                'affect a value to a variable of the context & return this value
                'WARNING : As parser eval all function in a single pass this function must be used as a surrunding one
                'it can not be used in the function if() as all parameters will be evaluated and it will not be then
                'a conditionnal affectation !!
                
                If len(parameters) = 2 Then
                    if typename(parameters(0))="String" then
                        result = parameters(1)
                        varContext.setVariable(parameters(0),parameters(1))
                    else
                        result= ""
                        executionState = False
                        errMsg = "The function '"+funcName+"' needs first the name of the variable and it must be a string."
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcName+"' needs 2 parameters, here we have " & cstr(len(parameters)) & " provided."
                end if
                
            case "chaine"
            'convert to string
                If len(parameters) = 1 Then
                    result=cstr(parameters(0))
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcName+"' needs 1 parameter, here we have " &cstr(len(parameters)) & " provided."
                end if  
                
            case "convertirvers"
            'Convert from original type to target type
                If len(parameters) = 2 Then
                    select case lcase(parameters(1))
                        case "chaine"
                            Try
                                result=cstr(parameters(0))
                            Catch
                                result= ""
                                executionState = False
                                errMsg = "The value given can not be converted to String."
                            End Try

                        case "numeric"
                            Try
                                result=cdbl(parameters(0))
                            Catch
                                result= ""
                                executionState = False
                                errMsg = "The value given can not be converted to Numeric."
                            End Try

                        case "date"
                            Try
                                result=cdate(parameters(0))
                            Catch
                                result= ""
                                executionState = False
                                errMsg = "The value given can not be converted to Date."
                            End Try
                    
                        case "boolean"
                            Try
                                result=cbool(parameters(0))
                            Catch
                                result= ""
                                executionState = False
                                errMsg = "The value given can not be converted to Boolean."
                            End Try
                  
                        case else
                            result= ""
                            executionState = False
                            errMsg = "The type '"&parameters(1)&"' is unknown."
                  
                    end select
                
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcName+"' needs 2 parameters, here we have " &cstr(len(parameters)) & " provided."
                end if
            
            case "dictionnaire"
            'Manage a global dictionary
            'arg : 0 - dicName, 1 - Key, 2 - Value, 3 - champliste, 4 - option (clef,valeur) (def clef)
                dimparam = len(parameters)
                removeop = false
                ValidKey = true
                oldValue = ""
                if dimparam = 2  then
                    if parameters(1)<>"" then
                        dicName = "_"+normalize(parameters(0))
                        if dbDictionary(dicName)="" then
                            result = ""
                        else
                            result = asjson(dbDictionary(dicName))(replace(parameters(1),chr(10),""))
                        end if
                    end if
                elseif dimparam >= 3  then
                    if parameters(1)<>"" then
                        dicName = "_"+normalize(parameters(0))
                        if dbDictionary(dicName)="" then
                            InitDic = true
                            if parameters(2)<>"" then
                                dbDictionary(dicName) = tojson(Dictionary(parameters(1),parameters(2)))
                            end if                   
                        else
                            if parameters(2)="" then
                                myDic = asjson(dbDictionary(dicName))
                                ValToRemove = myDic(parameters(1))
                                removeop = true
                                if myDic(parameters(1))<>"" then
                                    Erase myDic(parameters(1))
                                else
                                    ValidKey = false
                                end if
                                if len(myDic)=0 then
                                    Erase dbDictionary(dicName)
                                else
                                    dbDictionary(dicName) = tojson(myDic)
                                end if
                            else
                                myDic = asjson(dbDictionary(dicName))
                                if myDic(parameters(1))<>"" then
                                    oldValue = myDic(parameters(1))
                                end if
                                myDic(parameters(1))=parameters(2)
                                dbDictionary(dicName) = tojson(myDic)
                            end if
                        end if
                        if dimparam >= 4 then
                            if ValidKey then 
                                if parameters(3) then
                                    if InitDic then
                                        if parameters(2)<>"" then
                                            if dimparam = 5 then
                                                if lcase(parameters(4))="valeur" then 
                                                    dbDictionary(dicName+"_")=tojson(Array(parameters(2)))
                                                else
                                                    dbDictionary(dicName+"_")=tojson(Array(parameters(1)))
                                                end if
                                            else
                                                dbDictionary(dicName+"_")=tojson(Array(parameters(1)))
                                            end if
                                        else
                                            dbDictionary(dicName+"_")="[]"
                                        end if
                                    else
                                        if len(myDic)=0 then
                                            dbDictionary(dicName+"_")="[]"
                                        else
                                            if dbDictionary(dicName+"_")="" then
                                            'there is nothing we fullfill with data from dictionary
                                                listValArray=Array
                                                for each key in myDic
                                                    if dimparam = 5 then
                                                        if lcase(parameters(4))="valeur" then                                                    
                                                            if ExistInArray(listValArray,myDic(key)) then
                                                                val = myDic(key)+"_"+key
                                                            else
                                                                val = myDic(key)
                                                            end if
                                                        else
                                                            val = key
                                                        end if
                                                    else
                                                        val = key
                                                    end if
                                                    AppendToArray(listValArray,val)
                                                next
                                            else
                                                listValArray= asjson(dbDictionary(dicName+"_"))
                                                if simpleif(removeop,len(listValArray)-1<>len(myDic),len(listValArray)+1<>len(myDic)) then
                                                    listValArray=Array
                                                    for each key in myDic
                                                        if dimparam = 5 then
                                                            if lcase(parameters(4))="valeur" then                                                    
                                                                if ExistInArray(listValArray,myDic(key)) then
                                                                    val = myDic(key)+"_"+key
                                                                else
                                                                    val = myDic(key)
                                                                end if
                                                            else
                                                                val = key
                                                            end if
                                                        else
                                                            val = key
                                                        end if
                                                        AppendToArray(listValArray,val)
                                                    next
                                                else
                                                    if dimparam = 5 then
                                                        val = parameters(1)
                                                        if lcase(parameters(4))="valeur" then
                                                            if removeop then
                                                                if ExistInArray(listValArray,ValToRemove+"_"+parameters(1)) then
                                                                    RemoveFromArray(listValArray,ValToRemove+"_"+parameters(1))
                                                                else
                                                                    RemoveFromArray(listValArray,ValToRemove)
                                                                end if
                                                            else
                                                                val = parameters(2)
                                                                if oldValue<>"" then
                                                                    RemoveFromArray(listValArray,val+"_"+oldValue)
                                                                end if
                                                                if ExistInArray(listValArray,val) then
                                                                    val = val + "_" + parameters(1)
                                                                end if
                                                            end if
                                                        end if
                                                    else
                                                        val = parameters(1)
                                                        if removeop then
                                                            RemoveFromArray(listValArray,val)
                                                        end if
                                                    end if
                                                    if not removeop then                                            
                                                        AppendToArray(listValArray,val)
                                                    end if
                                                end if
                                            end if
                                            dbDictionary(dicName+"_")=tojson(listValArray)
                                        end if
                                    end if
                                end if
                            end if
                        end if
                    end if
                else 
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcName+"' needs 2,3,4,5 parameters, here we have " &cstr(len(parameters)) & " provided."
                end if
                
            
            '##################################################
            '#              HOURS & DATES FUNCTIONS   (2)     #
            '##################################################

            case "now"
                'Return the current system date with date type
                result = now

            case "today"
                'Return the current system date with date type
                result = now

            case "date"
                'Return the current date without time
                result = Date
            
            case "jour"
            'cut the date & return the Day
                If len(parameters) = 1 Then
                    if isdate(parameters(0)) then
                        result=day(parameters(0))
                    else
                        result= ""
                        executionState = False
                        errMsg = "The function 'jour' needs a date as parameter, here we have " &cstr(typename(parameters(0))) & " provided."
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function jour needs 1 parameters, here we have " &cstr(len(parameters)) & " provided."
                end if

            case "mois"
            'cut the date & return the month
                If len(parameters) = 1 Then
                    if isdate(parameters(0)) then
                        result=month(parameters(0))
                    else
                        result= ""
                        executionState = False
                        errMsg = "The function mois needs a date as parameter, here we have " &cstr(typename(parameters(0))) & " provided."
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function mois needs 1 parameters, here we have " &cstr(len(parameters)) & " provided."
                end if
                
            case "annee"
            'cut the date & return the year
                If len(parameters) = 1 Then
                    if isdate(parameters(0)) then
                        result=year(parameters(0))
                    else
                        result= ""
                        executionState = False
                        errMsg = "The function mois needs a date as parameter, here we have " &cstr(typename(parameters(0))) & " provided."
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function annee needs 1 parameters, here we have " &cstr(len(parameters)) & " provided."
                end if
              
            case "todate"
                'todate(yyyy,mm,dd) Convert 3 parameters to a valid date with the type date
                If len(parameters) = 3 Then
                    Try
                        result = cdate(cstr(parameters(2)) & "." & cstr(parameters(1)) & "." & cstr(parameters(0)))
                    Catch
                        result= ""
                        executionState = False
                        errMsg = "The values provided doesn't represent a valid date."
                    end Try
                Else
                    result= ""
                    executionState = False
                    errMsg = "The function todate needs 3 parameters, here we have " &cstr(len(parameters)) & " provided."
                End If
                 
            '##################################################
            '#                  CONDITIONS       (3)          #
            '##################################################
            case "cas","case"
               'manage different statement with different cases
            If len(parameters)>=3 Then
                result=""
                if len(parameters) mod 2 <> 0 then
                    for i=1 to len(parameters)-1 Step 2
                        if parameters(0)=parameters(i) then
                            result=parameters(i+1)
                            exit for
                        end if
                    next
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' must have an odd amount of parameters."
                end if
            else
                result= ""
                executionState = False
                errMsg = "The function '"+funcname+"' at least 3 parameters, here we have " & cstr(len(parameters)) & " provided."            
            end if

            case "if"
                'manage a if statement to return one or an other like this if(condition,if_TRUE,if_FALSE)
                If len(parameters) = 3 Then
                    If TypeName(parameters(0))="Boolean" Then
                        if parameters(0) Then
                            result = parameters(1)
                        Else
                            result = parameters(2)
                        End If
                    Else
                        result= ""
                        executionState = False
                        errMsg = "The test parameter must be a boolean."
                    End If
                Else
                    result= ""
                    executionState = False
                    errMsg = "The function 'if' needs 3 parameters, here we have " & cstr(len(parameters)) & " provided."
                End If
                          
                        
            '##################################################
            '#                 USERS & RIGHTS      (5)        #
            '##################################################
            
            case "estadmin","isadmin"
              'this function return true if we are admin
               result = (ProAdmin.current_user.email="root")
            
			case "estdansgroupe","isingroup"			
				if len(parameters)=1 then
					set user = ProAdmin.current_user
					groups = user.get_groups
					result = false
					for each group in groups
						if lcase(group.name) = lcase(parameters(0)) then
                          result = true
                          exit for
                        end if					
					next
				else
					result= ""
                    executionState = False
                    errMsg = "The function 'estdansgroupe' need 1 parameter, here we have " & cstr(len(parameters)) & " provided."
				end if
			
            
            case "trouverutilisateurs"
            'Convert login list into guid list
                if len(parameters)=1 then
                    loginList= Split(parameters(0),";")
                    result= ""
                    if len(loginList)<>0 then
                        for each login in loginlist
                            try
                                result = result + ProAdmin.search_users(login)(0).guid +";"
                            catch
                                logger ("TrouverUtilisateurs() user: " & login & " not found." )
                            end try
                        next
                        if len(result)<>0 then
                            result = mid(result,1,len(result)-1)
                        end if
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function 'TrouverUtilisateurs' need 1 parameters, here we have " & cstr(len(parameters)) & " provided."
                end if
                
            case "utilisateur"
            'Return usefull data about the current user
                if len(parameters)=2 then
                    UserList = ProAdmin.searchuser(Empty,parameters(0))
                    if len(UserList)=1 then
                        select case lcase(parameters(1))
                            case "prénom"
                                result = UserList(0).first_name
                            case "nom de famille"
                                result = UserList(0).last_name
                            case "nom"
                                result = UserList(0).name
                            case "email"
                                result = UserList(0).notification_email
                            case "téléphone"
                                result = UserList(0).cellphone
                            case else
                                result = ""
                                executionState = False
                                errMsg = "Invalid field request : '" & parameters(1) & "'"
                        end select
                    else
                        if len(UserList)=0 then
                            result= ""
                            executionState = False
                            errMsg = "Impossible to find user with GUID : '"&parameters(0)&"'"
                        else
                            result= ""
                            executionState = False
                            errMsg = "Invalid request more than one user was found."                  
                        end if
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function 'utilisateur' need 2 parameters, here we have " & cstr(len(parameters)) & " provided."
                end if
            
            case "utilisateurencours"
              result=ProAdmin.current_user.guid
            
            
            '##################################################
            '#              VARIOUS FUNCTIONS      (6)        #
            '##################################################
            
            case "abs"
                If len(parameters) = 1 Then
                    Try
                        result = abs (parameters(0))
                    Catch
                        result= ""
                        executionState = False
                        errMsg = err.message
                    End Try
                Elseif len(parameters) > 1 Then
                    result= ""
                    executionState = False
                    errMsg = "The function abs accept only one parameter, here " &cstr(len(parameters)) & " provided."
                else
                    result= ""
                    executionState = False
                    errMsg = "The function abs need exactly one parameter, 0 provided."
                End If
                    
            case "boutonvisible"
            'this function activate or not the visibility of the button in the form
            'the first parameter must be a boolean the second is optional & is the GUID of the button 
                if len(parameters)=2 or len(parameters)=1 then
                    if Typename(parameters(0))="Boolean" then
                        fodlerPath = varContext.getVariable("fodlerPath").value
                        if DBDictionary(fodlerPath)<>"" then
                            rulesData = asjson(DBDictionary(fodlerPath))
                            rules = rulesData("rules")
                            if len(parameters)=2 then
                                values = rules(parameters(1))
                                if typename(values)="Dictionary" then
                                    values("visible") = SimpleIf(parameters(0),"yes","no")
                                    sessiondictionary("btn_" & parameters(1)) =  values("visible")
                                end if
                            else
                                btnGUID = varContext.getVariable("btnGUID").value
                                values = rules(btnGUID)
                                values("visible") = SimpleIf(parameters(0),"yes","no")
                                sessiondictionary("btn_" & btnGUID) =  values("visible")
                            end if
                            DBDictionary(fodlerPath) = tojson(rulesData)
                            result = SimpleIf(parameters(0),"yes","no")
                        else
                            result= ""
                            executionState = False
                            errMsg = "Rule data for folder path : " & varContext.getVariable("fodlerPath").value & " are absents."
                        end if
                    else
                        result= ""
                        executionState = False
                        errMsg = "The first parameter of the function must be a boolean."
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 1 or 2 parameters, here we have " & cstr(len(parameters)) & " provided."
                end if
            
            case "chemin"
            'guid of path
                if len(parameters)=1 then
                    Try
                        result=ProShare.get_by_guid(parameters(0)).path
                    catch
                        result= ""
                        executionState = False
                        errMsg = "Impossible to get path from this GUID : " & parameters(0) & "."
                    end try
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 1 parameter, here we have " & cstr(len(parameters)) & " provided."
                end if
            
            case "couper"
            'Allow to remove spaces in a string
				If len(parameters) = 2 Then
					if len(parameters(0))>parameters(1) then
						result = mid(cstr(parameters(0)),1,parameters(1))+"..."
					else
						result = cstr(parameters(0))
					end if
               else
                  result= ""
                  executionState = False
                  errMsg = "The function '"+funcname+"' need 2 parameter, here we have " & cstr(len(parameters)) & " provided."
               end if

            
            case "csv"
                result = csv(parameters,executionState,errMsg,varContext)
            
            case "estnumeric","isnumeric"
                'Check if the value is kind of number ""
                If len(parameters) = 1 Then
                    try
                        testVal=CDbl(parameters(0))
                        if Typename(testVal)="Double" then
                            result=true
                        else
                            result=false
                        end if
                    catch
                        result=false
                    end try
                Else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 1 parameter, here we have " & cstr(len(parameters)) & " provided."
                end If
                
            case "apianswer","reponseapi"
            'return data to the caller
            '0- data
                If len(parameters) = 1 Then
                    response.write(cstr(parameters(0)))
                    result= cstr(parameters(0))
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 1 parameter, here we have " & cstr(len(parameters)) & " provided."                
                end if
            
            case "ecriredansfichier"
			'write data in a file
			'0- path, 1- filename, 2- data
				If len(parameters) = 3 Then
					targetPath = parameters(0)
                    set objPath = ProShare.get_by_guid(varContext.getVariable("currentPathGUID").value)
                    if mid(targetPath,1,2)="./" then
                        targetPath = replace (targetPath,"./",objPath.path + "/")  
                    elseif mid(targetPath,1,3)="../" then
                        targetPath = replace (targetPath,"../",objPath.parent.path + "/")
                    elseif mid(targetPath,1,4)=".../" then
                        targetPath = replace (targetPath,".../",objPath.parent.parent.path + "/")
                    end if
                    if mid(targetPath,len(targetPath))="/" then
                        targetPath = mid(targetPath,1,len(targetPath)-1)
                    end if
                    result=targetPath
					
					'Build file
                    set folder = ProShare.get_by_path(targetPath)
                    set buf = buffer.create
					set r = new RegExp
						r.IgnoreCase = true
						r.Global = true
						r.Pattern = "\{=?([^\{^\}]*)\}"
					Template = parameters(2)
					
					set Matches = r.execute (Template)
					set ExpToEval = New Evaluator
					set ExpToEval.EvaluatorContext = varContext
                    
					varError = false
					ErrData = Dictionary
					ErrCpt = 0
					for each match in matches
						ExpResult = ExpToEval.eval(match.value(1))
                        posErr = instr(lcase(ExpResult),"error") 
						if posErr<>0 and posErr<3  then
							varError = true
							ErrCpt = ErrCpt + 1
							ErrData(cstr(ErrCpt))=Array(match.value(1),ExpResult)
                            ExpResult = "?" & match.value(1) & "?"
						end if
						Template = replace(Template,match.value,SimpleIf(mid(match.value,1,2)="{=",ExpResult,""))
					next
					if not varError then
						buf.write(Template)
						result = true
						buf.seek(0)                             
                        FileName =  parameters(1)                                                  
                        while FileExists(targetPath & "/" & FileName) 
							FileName = "_" & FileName
                        wend                            
                        'Write file
                        folder.create_file( FileName, buf )
					else
						result= false
						executionState = False
						errMsg = "Error in function '"+funcname+"' >>> " & tojson(ErrData)
					end if
				else
					result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 3 parameters, here we have " & cstr(len(parameters)) & " provided."
				end if
				
			case "ecriredanshistorique"
            'ecriredanshistorique("msg")
                If len(parameters) = 1 Then
                    WriteToHistory(varContext.getVariable("fileGUID").value,"comment",parameters(0))
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 1 parameter, here we have " & cstr(len(parameters)) & " provided."
                end if
                
            case "goto"
				'make it more clean to use goto
				If len(parameters) = 1 Then
					result= "goto "+parameters(0)
				else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 1 parameter, here we have " & cstr(len(parameters)) & " provided."
				end if

            case "nonvide","isnotempty"
                'Check if the value is different ""
                If len(parameters) = 1 Then
                    if cstr(parameters(0))="" then
                        result=false
                    else
                        result=true
                    end if
                Else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 1 parameter, here we have " & cstr(len(parameters)) & " provided."
                end If
            
            case "partager"
            'this function allow to split data with a separator
            'Arg 0 - exp; Arg 1 - separator; Arg 2 - item
            if len(parameters)=3 then
              try
                arrayResults = split(parameters(0),parameters(1))
                result=arrayResults(cint(parameters(2)))
              catch
                result=""
                executionState = False
                errMsg = "An error occur in the function 'partager'" 
              end try
            else
              result= ""
              executionState = False
              errMsg = "The function '"+funcname+"' need 3 parameters, here we have " & cstr(len(parameters)) & " provided."
            end if
            
            case "remplace","remplacer"
                'Allow to replace something by something else
                If len(parameters) = 3 Then
                    result = replace (cstr(parameters(0)),parameters(1),parameters(2))              
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 3 parameter, here we have " & cstr(len(parameters)) & " provided."
              end if
            case "supprimerespace"
            'Allow to remove spaces in a string
              If len(parameters) = 1 Then
                result = replace (cstr(parameters(0))," ","")              
              else
                  result= ""
                  executionState = False
                  errMsg = "The function '"+funcname+"' need 1 parameter, here we have " & cstr(len(parameters)) & " provided."
              end if
                          
            case "sendmail"
            'Send a mail will send a mail as system account to a user.
            'The mail can be either mails, either users. The content of the mail support {exp}
                If len(parameters) = 3 Then                
                    ResultData = Dictionary
                    ResultData("nbEmail")=0
                    ResultData("EmailList")=Array
                    ResultData("nbUser")=0
                    ResultData("UserList")=Array
                    MailTemplate = parameters(2)
                    'Replace {Expression} by its value in the template
                    set r = new RegExp
                        r.IgnoreCase = true
                        r.Global = true
                        r.Pattern = "\{([^\{^\}]*)\}"
                    set Matches = r.execute (MailTemplate)
                    
                    set ExpToEval = New Evaluator
                    set ExpToEval.EvaluatorContext = varContext

                    varError = false
                    for each match in matches
                        ExpResult = ExpToEval.eval(match.value(1))
                        if instr(lcase(ExpResult),"error")<>0 then
                            varError = true
                            ExpResult = "?" & match.value(1) & "?"
                            logger ("SendMail Error in evaluating : " &  match.value(1))
                        end if
                        MailTemplate = replace(MailTemplate,match.value,ExpResult)
                    next
                    
                    if not varError then
                        MailTemplate = replace(MailTemplate,chr(10),"<br>")
                        'MailTemplate = replace(MailTemplate," ","&nbsp;")
                        'Send Email to true mail
                        Set regex=new regexp
                            regex.pattern="\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,6}\b"
                            regex.ignorecase=true
                            regex.global=true
                        set matches=regex.execute(parameters(0))
                        for each match in matches
                            ResultData("nbEmail") = ResultData("nbEmail") + 1
                            server.sendmail("Workflow",match.value,parameters(1),MailTemplate)
                            AppendToArray(ResultData("EmailList"),match.value)
                        next
                    
                        'Send mail to users
                        Set regex=new regexp
                            regex.pattern="(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}"
                            regex.ignorecase=true
                            regex.global=true
                        set matches=regex.execute(parameters(0))
                    
                        for each match in matches
                            ResultData("nbUser") = ResultData("nbUser") + 1
                            try
                                server.sendmail("Workflow",ProAdmin.searchuser(Empty,match.value)(0).notification_email,parameters(1),MailTemplate)
                                AppendToArray(ResultData("UserList"),ProAdmin.searchuser(Empty,match.value)(0).name)
                            catch
                                AppendToArray(ResultData("UserList"),"??>" & match.value)
                            end try                      
                        next
                        result = tojson(ResultData)
                    else
                        result= ""
                        executionState = False
                        errMsg = "Evaluation of expression in the template have generated some error, it's impossible to send the mail."
                    end if
                Else
                    result= ""
                    executionState = False
                    errMsg = "The function 'sendmail' needs 3 parameters, here we have " & cstr(len(parameters)) & " provided."
                end If

            case "windows"
            'windows ("titre","type","message")  
                If len(parameters) = 3 Then              
                    formObject.addScreen("windows")
                    formObject.Screen("windows").addComponent("comment","Text")
                    formObject.Screen("windows").Component("comment").setCenter(true)
                    formObject.Screen("windows").Component("comment").value("<br>" & parameters(2))
                    formObject.Screen("windows").Title = parameters(0)
                    select case lcase(parameters(1))
                        case "timer"
                            if typename(sessiondictionary("nodelist"))="Array" then
                                if len(sessiondictionary("nodelist"))=0 then
                                    sessiondictionary("filemode")="single"
                                    sessiondictionary("nodelist") = ""
                                    formObject.Screen("windows").addComponent("timeout","Timer")
                                    formObject.Screen("windows").Component("timeout").setTimer("EndGroupFileProcessing",3000)
                                else
                                    formObject.Screen("windows").addComponent("timeout","Timer")
                                    formObject.Screen("windows").Component("timeout").setTimer("Init",1000)
                                end if
                            else
                                formObject.Screen("windows").addComponent("timeout","Timer")
                                formObject.Screen("windows").Component("timeout").setTimer("Exit",3000)
                            end if
                            result= "[QUIT]"
                        
                        case "abandon","abandonne"
                            sessiondictionary("filemode")="single"
                            sessiondictionary("nodelist") = ""
                            formObject.Screen("windows").addComponent("timeout","Timer")
                            formObject.Screen("windows").Component("timeout").setTimer("Exit",1000)
                            result= "[QUIT]"
                        
                        case "reload"
                            formObject.Screen("windows").addComponent("timeout","Timer")
                            formObject.Screen("windows").Component("timeout").setTimer("Init",1000)
                            result= "[RELOAD > fileGUID:'"+SessionDictionary("fileguid")+"', filemode:'"+sessiondictionary("filemode")+"']"
                        case "file"
                            formObject.Screen("windows").Component("comment").value("<br>Chargement du fichier")
                            'Adding some additionnal textbox to store MacroID value
                            formObject.Screen("windows").addComponent("fileGUID","TextBox")
                            formObject.Screen("windows").component("fileGUID").visible(false)
                            formObject.Screen("windows").component("fileGUID").setProp("defaultValue",parameters(2))
                            formObject.Screen("windows").addComponent("timeout","Timer")
                            formObject.Screen("windows").Component("timeout").setTimer("Init",100)

                    end select
                    formObject.ShowScreen("windows")
                else
                    result= ""
                    executionState = False
                    errMsg = "The function 'Windows' need 3 parameters, here we have " & cstr(len(parameters)) & " provided."
                end if
              
            '##################################################
            '#             PROSUITE FUNCTIONS      (7)        #
            '##################################################
            
            case "creeevtproplanning"
                'date date, calendrier chaine, description chaine,emails_Invitations chaine
                if len(parameters)=3 or len(parameters)=4 then
                    cnxProPlanning = True
                    try
                        cnxState = ApplicationConnection( "ProPlanning", "127.0.0.1", "prosuite", "prosuite" )
                        set ProPlanning = cnxState(1)
                        users = ProAdmin.users("root")
                        set ROOT_USER = users(0)                    
                        LoginData = Dictionary("token", ProAdmin.accessTokenForUser(ROOT_USER))
                        login = asjson(ProPlanning.Invoke("login", ToJson(LoginData)))                 
                        logger ("Connecting to ProPlanning...")
                        if login(0)="error" then
                            cnxProPlanning = false
                            result= ""
                            executionState = False
                            errMsg = "Impossible to log into ProPlanning. Because : " & login(1)
                        end if
                        if cnxProPlanning then
                            dateok=true
                            try
                                fuck = cdate(Parameters(0))
                            catch
                                dateok=false
                            end try
                            if dateok then
                                Query = New Dictionary
                                Query("access")="o"
                                if DBDictionary("_calandar_Guid_" & normalize(Parameters(1)))<>"" then
                                    calandarGUID = DBDictionary("_calandar_Guid_" & normalize(Parameters(1)))
                                else
                                    Query = New Dictionary
                                    Query("access")="o"
                                    calendars  = asjson(ProPlanning.Invoke("retrieve_calendars", ToJson(Query)))
                                    if calendars(0)="success" then
                                        for each cal in calendars(1)
                                            if cal("name")=Parameters(1) then
                                                calandarGUID=cal("guid")
                                                DBDictionary("_calandar_Guid_" & normalize(Parameters(1)))=calandarGUID
                                                exit for
                                            end if
                                        next
                                        if calandarGUID="" then
                                            CalendarData = New Dictionary
                                            CalendarData("name")=Parameters(1)
                                            CalendarData("color")="#FFA733"
                                            newCalendar = asjson(ProPlanning.Invoke("create_calendar", ToJson(CalendarData)))
                                            if newCalendar(0)="success" then
                                                calandarGUID=newCalendar(1)
                                                DBDictionary("_calandar_Guid_" & normalize(Parameters(1)))=calandarGUID
                                            else
                                                cnxProPlanning = false
                                                result= ""
                                                executionState = False
                                                errMsg = "New Calandar (" & Parameters(1) & ") creation fail."
                                            end if
                                        end if
                                    else
                                        cnxProPlanning = false
                                        result= ""
                                        executionState = False
                                        errMsg = "Calandar retriving call fail."
                                    end if
                                end if                       
                            else
                                cnxProPlanning = false
                                result= ""
                                executionState = False
                                errMsg ="The first parameter is not a date."
                            end if
                        else
                            cnxProPlanning = false
                            result= ""
                            executionState = False
                            errMsg ="Impossible to log in ProPlanning."
                        end if
                    catch
                        cnxProPlanning = false
                        result= ""
                        executionState = False
                        errMsg = "Connexion to ProPlanning application is impossible."
                    end try  
                    if cnxProPlanning then
                        EventData = New Dictionary
                        EventData("start_date")=cstr(dateDiff("s",cdate("01/01/1970"),cdate(Parameters(0))))
                        EventData("end_date")=cstr(dateDiff("s",cdate("01/01/1970"),cdate(Parameters(0))))
                        EventData("summary")=Parameters(2)
                        EventData("description")="Evénement généré par le Plug-IN Automation"
                        EventData("all_day")=true
                        EventData("rrule")=""
                        EventData("calendar_guid")=calandarGUID
                    
                        if len(parameters)=4 then
                            Invite = New Dictionary
                            Invite("to_email")="some@email.com"
                            EventData("invites")=Array(Invite)
                        end if  
                        EventGuid = asjson(ProPlanning.Invoke("create_event", ToJson(EventData)))
                        if EventGuid(0)="success" then
                            result = EventGuid(1)
                        else
                            result= ""
                            executionState = False
                            errMsg = "Event creation failed : " & EventGuid(1) & "."
                        end if
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function 'CreeEvtProPlanning' need 3 or 4 parameters, here we have " & cstr(len(parameters)) & " provided."
                end if
                
            case "ajoutcontact","ajoutercontact"
                result = addContact(parameters,executionState,errMsg,varContext)
                
            case "contacts"
              'This function allow to retrieve contacts from a given contact list
              '(0) - contact_list_name, (1) field to get nom/email/telphone
              
              cnxProProContact = true
              if len(parameters)=2 then
                  cnxState = ApplicationConnection( getAppName("ProContact"), "127.0.0.1", "prosuite", "prosuite" )
                  set ProContact = cnxState(1)
                  users = ProAdmin.users("root")
                  set ROOT_USER = users(0)                    
                  LoginData = Dictionary("token", ProAdmin.accessTokenForUser(ROOT_USER))
                  login = asjson(ProContact.Invoke("login", ToJson(LoginData)))                 
                  'logger ("Connecting to ProContact...")
                  if login(0)="error" then
                      cnxProProContact = false
                      result= ""
                      executionState = False
                      errMsg = "Impossible to log into ProContact. Because : " & login(1)
                  end if
                  foundContactList = false
                  if DBDictionary("_contact_list_Guid_" & normalize(Parameters(0)))<>"" then
                      contactListGUID = DBDictionary("_contact_list_Guid_" & normalize(Parameters(0)))
                      foundContactList = true
                  else
                      Query = New Dictionary
                      Query("access")="r"
                      retrieve_contactlists = asjson(ProContact.Invoke( "retrieve_contactlists", ToJson(Query)))
                      if retrieve_contactlists(0)="success" then
                          for each contactlist in retrieve_contactlists(1)
                              if lcase(Parameters(0))=lcase(contactlist("name")) then
                                  contactListGUID = contactlist("guid")
                                  DBDictionary("_contact_list_Guid_" & normalize(Parameters(0))) = contactListGUID
                                  foundContactList = true
                                   exit for
                              end if
                          next
                      else
                          'No conctact list
                          result= ""
                          executionState = False
                          errMsg = "No contact list for this account."
                      end if
                  end if
                  if foundContactList and cnxProProContact then
                      'logger("Contact list found :" & contactListGUID)
                      Query = New Dictionary
                      Query("contactlists_guids")=Array(contactListGUID)
                      retrieve_contacts = asjson(ProContact.Invoke( "retrieve_contacts", ToJson(Query)))
                      contacts = ""
                      if retrieve_contacts(0)="success" then
                            for each contact in retrieve_contacts(1)
                                select case lcase(Parameters(1))
                                    case "nom","name"
                                        contacts = contacts + contact("first_name") + " " + contact("last_name") + ";"
                                    case "email"
                                        contacts = contacts + contact("email")(0)("value") + ";"
                                    case "telephone","phone"
                                        contacts = contacts + trim(replace(contact("phone")(0)("value")," ","")) + ";"
                                    case else
                                        result= ""
                                        executionState = False
                                        errMsg = "The field '"+Parameters(1)+"' doesn't exist."
                                        exit for
                                end select
                            next
                            if len(contacts)<>0 then
                                  result = mid(contacts,1,len(contacts)-1)
                            end if
                      else
                           'No conctact list
                           result= ""
                           executionState = False
                           errMsg = "Cache for Contact list '"+Parameters(0)+"' wrong, try again."
                           erase DBDictionary("_contact_list_Guid_" & normalize(Parameters(0)))
                      end if  
                  else
                      'No conctact list
                       result= ""
                       executionState = False
                       errMsg = "Contact list'"+Parameters(0)+"' list not found."
                  end if
              else
                result= ""
                executionState = False
                errMsg = formatstring("The function {0} need at least {1} parameter{2}, here we have {3} provided.",funcname,"2","",cstr(len(parameters)))                        
              end if
            
            case "pourrecherche"
                'Cette fonction permet de transmettre des metas données pour la recherche
                'un premier filtre se fait sur un type de base qui est le premier paramétre
                if len(parameters)>=1 then
                    set file = ProShare.get_by_guid(varContext.getVariable("fileGUID").value)
                    set current_user = ProAdmin.current_user
                    
                    '--- Main Index data
                    IndexData = Dictionary
                    if varContext.getVariable("_prosearch_id").isNull then
                        IndexData( "id" ) = "wf-" & file.guid
                        varContext.setVariable("_prosearch_id",IndexData( "id" ))
                    else
                        IndexData( "id" ) = varContext.getVariable("_prosearch_id").value
                    end if
                    IndexData( "name" ) = file.name
                    '--- Main meta data about the file
                    metaData = Dictionary
                    metaData( "path" )              = file.path
                    metaData( "modification_date" ) = file.modificationDate
                    metaData( "author" )            = current_user.name
                    metaData( "mime_type" )         = file.mimeType
                    metaData( "url" )               = FormatString( "/home.vdom#{0}", file.guid )
                    metaData( "size" )              = file.size
                    metaData( "index" )             = file.serializeString
                    metaData( "keyword" )           = parameters(0)
                    IndexData( "meta_data" ) = metaData
                    
                    'Set rights according to the right the file have
                    aclDict = Dictionary
                    aclArray = Array
                    fileRules = file.rules
                    for each fr in fileRules
                        aclDict( fr.subject.guid ) = fr.subject.guid
                    next
                    for each r in aclDict
                        AppendToArray(aclArray, r)
                    next
                    if len(aclArray) > 0 then
                        metaData( "acl" ) = aclArray
                    else
                        metaData( "acl" ) = Array("None")
                    end if            
                    IndexList = Dictionary( "path","string","modification_date","date","mime_type","string","url","string","size","float","index","string","author","string","acl","string","keyword","string")
                    metaDataList=Array
                    if len(parameters)>1 then
                        for i=1 to len(parameters)-1 step 2
                            if IndexList(normalize(parameters(i)))="" then
                                AppendToArray(metaDataList,parameters(i))
                                IndexList(normalize(parameters(i)))=lcase(typename(parameters(i+1)))
                            end if
                            metaData(normalize(parameters(i)))=parameters(i+1)
                        next
                    end if
                    metaData("metadatalist")=tojson(metaDataList)
                  
                    set cnxProSearch = new ProSearchApp
                        cnxProSearch.agentName = "WF Documents"
                        cnxProSearch.sourceName = "IndexedMetaData"
                        cnxProSearch.sourceindexes = IndexList
                        
                    if cnxProSearch.connect then
                        IndexData( "source_id" ) = cnxProSearch.GetSourceId
                        if cnxProSearch.send(IndexData)  then
                            result = IndexData( "id" )
                        else
                            result= ""
                            executionState = False
                            errMsg = "The function 'pourrecherche' returned an error while sending data to ProSearch : " & cnxProSearch.ErrMsg                                    
                        end if
                    else
                        result= ""
                        executionState = False
                        errMsg = "The function 'pourrecherche' returned an error while connecting to ProSearch : " & cnxProSearch.ErrMsg                                    
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function 'pourrecherche' need at least 1 parameter, here we have " & cstr(len(parameters)) & " provided."
                end if
            
            case "tache"
            'Cette fonction permet de créer une tâche pour le suivi des WF dans ProSearch
            'le premier paramétre définie la fonction à utiliser
            'Les données sont >
            '    - WF_priority     : 1 (Faible), 2 (Normal), 3 (Important),4 (Urgent)
            '    - WF_name         : Nom du WF
            '    - WF_groups       : id groupes
            '    - WF_step_name    : nom de l'étape
            '    - WF_limit        : date limite
            '    - WF_step         : numéro de l'étape
            '    - WF_step_max     : nombre max d'étape
            '    - WF_file_name    : nom du ficher
            '    - WF_file_path    : chemin absolue du fichier
            '    - WF_comments     : commentaires.
            '    - WF_file_content : contenu du fichier.
                if Typename(parameters(0))="String" then
                    set file = ProShare.get_by_guid(varContext.getVariable("fileGUID").value)
                    set current_user = ProAdmin.current_user
                    select case parameters(0)
                        case "new" 
                        'General Data about the file
                        'Parmeters given to the function
                        'Tache ("new" ,wf_priority,wf_name,wf_groups,wf_step_name,wf_limit,wf_step,wf_step_max,wf_comments)
                    
                            IndexData = Dictionary
                            IndexData( "id" ) = CreateGUID
                            IndexData( "name" ) = file.name
                            'Main meta data about the file
                            metaData = Dictionary
                            metaData( "path" )              = file.path
                            metaData( "modification_date" ) = file.modificationDate
                            metaData( "author" )            = current_user.name
                            metaData( "mime_type" )         = file.mimeType
                            metaData( "url" )               = FormatString( "/home.vdom#{0}", file.guid )
                            metaData( "size" )              = file.size
                            metaData( "index" )             = file.serializeString
                            metaData( "Task_Id" )           = IndexData("id")
                            IndexData( "meta_data" ) = metaData
                            'Set rights according to the right the file have
                            aclDict = Dictionary
                            aclArray = Array
                            fileRules = file.rules
                            for each fr in fileRules
                                aclDict( fr.subject.guid ) = fr.subject.guid
                            next
                            for each r in aclDict
                                AppendToArray( aclArray, r )
                            next
                            if len(aclArray) > 0 then
                                metaData( "acl" ) = aclArray
                            else
                                metaData( "acl" ) = Array("None")
                            end if
                            'Set specific data about the WF
                            if typename(parameters(1))="String" then
                                select case parameters(1)
                                    case "URGENT"
                                        priority=4
                                    case "HAUT"
                                        priority=3
                                    case "MOYEN"
                                        priority=2
                                    case "BAS"
                                        priority=1
                                    case else
                                        priority=2
                                end select
                            else
                                priority = parameters(1)
                            end if
                    
                            metaData( "wf_priority" ) = priority '1-Faible, 2-Normal, 3- Important, 4-Urgent
                            metaData( "wf_name" )     = parameters(2) 'WF name 
                            metaData( "wf_groups" )   = parameters(3) 'name of the group of pepole who are 
                            metaData( "wf_step_name" )= parameters(4) 'name of the step
                            try
                                limitDate = cdate(parameters(5))
                            catch
                                limitDate = dateadd("d",1,now)
                            end try
                            metaData( "wf_limit" )    = DateToUnixInt(limitDate) 'Limit date for the WF
                            metaData( "wf_step" )     = parameters(6) 'current step
                            metaData( "wf_step_max" ) = parameters(7) 'max amount of step
                            metaData( "wf_file_name" )= file.name 
                            metaData( "wf_file_path" )= file.parent.path
                            metaData( "wf_comments" ) = parameters(8)
                    
                            IndexList = Dictionary( "path","string","modification_date","date","mime_type","string","url","string","size","float","index","string","author","string","acl","string","wf_priority","float","wf_name","string","wf_groups","string","wf_step_name","string","wf_limit","float","wf_step_max","float","wf_file_name","string","wf_file_path","string","wf_comments","string","Tasks_Id","string")

                            set cnxProSearch = new ProSearchApp
                                cnxProSearch.agentName = "Automation"
                                cnxProSearch.sourceName = "Tasks"
                                cnxProSearch.sourceindexes = IndexList
                        
                            if cnxProSearch.connect then
                                IndexData( "source_id" ) = cnxProSearch.GetSourceId
                                if cnxProSearch.send(IndexData)  then
                                    result = IndexData( "id" )
                                else
                                    result= ""
                                    executionState = False
                                    errMsg = "The function 'tache' returned an error while sending data to ProSearch : " & cnxProSearch.ErrMsg                                    
                                end if
                            else
                                result= ""
                                executionState = False
                                errMsg = "The function 'tache' returned an error while connecting to ProSearch : " & cnxProSearch.ErrMsg                                    
                            end if
                        case "update"
                            IndexData = Dictionary
                            IndexData( "id" ) = parameters(1)
                            IndexData( "name" ) = file.name
                    
                            'Main meta data about the file
                            metaData = Dictionary
                            metaData( "path" )              = file.path
                            metaData( "modification_date" ) = file.modificationDate
                            metaData( "author" )            = current_user.name
                            metaData( "mime_type" )         = file.mimeType
                            metaData( "url" )               = FormatString( "/home.vdom#{0}", file.guid )
                            metaData( "size" )              = file.size
                            metaData( "index" )             = file.serializeString
                            metaData( "Task_Id" )           = IndexData("id")

                            if len(parameters)>2 then
                                for i=2 to len(parameters)-1 step 2
                                    metaData(parameters(i))=parameters(i+1)
                                next
                            end if
                    
                            IndexData( "meta_data" ) = metaData                    
                            IndexList = Dictionary( "path","string","modification_date","date","mime_type","string","url","string","size","float","index","string","author","string","acl","string","wf_priority","float","wf_name","string","wf_groups","string","wf_step_name","string","wf_limit","float","wf_step_max","float","wf_file_name","string","wf_file_path","string","wf_comments","string","Tasks_Id","string")

                            set cnxProSearch = new ProSearchApp
                                cnxProSearch.agentName = "Automation"
                                cnxProSearch.sourceName = "Tasks"
                                cnxProSearch.sourceindexes = IndexList
                        
                            if cnxProSearch.connect then
                                IndexData( "source_id" ) = cnxProSearch.GetSourceId
                                if cnxProSearch.send(IndexData)  then
                                    result = IndexData( "id" )
                                else
                                    result= ""
                                    executionState = False
                                    errMsg = "The function 'tache' returned an error while sending data to ProSearch : " & cnxProSearch.ErrMsg                                    
                                end if               
                            else
                                result= ""
                                executionState = False
                                errMsg = "The function 'tache' returned an error while connecting to ProSearch : " & cnxProSearch.ErrMsg                                    
                            end if
                  
                        case "delete"
                            set cnxProSearch = new ProSearchApp
                                cnxProSearch.agentName = "Automation"
                                cnxProSearch.sourceName = "Tasks"
                                cnxProSearch.sourceindexes = IndexList
                        
                            if cnxProSearch.connect then
                                if cnxProSearch.delete(parameters(1)) then
                                    result = true
                                else
                                    result= ""
                                    executionState = False
                                    errMsg = "The function 'tache' returned an error while deleting data in ProSearch : " & cnxProSearch.ErrMsg                                    
                                end if
                            else
                                result= ""
                                executionState = False
                                errMsg = "The function 'tache' returned an error while connecting to ProSearch : " & cnxProSearch.ErrMsg                                    
                            end if
                        case else
                            result= ""
                            executionState = False
                            errMsg = "The function 'tache' needs a string as a first parameter with the value 'new','update','delete'."                
                    end select               
                else
                    result= ""
                    executionState = False
                    errMsg = "The function 'tache' needs a string as a first parameter."              
                end if

            '##################################################
            '#                 FORM FUNCTIONS   (8)           #
            '##################################################
            '///// FORM BUILDER
            case "annulecheck"
                try 
                    avoidCheck = asjson(varContext.getVariable("_avoidCheck").value)
                    AppendToArray(avoidCheck,varContext.getVariable("btnGUID").value)
                    varContext.setVariable("_avoidCheck",tojson(avoidCheck))
                    result= varContext.getVariable("btnGUID").value
                catch scripterror as err
                    result= ""
                    executionState = False
                    errMsg = err.message
                end try
            
            case "onglet"
                chkFunctionOpt(funcName,parameters)
                If len(parameters) = 1 Then 
                    SessionDictionary("tab")=parameters(0)
                    formObject.Screen("Automation").Component("tabgroup1").addTab(SessionDictionary("tab"))
                else
                    result= ""
                    executionState = False
                    errMsg = "The function 'Onglet' need 1 parameter, here we have " & cstr(len(parameters)) & " provided."
                end if
            
            case "listeoptions"
            'listeoptions("nom","titre",height,liste_de_valeurs)
            'this function create a fieldset, the options are by 2 title & true or false to set if it's selected or not
            
                chkFunctionOpt(funcName,parameters)
                if SessionDictionary("tab")="" then
                    SessionDictionary("tab") = SessionDictionary("tabDefault")
                    formObject.Screen("Automation").Component("tabgroup1").addTab(SessionDictionary("tab"))
                end if
                If len(parameters) >= 4 Then
                    if typename(sessiondictionary("_listeoptions_"))="Array" then
                        AppendToArray(sessiondictionary("_listeoptions_"),normalize(parameters(0)))
                    else
                        sessiondictionary("_listeoptions_")=Array(normalize(parameters(0)))
                    end if
                    sessiondictionary("_listeoptions_" & normalize(parameters(0)))=Array
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).addComponent(parameters(0),"fieldset")
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).identifier = normalize(parameters(0))
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).label = parameters(1) 
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).height = parameters(2)
                    for i=3 to len(parameters)-1
                        AppendToArray(sessiondictionary("_listeoptions_" & normalize(parameters(0))),normalize(parameters(i)))
                        formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).addOption(parameters(i))
                        if varContext.getVariable(normalize(parameters(0))+"_"+normalize(parameters(i))).isNull then 
                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).setOptionState(parameters(i),false)                           
                        elseif varContext.getVariable(normalize(parameters(0))+"_"+normalize(parameters(i))).value = "on" then
                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).setOptionState(parameters(i),true)
                        else
                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).setOptionState(parameters(i),false)
                        end if
                    next
                else
                    result= ""
                    executionState = False
                    errMsg = "The function 'ListeOptions' need at least 4 parameters, here we have " & cstr(len(parameters)) & " provided."
                end if

            case "option"
                If len(parameters) = 2 Then
                    if varContext.getVariable(normalize(parameters(0))+"_"+normalize(parameters(1))).isNull then
                        result=false
                    elseif  varContext.getVariable(normalize(parameters(0))+"_"+normalize(parameters(1))).value = "on" then
                        result=true
                    else
                        result=false
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function 'Option' need 2 parameters, here we have " & cstr(len(parameters)) & " provided."
                end if
            case "champtexte"
            'ChampTexte("nom","label","valeur par default","Obligatoire oui | non")
            'This function create a component textbox for the given formobject 
            
                chkFunctionOpt(funcName,parameters)
                if SessionDictionary("tab")="" then
                    SessionDictionary("tab") = SessionDictionary("tabDefault")
                    formObject.Screen("Automation").Component("tabgroup1").addTab(SessionDictionary("tab"))
                end if
                If len(parameters) = 4 Then              
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).addComponent(parameters(0),"InputText")
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).label = cstr(parameters(1))
                    if cstr(parameters(2))<>"" then
                        if isDate(parameters(2)) then
                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).setToDate
                        end if
                        formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).defaultValue = parameters(2)
                    end if
                    if isDate(parameters(2)) then
                        formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).setToDate
                    end if
                    if parameters(3) then
                        formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).addRule("if(isNotEmpty(me),'ok','Ce champ ne peut pas être vide !')")
                    end if
                    result=parameters(0)    
                Else
                    result= ""
                    executionState = False
                    errMsg = "The function 'ChampTexte' need 4 parameters, here we have " & cstr(len(parameters)) & " provided."
                end If
            
            case "champliste"
            'ChampListe("nom","label","editable True|False","valeur prédéfinies séparées par ;","Obligatoire oui | non")
            'This function create a component textbox for the given formobject
                chkFunctionOpt(funcName,parameters)
                if SessionDictionary("tab")="" then
                    SessionDictionary("tab") = SessionDictionary("tabDefault")
                    formObject.Screen("Automation").Component("tabgroup1").addTab(SessionDictionary("tab"))
                end if
                sortData = true
                If len(parameters) = 5 and typename(parameters(2)) ="Boolean" and typename(parameters(4))="Boolean" Then              
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).addComponent(parameters(0),"DropDown")
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).label = parameters(1)
                    if parameters(2) then
                        formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).loadfromdb
                    end if
                    if parameters(4) then
                        formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).addRule("if(isNotEmpty(me),'ok','Vous devez choisir une valeur !')")
                    end if
                    editable = true
                    if instr(parameters(3),";")<>0 then
                        editable = false
                        opt = split(parameters(3),";")
                        for each elt in opt
                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).addOption(elt)
                        next
                    elseif instr(parameters(3),"[")<>0 and instr(parameters(3),"]")<>0 then
                        editable = false
                        listData = asjson(replace(parameters(3),"'",""""))
                        if len(listData)=1 or len(listData)=2 or len(listData)=3 then
                            select case lcase(listData(0))
                                case "utilisateurs"
                                    group = ""
                                    if len(listData)=2 then
                                        group = listData(1)
                                    end if
                                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).buildFromUsers(group)
                         
                                case "répertoires"
                                    sortData = false
                                    if len(listData)=2 or len(listData)=3  then
                                        targetPath = listData(1) 
                                        set objPath = ProShare.get_by_guid(varContext.getVariable("currentPathGUID").value)
                                        if mid(targetPath,1,2)="./" then
                                            targetPath = replace (targetPath,"./",objPath.path)  
                                        elseif mid(targetPath,1,3)="../" then
                                            targetPath = replace (targetPath,"../",objPath.parent.path + "/")
                                        elseif mid(targetPath,1,4)=".../" then
                                            targetPath = replace (targetPath,".../",objPath.parent.parent.path + "/")
                                        end if
                                        'logger "target PATH : " & targetPath
										if len(listData)=3 then
											howdeep = listData(2)
										else
											howdeep = 3
										end if
                                        formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).buildFromPath(targetPath,howdeep)
                                        if parameters(4) then
                                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).addRule("if(me<>'" & ProShare.get_by_path(targetPath).guid & "','ok','Vous devez choisir un répertoire !')")
                                        end if
                                    else
                                        result= ""
                                        executionState = False
                                        errMsg = "The function 'ChampListe' pamareter default Value must be 'utilisateur' or 'répertoire'."
                                    end if
                         
                                case "db"
                                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).loadfromdb
                                    if len(listData)=2 then
                                        if listData(1)<>"" then
                                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).selectedValue=listData(1)
                                        end if
                                    end if
                         
                                case else
                                    result= ""
                                    executionState = False
                                    errMsg = "The function 'ChampListe' pamareter default Value must be 'utilisateur' or 'répertoire'."
                       
                            end select      
                        else
                            result= ""
                            executionState = False
                            errMsg = "The function 'ChampListe' pamareter default Value must be an array of 1 or 2 value(s), a list with ';' to separate or 'db'."
                        end if
                    else
                        formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).selectedValue=parameters(3)
                    end if
                 
                    if parameters(2) and editable then
                        formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).setEditable
                    end if
                    if sortData then
                        formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component(parameters(0)).sort
                    end if
                        result=parameters(0)
                Else
                    errMsg = ""
                    result = ""
                    executionState = False               
                    if typename(parameters(2)) <> "Boolean" then
                        errMsg = "The parameter 'editable' must be a boolean value oui or non."
                    end if
                    if typename(parameters(4)) <> "Boolean" then
                        errMsg = errMsg + "The parameter 'Obligatoire' must be a boolean value oui or non."
                    end if
                    if len(parameters) <> 5 then
                        errMsg = errMsg + "The function '"+funcname+"' need 5 parameters, here we have " & cstr(len(parameters)) & " provided."
                    end if  
                end If
            
            case "commentaires"
            'Commentaires("nblignes")
                chkFunctionOpt(funcName,parameters)
                if SessionDictionary("tab")="" then
                    SessionDictionary("tab") = SessionDictionary("tabDefault")
                    formObject.Screen("Automation").Component("tabgroup1").addTab(SessionDictionary("tab"))
                end if
                If len(parameters) = 1 Then
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).addComponent("commentaires","textarea")
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component("commentaires").label = "Commentaires :"
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component("commentaires").row = parameters(0)
                Else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 1 parameters, here we have " & cstr(len(parameters)) & " provided."
                end If
           
            case "creerliste"
                'remplir une liste depuis un fichier CSV
                'arg 0 - chemin; arg 1 - separateur, arg 2 - colonne, arg 3 - dbname, arg 4 - entête, arg 5 - trié
                if len(parameters)=6 then
                    if dbdictionary("sourcelist")="" then
                        sourcelist = array
                    else
                        sourcelist = asjson(dbdictionary("sourcelist"))
                    end if
                    reloaddata = true
                    for each elt in sourcelist
                        if elt=parameters(0) then
                            reloaddata = false
                            exit for
                        end if                  
                    next
                    if dbdictionary("sourcelist_"+normalize(parameters(0)))<>"loaded" then
                        try
                            set filedata = ProShare.get_by_path(parameters(0)).open
                            boolskip=parameters(4)
                            optiondata = array
                            eof = false
                            do until eof
                                try
                                    dataline = filedata.readline
                                catch
                                    eof=true
                                end try
                                if boolskip or eof then
                                    boolskip = false
                                else
                                    if dataline="" then
                                        eof=true
                                    else
                                        val = split(dataline,parameters(1))(parameters(2))
                                        AppendToArray(optiondata,val)
                                    end if
                                end if
                            Loop
                            if parameters(5) then
                                ArraySort(optiondata)
                            end if

                            dbdictionary("_" & normalize(parameters(3)) & "_")=tojson(optiondata)
                            dbdictionary("sourcelist_"+normalize(parameters(0)))="loaded"
                            result="données chargées"
                        catch
                            result= ""
                            executionState = False
                            errMsg = "An error occured while loading file : " & parameters(0)
                        end try
                    else
                        result="données déjà chargées"
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 6 parameters, here we have " & cstr(len(parameters)) & " provided."
                end if
            
            case "listerecherchable"
            'this function allow to turn on searchable dropdown
                If len(parameters) = 1 Then 
                    formObject.Screen("Automation").DropDownSearchable = parameters(0)
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 1 parameter, here we have " & cstr(len(parameters)) & " provided."
                end if
            
            case "pdfviewer"
            'PDFViewer("Taille")
            'This add a component to see the doc if it's PDF.
                chkFunctionOpt(funcName,parameters)
                if SessionDictionary("tab")="" then
                    SessionDictionary("tab") = SessionDictionary("tabDefault")
                    formObject.Screen("Automation").Component("tabgroup1").addTab(SessionDictionary("tab"))
                end if
                If len(parameters) >= 1 Then
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).addComponent("PDF","PdfViewer")
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component("PDF").height = cstr(parameters(0))
                    if len(parameters)=2 then
                        if parameters(1)<0.75 then
                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component("PDF").ratio = 0.75
                        else
                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component("PDF").ratio = parameters(1)
                        end if
                    else
                        formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component("PDF").ratio = 0.75
                    end if
                Else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 1 parameter, here we have " & cstr(len(parameters)) & " provided."
                end If              
            
            case "setparams"
            'setParams("title",width,height,pdfAgauche)
                chkFunctionOpt(funcName,parameters)
                if SessionDictionary("tab")="" then
                    logger ("got tab")
                    SessionDictionary("tab") = SessionDictionary("tabDefault")
                    formObject.Screen("Automation").Component("tabgroup1").addTab(SessionDictionary("tab"))
                end if
                If len(parameters) = 3 or len(parameters) = 4  Then
                    if sessiondictionary("filemode")="multiple" then
                        filename = ProShare.get_by_guid(varContext.getVariable("fileGUID").value).name
                        formObject.Screen("Automation").Title = parameters(0) + " | " + SimpleIf(len(filename)>15,mid(filename,1,15)+"...",filename) + formatstring(" | ({0}/{1})",sessiondictionary("nbFileMax")-len(sessiondictionary("nodelist")),sessiondictionary("nbFileMax"))
                    else
                        formObject.Screen("Automation").Title = parameters(0)
                    end if
                    formObject.Screen("Automation").Width = parameters(1)
                    formObject.Screen("Automation").Height = parameters(2)
                    if len(parameters) = 4 then
                        if TypeName(parameters(3))="Boolean" then
                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).pdfTemplate = parameters(3)
                        end if
                    end if
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 3 or 4 parameters, here we have " & cstr(len(parameters)) & " provided."
                end if
            
            case "texte"
            'Texte("Ici le texte à afficher","Alignement gauche|centre|droit")
                chkFunctionOpt(funcName,parameters)
                if SessionDictionary("tab")="" then
                    logger ("got tab")
                    SessionDictionary("tab") = SessionDictionary("tabDefault")
                    formObject.Screen("Automation").Component("tabgroup1").addTab(SessionDictionary("tab"))
                end if
                If len(parameters) = 2 Then
                    randomize
                    'val = replace(cstr(rnd * 100000),".","-")
                    val=cstr(rnd * 100000)
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).addComponent("groupTXT-" & cstr(val),"group")
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component("groupTXT-" & cstr(val)).CreateCell(1)
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component("groupTXT-" & cstr(val)).Size(0)="100%"
                    select case lcase(parameters(1))
                        case "gauche"
                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component("groupTXT-" & cstr(val)).align(0) = "left"
                        case "centre"
                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component("groupTXT-" & cstr(val)).align(0) = "middle"
                        case "droit"
                            formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component("groupTXT-" & cstr(val)).align(0) = "right"
                    end select
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).addComponent("text-" & val,"Text")
                    formObject.Screen("Automation").Component("tabgroup1").Tab(SessionDictionary("tab")).Component("text-" & cstr(val)).value=parameters(0)
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 1 parameter, here we have " & cstr(len(parameters)) & " provided."
                end if  
            '///// FORM GETTER
            case "formulaire"
            'Formulaire("nom de la variable")
                If len(parameters) = 1 Then
                    result = xml_dialog.get_answer(normalize(parameters(0)))               
                else
                    result= ""
                    executionState = False
                    errMsg = "The function '"+funcname+"' need 1 parameters, here we have " & cstr(len(parameters)) & " provided."
                end if
            
            case "memoriseformulaire"
                for each variable in xml_dialog.get_answer
                    select case variable
                        case "macros_id","tabselected","step","sender","fileguid"
                            'we do nothing they are not form variable
                        case else
                            if not(instr(variable,":")<>0 and xml_dialog.get_answer(variable)="on") then
                                varContext.setVariable(variable,xml_dialog.get_answer(variable))
                            end if
                    end select
                next
    
                
            '##################################################
            '#           EAC Managment        (11)            #
            '##################################################
            
            case "sendeacmail"
              'Send an EAC email that run an EAC content linked to the current folder
              '0 - destinataires, 1- subject, 2- EAC method
              result=sendEACmail(parameters,executionState,errMsg,varContext)
            
            case "eacanswer"
              'function that return a valid XML for EAC
              '0- EAC Template
              result=EACAnswer(parameters,executionState,errMsg,varContext)
           
            case "toxmlentity"
              'convert non ASCII char into XML entity &#xxxx;
              '0- input string
              result=ToXMLEntity(parameters,executionState,errMsg,varContext)
              
            case "cdata"
              'Add CDATA
              '0- input string
              result=cdata(parameters,executionState,errMsg,varContext)
              
            case "eacevt"
              'Add CDATA
              '0- input string
              result=EACEvt(parameters,executionState,errMsg,varContext)
              
            case "eacaction"
              'Add CDATA
              '0- input string
              result=EACAction(parameters,executionState,errMsg,varContext)
            
            case Else
                result= ""
                executionState = False
                errMsg = "The function (" & funcName & ") was not found in the library"
        End Select
    End Sub
End Class