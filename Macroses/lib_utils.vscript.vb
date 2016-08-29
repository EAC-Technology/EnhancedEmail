'####################################################################################################
'#                        Various functions that can simplify some operation                        #
'#--------------------------------------------------------------------------------------------------#
'# Some functions commonly used in macro.                                                           #
'#                                                                                                  #
'####################################################################################################

'Guid of macro
macrodic = Dictionary("Action","0b8b7eaf-439a-40f4-80e6-f9532ec980b7")


'-------------------------------------- ParseArg -----------------------------------------------
' if the key exist in args dictionary that should be this one made with arguments passed
' to the macro, return the value, if not return default value
'-----------------------------------------------------------------------------------------------

Function ParseArg( key, defaultValue )
  if key in args then
    if args(key) <> "" then
      ParseArg = args( key )
    else
      ParseArg = defaultValue
    end if
  else
    ParseArg = defaultValue
  end if
End Function

'-------------------------------------- SaveToDBdictionary -------------------------------------
' Save the value un the internal db dictionary, it value is Empty distroy the key.
' 
'-----------------------------------------------------------------------------------------------

Sub SaveToDBdictionary( key, value )
  if isEmpty( value ) then
    dbdictionary.remove( key )
  else
    dbdictionary( key ) = value 
  end if
End Sub

'-------------------------------------- IsEmptyArray -------------------------------------------
' Return true is the Array is empty, false if not.
' 
'-----------------------------------------------------------------------------------------------

Function IsEmptyArray( massive )
  if typename(massive)="Array" then
    if len(massive) <> 0 then
      isEmptyArray = False
    else
      isEmptyArray = True
    end if
  end if
End Function

'-------------------------------------- AppendToArray ------------------------------------------
'append new value to the end of array
' 
'-----------------------------------------------------------------------------------------------

Function AppendToArray( mass, value )
  if typename(mass)="Array" then
    array_len = -1
    if not IsEmptyArray( mass ) then
      array_len = ubound( mass )
    end if
    Redim Preserve mass( array_len + 1 )
    mass( array_len + 1 ) = value
    AppendToArray =  array_len + 1
  end if
End Function

'-------------------------------------- RemoveFromArray ------------------------------------------
'remove a value from an array
' 
'-----------------------------------------------------------------------------------------------

Function RemoveFromArray(mass, value )
'  logger ("RemoveFromArray IN :" & tojson(mass))
'  logger ("Value to DELETE :" & value)
  if typename(mass)="Array" then
    oldArray = mass
    mass = Array
    for each element in oldArray
      if element<>value then
        AppendToArray(mass,element)
      end if
    next
  end if
'  logger ("RemoveFromArray OUT :" & tojson(mass))
End Function

'-------------------------------------- ExistInArray ------------------------------------------
'Check if a value exist already in an Array
' 
'-----------------------------------------------------------------------------------------------
Function ExistInArray(mass, value )
  ExistInArray = false
  if typename(mass)="Array" then
    for each element in mass
      if element=value then
        ExistInArray = true
        exit for
      end if
    next
  end if
end function

'-------------------------------------- normalize ----------------------------------------------
'Remove space & accents
' 
'-----------------------------------------------------------------------------------------------

function normalize(byval str)

  normint=""
  accents     = "àáâãäçèéêëìíîïñòóôõöùúûüýÿÀÁÂÃÄÇÈÉÊËÌÍÎÏÑÒÓÔÕÖÙÚÛÜÝ"
  noaccents   = "aaaaaceeeeiiiinooooouuuuyyAAAAACEEEEIIIINOOOOOUUUUY"

  str = lcase(str)
  'delete space
  str = replace (str," " ,"_")
  str = replace (str,"'" ,"_")
  str = replace (str,"’" ,"_")
  str = replace (str,"°" ,"_")
  'delete special char
  set r = new RegExp
  r.IgnoreCase = true
  r.Global = true
  r.Pattern = "[\\/:""*?<>|]+"
  str = r.Replace(str, "")
  
  'remove accents
  for i=1 to len(str)
    v = mid(str,i,1)
    if InStr(accents,v) <> 0 then
     normint = normint + mid(noaccents,InStr(accents,v),1)
    else
     normint = normint + v
    end if
  next
  
  normalize = normint
  
end function

'-------------------------------------- toascii ------------------------------------------------
' convert with the notation &#xxx; non ascii stuff
' 
'-----------------------------------------------------------------------------------------------

function toascii(byval str)

  toascii=""
  for i=1 to len(str)
    v = mid(str,i,1)
    a = asc(v)
    if (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=95 and a<=122) then
      toascii=toascii+v
    else
      toascii=toascii+"&#"+cstr(a)+";"
    end if
  next

end function

'-------------------------------------- getAppName ---------------------------------------------
' Return the name of the application stored in ProAdmin
' 
'-----------------------------------------------------------------------------------------------

function getAppName(appName)
  
    select case lcase(appName)
  
    case "proadmin"
      appGuid = "491d4c93-4089-4517-93d3-82326298da44"
    case "procontact"
      appGuid = "526ae088-8004-469c-9d8e-cea715f8f63b"
    case "proplanning"
      appGuid = "22d43054-9861-48e8-875f-53d09bb1fd11"
    case "prosearch"
      appGuid = "fc2221b2-794b-4c40-991f-6c7c2f61dbc2"
    case "proshare"
      appGuid = "b0a274f0-22bc-44be-be48-da6ec9180268"
      
  end select
  
  AppSearchName = ""
  
  for each app in proadmin.registeredApplications
    if app.guid = appGuid then
      AppSearchName = app.title
      Exit For
    end if
  next
  
  if AppSearchName="" then
    getAppName = ""
  else
    getAppName = AppSearchName
  end if

end function

'-------------------------------------- ApplicationConnection -------------------------------
' Return an object that represent a connection to the App given in parameters
' 
'-----------------------------------------------------------------------------------------------

Function ApplicationConnection( appName, host, login, password )

  select case lcase(appName)
  
    case "proadmin"
      appGuid = "491d4c93-4089-4517-93d3-82326298da44"
    case "procontact"
      appGuid = "526ae088-8004-469c-9d8e-cea715f8f63b"
    case "proplanning"
      appGuid = "22d43054-9861-48e8-875f-53d09bb1fd11"
    case "prosearch"
      appGuid = "fc2221b2-794b-4c40-991f-6c7c2f61dbc2"
    case "proshare"
      appGuid = "b0a274f0-22bc-44be-be48-da6ec9180268"
      
  end select
  
  AppSearchName = ""
  
  for each app in proadmin.registeredApplications
    if app.guid = appGuid then
      AppSearchName = app.title
      Exit For
    end if
  next
  
  if AppSearchName="" then
    ApplicationConnection = array(0)
    Exit Function
  else
    'logger "App Name for " & appName & " is: " & AppSearchName

    Try
      Set ServerConnection = New WHOLEConnection
      ServerConnection.Open( host, login, password )
    Catch 
      ApplicationConnection = array(1)
      Exit Function
    End Try
    
    Set App = Nothing
    Try
      Set App = ServerConnection.Applications( AppSearchName )
    Catch
      ApplicationConnection = array(2)
      Exit Function
    End Try
    t = array(3,1)
    set t(1) = App
    ApplicationConnection = t
  end if
End Function

'-------------------------------------- BuildPath ----------------------------------------------
' If doesn't exist build a path given in paramter
' 
'-----------------------------------------------------------------------------------------------
function BuildPath(path)
  try
    pathArray = split(path,"/") 
    for each fol in pathArray
      if fol="" then 
        curPath=""
      else
        set getNode = ProShare.get_by_path(curPath & "/" & fol)
        if TypeName(getNode) = "Nothing" then
          if curPath="" then curPath="/"
            ProShare.get_by_path(curPath).mkdir(fol)
          end if
          if curPath="/" then
            curPath = "/" & fol
          else
            curPath = curPath & "/" & fol
          end if
        end if
    next
    BuildPath = true
  catch
    BuildPath = false
  end try
end function

'-------------------------------------- EscapeStr ----------------------------------------------
' Escape special chat to fit with HTML representation
' 
'-----------------------------------------------------------------------------------------------

Function EscapeStr( data )  
  
  data = Replace( data, "&", "&" & "amp;"  )
  data = Replace( data, "<", "&" & "lt;"   )
  data = Replace( data, ">", "&" & "gt;"   )
  data = BBCode( data )
  EscapeStr = data
  
End Function

'-------------------------------------- BBCode ----------------------------------------------
' Replace specific code to its equivalent in HTML
' 
'-----------------------------------------------------------------------------------------------

Function BBCode( data )

  data = Replace( data, "[b]", "<strong>" )
  data = Replace( data, "[/b]", "</strong>" ) 
  
  data = Replace( data, "[i]", "<em>" )
  data = Replace( data, "[/i]", "</em>" ) 
  
  data = Replace( data, "[u]", "<span style=""text-decoration: underline;"">" )
  data = Replace( data, "[/u]", "</span>" ) 
    
  data = Replace( data, "[s]", "<span style=""text-decoration:line-through"">" )
  data = Replace( data, "[/s]", "</span>" ) 
  
  data = Replace( data, "[code]", "<pre>" )
  data = Replace( data, "[/code]", "</pre>" ) 
  
  data = Replace( data, "[crlf]", "<br/>" )
  
  BBCode = data

End Function

'-------------------------------------- WriteToHistory ------------------------------------------
' Write data into a specific dictionary attached to a file guid key.
' 
'-----------------------------------------------------------------------------------------------

sub WriteToHistory(file_guid,type,comment)
 
    set file = ProShare.node.get_by_guid( file_guid )
    set user = ProAdmin.current_user
    
    if TypeName(DBdictionary(file.guid)) = "Empty" then
      FileData = Dictionary("history",dictionary,"context","")
    else
      FileData = asjson(replace(DBdictionary(file_guid),"""history"": {}","""history"": """""))
    end if
    if typename(FileData("history"))="Dictionary" then
      InternalData       = FileData("history")
    else
      InternalData       = Dictionary
    end if
    
    newComment         = Dictionary
    newComment("who")  = user.name
    newComment("date") = now
    newComment("msg")  = EscapeStr(comment)
    
    select case type
      case "indoc"
        newComment("type")="arrive_doc.png"
      case "move"
        newComment("type")="deplacement.png"
      case "comment"
        newComment("type")="commentaire.png"
      case "updatestate"
        newComment("type")="change.png"
    end select
    randomize
    guid = md5(cstr(now)+cstr(rnd*1000))
    InternalData(guid)=newComment
    FileData("history") = InternalData
    DBdictionary(file.guid)= tojson(FileData)

end sub

'-------------------------------------- setRights ----------------------------------------------
' Generate tasks to modify the rights of a user over a given path. 
' 
'-----------------------------------------------------------------------------------------------

function setRights(path,userEmail)

  folder_list = Array              
  parents_path = split(path,"/")
  pathstring=""
  AppendToArray(folder_list,"/")
              
  for each elt in parents_path
    if elt <> "" then
      pathstring = pathstring & "/" & elt
      AppendToArray(folder_list,pathstring)
    end if
  next
              
  if dbdictionary( "RightToManage" )= "" then
    ArrayOfTasks=Array
  else
    ArrayOfTasks=asjson(dbdictionary( "RightToManage" ))
  end if
              
  AppendToArray(ArrayOfTasks,dictionary(userEmail,folder_list))
  dbdictionary( "RightToManage" )=ToJson(ArrayOfTasks)
  set objTimer = GetTimer("BackgroundSetRights")
      objTimer.is_active=true
              
  Logger ("Task to manage rights for users had been added to queue : " & dbdictionary( "RightToManage" ))

end function

'-------------------------------------- FormatEmail --------------------------------------------
' Replace each key like {KeyWord} in a given string by it's value & format it to be sent by mail. 
' 
'-----------------------------------------------------------------------------------------------

Function FormatEmail(mail,data)

  mailformated = EscapeStr(mail)
  mailformated = replace(mailformated,chr(10),"<br>")
  if lcase(TypeName(data))="dictionary" then
    for each key in data
       mailformated = replace(mailformated,"{" & key & "}",data(key))
    next
  end if
  
  FormatEmail = "<HTML><BODY> " & mailformated & "</BODY></HTML>"
end function

'-------------------------------------- FileExists ---------------------------------------------
' Return true if the file with absolute path exist & false if not. 
' 
'-----------------------------------------------------------------------------------------------

Function FileExists(path)
  set file = ProShare.get_by_path(path)
  if isNothing(file) then
    FileExists = false
  else
    FileExists = true
  end if
End Function

'-------------------------------------- FixFilename ---------------------------------------------
' Delete from the file name the unapropriate char
' 
'-----------------------------------------------------------------------------------------------

Function FixFilename(filename) 
  set r = new RegExp
  r.IgnoreCase = true
  r.Global = true
  r.Pattern = "[\\/:""*?<>|]+"
  FixFileName = r.Replace(filename, "")
End Function

'-------------------------------------- MoveFile -----------------------------------------------
' Move a file from the current place to the target one. 
' 
'-----------------------------------------------------------------------------------------------

function MoveFile(ObjFile,targetPath)

  fileGUID = objFile.guid
  fileData = dbdictionary( fileGUID )
  
  if mid(targetPath,1,2)="./" then
    targetPath = replace (targetPath,"./",ObjFile.parent.path + "/")  
  elseif mid(targetPath,1,3)="../" then
    targetPath = replace (targetPath,"../",ObjFile.parent.parent.path + "/")
  elseif mid(targetPath,1,4)=".../" then
    targetPath = replace (targetPath,".../",ObjFile.parent.parent.parent.path + "/")
  end if

  buildPath(targetPath)
  
  while FileExists(targetPath & "/" & ObjFile.name) 
      ObjFile.name = "_" & ObjFile.name
  wend

  filename = ObjFile.name  
  ObjFile.copy(targetPath)
  set objFolderTarget = ProShare.get_by_path(targetPath)
  readRules = objFolderTarget.rules( Empty, "r" )
  writeRules = objFolderTarget.rules( Empty, "w" )
  ownerRules = objFolderTarget.rules( Empty, "d" )
  ObjFile.delete
  set newFile = ProShare.get_by_path(targetPath & "/" & filename)
  for each srule in readRules
    newFile.addRule(srule.subject,"r")
  next
  for each srule in writeRules
    newFile.addRule(srule.subject,"w")
  next
  for each srule in ownerRules
    newFile.addRule(srule.subject,"d")
  next
  dbdictionary( newFile.guid ) = fileData
  WriteToHistory(newFile.guid,"move","Déplacement du fichier vers : " & targetPath) 
  
  set MoveFile = newFile

end function

'-------------------------------------- toISODate -----------------------------------------------
' Take a local date & transform it to ISO format
' 
'-----------------------------------------------------------------------------------------------

Function toISODate(inDate)
	'if isDate(inDate) then
		jour	= mid(cstr(indate),1,2)
		mois  = mid(cstr(indate),4,2)
		annee   = mid(cstr(indate),7,4)
		toISODate = annee & "-" & mois & "-" & jour
	'Else
	'	toISODate = ""
	'End If
End Function

'-------------------------------------- fromISODate -----------------------------------------------
' reverse teh toISODate function
' 
'-----------------------------------------------------------------------------------------------

Function fromISODate(inDate)

	jour	= mid(cstr(indate),9,2)
	mois  = mid(cstr(indate),6,2)
	annee   = mid(cstr(indate),1,4)
	fromISODate = cstr(jour & "/" & mois & "/" & annee)
End Function

'-------------------------------------- SimpleIf -----------------------------------------------
' make If as a function
' 
'-----------------------------------------------------------------------------------------------

Function SimpleIf( expr, ret1, ret2 )
  if expr then
    SimpleIf = ret1
  else
    SimpleIf = ret2
  end if
End Function

'-------------------------------------- CreateGUID ---------------------------------------------
' generate a GUID like string
' 
'-----------------------------------------------------------------------------------------------

Function CreateGUID()
  Randomize
  Dim tmpCounter, tmpGUID
  Const strValid = "0123456789abcdef"
  For tmpCounter = 1 To 8
    tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  Next
  tmpGUID = tmpGUID & "-"
  For tmpCounter = 1 To 4
    tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  Next
  tmpGUID = tmpGUID & "-"
  For tmpCounter = 1 To 4
    tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  Next
  tmpGUID = tmpGUID & "-"
  For tmpCounter = 1 To 4
    tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  Next
  tmpGUID = tmpGUID & "-"
  For tmpCounter = 1 To 12
    tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  Next
  CreateGUID = tmpGUID
End Function

'-------------------------------------- FileDataManagement -------------------------------------
' Object that represent a file & its meta data
' 
'-----------------------------------------------------------------------------------------------

class FileDataManagement

  dim fileGUID
  dim FileData
  

  Sub Class_Initialize( )
    try      
      set file = page_status.selected_nodes(0)
      fileGUID = file.guid
      loadFileDataFromDB
    catch
    end try
  end sub
  
  sub loadFileDataFromDB
    FileData= Dictionary
    if TypeName(DBdictionary(fileGUID)) = "Empty" then
      FileData = Dictionary("history",dictionary,"context","")
    else
      FileData = asjson(replace(DBdictionary(fileGUID),"""history"": {}","""history"": """""))
    end if

  end sub
  
  function getHistory
    getHistory = tojson(FileData("history"))
  end function
  
  sub WrToHistory(type,msg)
    WriteToHistory(fileGUID,type,msg)
  end sub
  
  function getContext
 
    getContext = FileData("context")
  
  end function
  
  function storeContext(data)
    loadFileDataFromDB
    if typename(data)="Dictionary" then
      if len(data) = 0 then
        FileData("context") = ""
      else
        FileData("context") = data
      end if
    else
      FileData("context") = data
    end if
    DBdictionary(FileData("context")("fileGUID")("value"))= tojson(FileData)
  end function
  
  function getRulesList
     getActionList = FileData("rules")
  end function
  
end class

'-------------------------------------- Progress -------------------------------------
' This class allow to transmit data to the API progress indicator
' 
'-------------------------------------------------------------------------------------

class ProgressIndicator

  Sub Class_Initialize( )
     if sessiondictionary("ProgressID")="" then
       sessiondictionary("ProgressID")=CreateGUID()
     end if
  end sub

  sub Progress (label,value,max)

    if sessiondictionary("ProgressID")<>"" then
      dbdictionary(sessiondictionary("ProgressID"))= tojson(dictionary ("label",cstr(label),"value",value/max))
    end if
  
  end sub

  Sub Class_Terminate()
    if sessiondictionary("ProgressID")<>"" then
      if dbdictionary(sessiondictionary("ProgressID"))<>"" then
        dbdictionary.remove(sessiondictionary("ProgressID"))
      end if
    end if
  end sub

end class

'-------------------------------------- DateToUnixInt ------------------------------------------
' Convert a date to its unix format as a int in second from 01/01/1970
' 
'-----------------------------------------------------------------------------------------------

function DateToUnixInt(dateToConvert)
 
  if typename(dateToConvert)="Date" then
    DateToUnixInt=DateDiff ("s",cdate("01/01/1970"),dateToConvert)
  else
    DateToUnixInt=0
  end if

end function


'-------------------------------------- UnixIntToDate ------------------------------------------
' Convert a date to its unix format as a int in second from 01/01/1970
' 
'-----------------------------------------------------------------------------------------------

function UnixIntToDate(intToConvert)

  if true then
    UnixIntToDate = DateAdd("s",intToConvert,cdate("01/01/1970"))  
  end if

end function


'-------------------------------------- numInvoieTokenizer ------------------------------------------
' Usefull function needed for Analyse function in lib
' 
'-----------------------------------------------------------------------------------------------

function numInvoieTokenizer(curchar)
    select case lcase(curchar)
      case "|"
          numInvoieTokenizer ="[NEEDED]"
      case "?"
          numInvoieTokenizer ="[ANY]"
      case "µ"
          numInvoieTokenizer ="[INVARIANT]"
      case "0","1","2","3","4","5","6","7","8","9","~"
          numInvoieTokenizer ="[NUMERIC]"
      case "a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","^"
          numInvoieTokenizer ="[ALPHA]"
      case "/","-",",","_","#"," "
          numInvoieTokenizer ="[SPECIAL]"
      case else
          numInvoieTokenizer ="[UNKNOWN]"
	end select
end function

'-------------------------------------- opMerging ----------------------------------------------
' function to merge 2 char in the context of learning procedure
' 
'-----------------------------------------------------------------------------------------------

function opMerging(mergedchar,curchar)

  select case numInvoieTokenizer(curchar)
      
      case "[NUMERIC]"
          select case numInvoieTokenizer(mergedchar)
              case "[NUMERIC]"
                  if mergedchar=curchar then
                      opMerging=curchar
                  else
                      opMerging ="~"
                  end if
              case "[INVARIANT]"
                  opMerging=curchar
              case else
                  opMerging ="?"             
          end select
          
      case  "[ALPHA]"
          select case numInvoieTokenizer(mergedchar)
              case "[ALPHA]"
                  if mergedchar=curchar then
                    opMerging=curchar
                  else
                    opMerging="^"
                  end if
              case "[INVARIANT]"
                  opMerging=curchar
              case else
                  opMerging ="?"             
          end select      
      
      case "[SPECIAL]"
          select case numInvoieTokenizer(mergedchar)
              case "[SPECIAL]"
                  if mergedchar=curchar then
                    opMerging=curchar
                  else
                    opMerging="?"
                  end if
              case "[INVARIANT]"
                  opMerging=curchar
              case else
                  opMerging ="?"             
          end select      
      
      case "[UNKNOWN]"
          select case numInvoieTokenizer(mergedchar)
              case "[UNKNOWN]"
                  if mergedchar=curchar then
                    opMerging=curchar
                  else
                    opMerging="?"
                  end if
              case "[INVARIANT]"
                  opMerging=curchar
              case else
                  opMerging ="?"             
          end select      
      case else
          opMerging ="?"
  
  end select
      
end function

'-------------------------------------- BuildTemplate ------------------------------------------
' build a tempalte for the regEx generation based on exemples
' exmaxlen : Max len of the examples
'-----------------------------------------------------------------------------------------------

function BuildTemplate(exemples)
	if typename(exemples)="Array" then
		exmaxlen = 0
		for each ex in exemples
			if len(ex)>exmaxlen then
				exmaxlen = len(ex)
			end if
		next	
		outputTemplate = ""
		for i=1 to exmaxlen
			mergedchar = "µ"
			for each ex in exemples
				try
					curchar = mid(ex,i,1) 
				catch
					curchar = "µ"
				end try
				mergedchar = opMerging(mergedchar,curchar)
			next
			if mergedchar="^" then
				mergedchar="A"
			elseif numInvoieTokenizer(mergedchar) = "[ALPHA]" then
				mergedchar = "|"+mergedchar
			end if
            if mergedchar="~" then
				mergedchar="0"
			elseif numInvoieTokenizer(mergedchar) = "[NUMERIC]" then
				mergedchar = "|"+mergedchar
			end if
			outputTemplate = outputTemplate + mergedchar
			logger("outputTemplate : " & outputTemplate)
		next
		BuildTemplate = outputTemplate
	else
		BuildTemplate = ""
	end if
end function

'-------------------------------------- AnalyseBlock ------------------------------------------
' cut sequence with separators & process analyse on each block recursively
' 
'-----------------------------------------------------------------------------------------------

function AnalyseBlock(exemples,separators)

  if typename(exemples)="Array" and typename(separators)="Array" then   
    if len(separators)=0 then
      'end of recurtion 
      AnalyseBlock = BuildTemplate(exemples)
    else
        curSeparator = separators(0)
        RemoveFromArray(separators,curSeparator)
        nbBlock = len(split(exemples(0),curSeparator))
        BlockArray = Array
        template = ""
        for each ex in exemples
            AppendToArray(BlockArray,Split(ex,curSeparator))
        next
        for i=0 to nbBlock-1
            tempBlock = Array
            for each block in BlockArray
                try
                    AppendtoArray(tempBlock,block(i))
                catch
                    AppendtoArray(tempBlock,"")
                end try
            next
            template = template + curSeparator + AnalyseBlock(tempBlock,separators)
        next
      AnalyseBlock = mid(template,2)
    end if
  end if
end function

'-------------------------------------- VDOMEvent ----------------------------------------------
' Generate a js to launch a macro.
' 
'-----------------------------------------------------------------------------------------------

function VDOMEvent(macro,data)  
  VDOMEvent=""
  if TypeName(data)="Dictionary" then
    if macro in macrodic then
      data("macros_id")=macrodic(macro)
      VDOMEvent = "javascript:execEventBinded('724748df_a464_486f_a324_4e9b3fc36fe0', 'custom', " & replace(tojson(data),"""","'") & ")"
    end if
  end if
end function


'-------------------------------------- genTable ----------------------------------------------
' Allow to generate a table with data
' 
'-----------------------------------------------------------------------------------------------

function genTable(header,data,nbLigne,pos,color,csvname,token)

  refreshtable = false
  pos = cint(pos)
  mode = "normal"
  if nbLigne = 0 then
    mode="full"
    nbligne = len(data)
  end if
  if mode="normal" then
    if token = "" then
        token = cstr(cint(rnd * 1000))
        key = csvname+"_"+token 
        if dbdictionary("tabledata_session")<>"" then
            tabledata_session =  asjson(dbdictionary("tabledata_session"))
            tabledata_session(key) = now
            dbdictionary("tabledata_session") = tojson(tabledata_session)
            set objTimer = getTimer("Session Clean up")
            objTimer.is_active=true
        else
            tabledata_session =  Dictionary
            tabledata_session(key) = now
            dbdictionary("tabledata_session") = tojson(tabledata_session)
        end if
        dbdictionary(key)= tojson(Dictionary("header",header,"data",data,"nbLigne",nbLigne,"color",color))
    else
      refreshtable = true
    end if
  else
    token = cstr(cint(rnd * 1000))
  end if
  
  if mode="normal" then
      colorCss = Dictionary("blue","genTable-blue.css","brown","genTable-brown.css","gray","genTable-gray.css","green","genTable-green.css","purple","genTable-purple.css","red","genTable-red.css")
      if lcase(color) in colorCss then
          tableStyle = "<Style>" & replace(resources.open(colorCss(lcase(color))).getValue,"datagrid","datagrid-"+token) & "</Style>"
      else
          tableStyle = "<Style>" & replace(resources.open(colorCss("blue")).getValue,"datagrid","datagrid-"+token) & "</Style>"    
      end if
  end if
  
  set tableData = buffer.create
      if not refreshtable then
          tableData.write("<div id='datagrid_"+token+"' class='datagrid-"+token+"'>")
      end if
      tableData.write("<table "+SimpleIf(mode="normal","","style='border:solid 1px #777777; border-collapse:collapse; font-family:verdana; font-size:11px;'")+" >")
      tableData.write("   <thead>")
      tableData.write("      <tr "+SimpleIf(mode="normal","","style='background-color:lightgrey;'")+">")
      cpt=0
      selcol= -1
    for each col in header
      tableData.write("         <th>"&cstr(col)&"</th>")
      if cstr(col)="sel" then
        selcol=cpt
      end if
      cpt = cpt + 1
    next
      tableData.write("      </tr>")
      tableData.write("   <thead>")
      tableData.write("   <tfoot>")
      tableData.write("      <tr>")
      tableData.write("         <td colspan="""&cstr(len(header))&""">")
      tableData.write("            <div id=""paging"">")
    if len(data)>nbligne then
      tableData.write("               <ul>")
    if pos<>1 then
      tableData.write("                  <li><a href=""javascript:getTable('" + csvname + "','" + cstr(pos-1) + "','" + token + "')""><span>Précédent</span></a></li>")
    end if
    for k=1 to (int(len(data)/nbLigne))+1
      tableData.write("                  <li><a href=""javascript:getTable('" + csvname + "','" + cstr(k) + "','" + token + "')"" "&SimpleIf(k=pos,"class=""active""","")&" ><span>"&k&"</span></a></li>")
    next
    if pos*nbligne<len(data) then
      tableData.write("                  <li><a href=""javascript:getTable('" + csvname + "','" + cstr(pos+1) + "','" + token + "')""><span>Suivant</span></a></li>")
    end if
      tableData.write("              </ul>")
    end if
      tableData.write("           </div>")
      tableData.write("     </tr>")
      tableData.write("   </tfoot>")
      tableData.write("   <tbody>")
   if (pos-1)*nbligne+nbligne>len(data) then
     max = len(data)-1
   else
     max = (pos-1)*nbligne+nbligne-1
   end if
   for i=(pos-1)*nbligne to max
      line = data(i)
      tableData.write("      <tr>") 
      for j=0 to len(header)-1
        if j=selcol then
          if mode="normal" then
              tableData.write("        <td><INPUT type='checkbox' name='"+csvname+":"+cstr(line(j))+"'></INPUT></td>")
          else
              tableData.write("        <td></td>")
          end if
        else
          tableData.write("        <td "+SimpleIf(mode="normal","","style='padding-left:10px; color:midnightblue;'")+" >"&cstr(line(j))&"</td>")
        end if
      next
      tableData.write("      </tr>")
   next
      tableData.write("   </tbody>")
      tableData.write("</table>")
      if not refreshtable then
          tableData.write("</div>")
      end if
   genTable = tableStyle & tableData.getvalue
   
end function

'-------------------------------------- genEditableTable ---------------------------------------
' Allow to generate an editable table with data
' 
'-----------------------------------------------------------------------------------------------


function genEditableTable(header,data,nbLigne,pos,color,csvname,token,editable)

  refreshtable = false
  pos = cint(pos)
  mode = "normal"
  if nbLigne = 0 then
    mode="full"
    nbligne = len(data)
  end if
  if mode="normal" then
    if token = "" then
        token = cstr(cint(rnd * 1000))
        key = csvname+"_"+token 
        if dbdictionary("tabledata_session")<>"" then
            tabledata_session =  asjson(dbdictionary("tabledata_session"))
            tabledata_session(key) = now
            dbdictionary("tabledata_session") = tojson(tabledata_session)
            set objTimer = getTimer("Session Clean up")
            objTimer.is_active=true
        else
            tabledata_session =  Dictionary
            tabledata_session(key) = now
            dbdictionary("tabledata_session") = tojson(tabledata_session)
        end if
        dbdictionary(key)= tojson(Dictionary("header",header,"data",data,"nbLigne",nbLigne,"color",color,"editable",editable))
    else
      refreshtable = true
    end if
  else
    token = cstr(cint(rnd * 1000))
  end if
  
  if mode="normal" then
      colorCss = Dictionary("blue","genTable-blue.css","brown","genTable-brown.css","gray","genTable-gray.css","green","genTable-green.css","purple","genTable-purple.css","red","genTable-red.css")
      if lcase(color) in colorCss then
          tableStyle = "<Style>" & replace(resources.open(colorCss(lcase(color))).getValue,"datagrid","datagrid-"+token) & "</Style>"
      else
          tableStyle = "<Style>" & replace(resources.open(colorCss("blue")).getValue,"datagrid","datagrid-"+token) & "</Style>"    
      end if
  end if
  
  set tableData = buffer.create
      if not refreshtable then
          tableData.write("<div id='datagrid_"+token+"' class='datagrid-"+token+"'>")
      end if
      tableData.write("<table "+SimpleIf(mode="normal","","style='border:solid 1px #777777; border-collapse:collapse; font-family:verdana; font-size:11px;'")+" >")
      tableData.write("   <thead>")
      tableData.write("      <tr "+SimpleIf(mode="normal","","style='background-color:lightgrey;'")+">")
      cpt=0
      selcol= -1
    for each col in header
      tableData.write("         <th>"&cstr(col)&"</th>")
      if cstr(col)="sel" then
        selcol=cpt
      end if
      cpt = cpt + 1
    next
      tableData.write("      </tr>")
      tableData.write("   <thead>")
      tableData.write("   <tfoot>")
      tableData.write("      <tr>")
      tableData.write("         <td colspan="""&cstr(len(header))&""">")
      tableData.write("            <div id=""paging"">")
    if len(data)>nbligne then
      tableData.write("               <ul>")
    if pos<>1 then
      tableData.write("                  <li><a href=""javascript:getTable('" + csvname + "','" + cstr(pos-1) + "','" + token + "')""><span>Précédent</span></a></li>")
    end if
    for k=1 to (int(len(data)/nbLigne))+1
      tableData.write("                  <li><a href=""javascript:getTable('" + csvname + "','" + cstr(k) + "','" + token + "')"" "&SimpleIf(k=pos,"class=""active""","")&" ><span>"&k&"</span></a></li>")
    next
    if pos*nbligne<len(data) then
      tableData.write("                  <li><a href=""javascript:getTable('" + csvname + "','" + cstr(pos+1) + "','" + token + "')""><span>Suivant</span></a></li>")
    end if
      tableData.write("              </ul>")
    end if
      tableData.write("           </div>")
      tableData.write("     </tr>")
      tableData.write("   </tfoot>")
      tableData.write("   <tbody>")
   if (pos-1)*nbligne+nbligne>len(data) then
     max = len(data)-1
   else
     max = (pos-1)*nbligne+nbligne-1
   end if
   for i=(pos-1)*nbligne to max
      line = data(i)
      tableData.write("      <tr>") 
      for j=0 to len(header)-1
        if j=selcol then
          if mode="normal" then
              tableData.write("        <td><INPUT type='checkbox' name='"+csvname+"I"+cstr(line(j))+"'></INPUT></td>")
			  lign = cstr(line(j))
          else
              tableData.write("        <td></td>")
          end if
        else
          tableData.write("        <td " & SimpleIf(mode="normal","","style='padding-left:10px; color:midnightblue;'")+" >" & SimpleIf(editable(header(j)),"<input class='inputFullWidth ' type='text' name='"+lign+":"+header(j)+"' value='"+cstr(line(j))+"' onchange='$(""input[name="+csvname+"I"+lign+"]"").attr(""checked"",true)'>",cstr(line(j))) & "</td>")
        end if
      next
      tableData.write("      </tr>")
   next
      tableData.write("   </tbody>")
      tableData.write("</table>")
      if not refreshtable then
          tableData.write("</div>")
      end if
   genEditableTable = tableStyle & tableData.getvalue
   
end function

'-------------------------------------- chkFunctionOpt ----------------------------------------------
' Gnerate function for check process
' 
'-----------------------------------------------------------------------------------------------

sub chkFunctionOpt(funcname,parameters)

  if not SessionDictionary("checkscreen") then
    paramlist=""
	for each p in parameters
        if typename(p)="String" then
            paramlist = paramlist + """"+p+""","
		elseif typename(p)="Boolean" then
			if p then
				paramlist = paramlist + "oui,"
			else 
				paramlist = paramlist + "non,"
			end if
		elseif typename(p)="Date" then
            paramlist = paramlist + "convertirvers("""+ cstr(p) +""",""date""),"
        else
			paramlist = paramlist + cstr(p) +","
		end if
	next
	functoeval = funcname+"("+mid(paramlist,1,len(paramlist)-1)+")"
	AppendToArray(SessionDictionary("checkscreenData"),functoeval)   
  end if

end sub