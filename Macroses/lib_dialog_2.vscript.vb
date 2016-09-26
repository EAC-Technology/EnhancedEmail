'Remarks V2.1 // Rev 02-10-2014
'---------------------------------------
'Need searchit.js as a resource to work
'Because of bug in inheritence in VScript all variables of child class must be in mother class
'There is bug when a child call the same method as a function of its own but from his mother class
'the function can not return the 

class component
    
    'general variables
    dim name
    dim properties
    dim propertiesKeyWordsDict
    dim attributes
    dim componentTypeName
    dim datafromrender 'Variable used when a child call render to workaround the bug of calling function from child
    dim rulesCheck
    
    'varaibles for BtnGroup child class
    dim buttonType
    dim btnSize
    dim buttons 
    dim buttonsOrder
    dim buttonsHint
    dim showProgress
    
    'variables for PdfViewer
    dim FileGuid
    dim heightViewer
    
    'variables for Text
    dim center
    
    'variable for TabsGroup
    dim tab
    dim tabsOrder
    
    'Variable for AutoSubmit
    dim NextStep
    
    'Variable Timer
    dim timeout
    
    'Variable Reload
    dim urlGUID
    
    Sub Class_Initialize( )  
        me.componentTypeName = "component"
        me.properties = Dictionary
        me.attributes = Dictionary( "id", "", "visible", "1", "disabled", "0", "fullsize","1"  )
        me.propertiesKeyWordsDict = Dictionary("value","value","defaultvalue","defaultValue")
        me.rulesCheck = Array
    end sub
    
    Function wrappToCData( txt )
        wrappToCData = "<!--[CDATA["  & txt & "]]-->"
    End Function

    sub SetTypeName(name)
      componentTypeName = name
    end sub
    
    sub SetName(name)
      attributes("id")=name
    end sub
    
    sub visible(value)
        if value then
            attributes("visible")="1"
        else
            attributes("visible")="0"
        end if
    end sub
    
    sub disabled(value)
        if value then
            attributes("disabled")="1"
        else
            attributes("disabled")="0"
        end if
    end sub
    
    sub fullsize(value)
        if value then
            attributes("fullsize")="1"
        else
            attributes("fullsize")="0"
        end if
    end sub
    
    sub setAttr(name,value)
      attributes(lcase(name))=value
    end sub
    
    function getAttr(name)
      getAttr=attributes(lcase(name))
    end function
    
    sub setProp(name,value)
      if lcase(name) in propertiesKeyWordsDict then
        properties(propertiesKeyWordsDict(lcase(name)))=value
      else
        properties(lcase(name))=value
      end if
    end sub
    
    function getProp(name)
      if lcase(name) in propertiesKeyWordsDict then
        getProp=propertiesKeyWordsDict(lcase(name))
      else
        getProp=properties(lcase(name))
      end if
    end function
    
    sub addRule(rule)
      AppendToArray(rulesCheck,rule)
    end sub
        
    function render
        
        'generate Attributes
        dim CreateXMLAtributes
        dim beginText
        dim middleText
        dim endText
        
        set CreateXMLAtributes = buffer.create
        for each key in attributes
            CreateXMLAtributes.write( key & "='" & attributes( key ) & "' " )
        next
        
        beginText = "<" & componentTypeName & " " & CreateXMLAtributes.getvalue & "><Properties>"
        'generate properties
        set CreatePropertyXML = buffer.create
        for each key in properties
            if key = "options" then
                CreatePropertyXML.write( "<Property name='" & key & "'>" )
                options = properties( key )
                if lcase(TypeName(options))="dictionary" then
                    for each id in options
                        CreatePropertyXML.write( "<option id='" & id & "'>" & wrappToCData(options( id )) & "</option>" )
                    next
                else
                    for i=0 to len(options)-1
                        CreatePropertyXML.write( "<option id='" & options(i) & "'>" & wrappToCData(options( i )) & "</option>" )
                    next
                end if
                CreatePropertyXML.write( "</Property>" )
            else
                CreatePropertyXML.write( "<Property name='" & key & "'>" & wrappToCData(properties( key )) & "</Property>" )
            end if 
        next
        middleText = CreatePropertyXML.getvalue
      
        endText = "</Properties></" & componentTypeName & ">"
        render = beginText & middleText & endText
        datafromrender = beginText & middleText & endText
    
    end function

end class

class Heading: inherits component

    Sub init
      mybase.SetTypeName("Heading")
      mybase.setProp("label","")
      mybase.setProp("defaultValue","")
    end sub
    
    function check(byref globalContext)
      check=true
    end function
    
end class

class LiveSearch: inherits component

    Sub init
      mybase.SetTypeName("LiveSearch")
      mybase.setProp("label","")
    end sub
    
    sub label(value)
      mybase.setProp("label",value)
    end sub
    
    sub sourceURI(value)
      mybase.setProp("sourceURI",value)
    end sub
    
	sub sourceEvent(value)
      mybase.setProp("sourceEvent",value)
    end sub
	
	sub sourceData(value)
      mybase.setProp("sourceData",value)
    end sub
	
    function check(byref globalContext)
      check=true
    end function
    
end class

class FileUpload: inherits component

    Sub init
      mybase.SetTypeName("FileUpload")
      mybase.setProp("label","")
    end sub
    
    sub label(value)
      mybase.setProp("label",value)
    end sub

    function check(byref globalContext)
      check=true
    end function
    
end class

class RichTextArea: inherits component

    Sub init
      mybase.SetTypeName("RichTextArea")
      mybase.setProp("label","")
      mybase.setProp("defaultValue","")
	  mybase.setProp("height","400")
    end sub
    
    sub label(value)
      mybase.setProp("label",value)
    end sub
    
    sub height(value)
      mybase.setProp("height",value)
    end sub
    
	sub defaultValue(value)
      mybase.setProp("defaultValue",value)
    end sub
	
    function check(byref globalContext)
      check=true
    end function
    
end class

class TextBox: inherits component

    Sub init
      mybase.SetTypeName("TextBox")
      mybase.setProp("label","")
      mybase.setProp("defaultValue","")
    end sub
    
    sub label(value)
      mybase.setProp("label",value)
    end sub
    
    sub defaultValue(value)
      mybase.setProp("defaultValue",value)
    end sub
    
    function check(byref globalContext)
      check=true
    end function
    
end class

class DropDown: inherits component

    Sub init
      mybase.SetTypeName("DropDown")
      mybase.setProp("label","")
      mybase.setProp("defaultValue","")
    end sub
    
    function check(byref globalContext)
      check=true
    end function
    
    sub label(value)
      mybase.setProp("label",value)
    end sub
    
    sub AddOption
    
    end sub
    
    sub removeOption
    
    end sub
    
end class

class CheckBox: inherits component

    Sub init
      mybase.SetTypeName("DropDown")
      mybase.setProp("label","")
      mybase.setProp("defaultValue","")
    end sub
    
    function check(byref globalContext)
      check=true
    end function

    
    sub AddOption
    
    end sub
    
    sub removeOption
    
    end sub
    
end class

class SimpleButton: inherits component

    Sub init
      mybase.SetTypeName("Button")
      mybase.setProp("value","")
    end sub

    function check(byref globalContext)
      check=true
    end function

end class

class Hypertext: inherits component

    Sub init
      mybase.SetTypeName("Hypertext")
      mybase.setProp("value","")
    end sub
    
    sub value(data)
       mybase.setProp("value",data)
    end sub
    
    function check(byref globalContext)
      check=true
    end function

    
end class

class Text: inherits component

    Sub init
      mybase.SetTypeName("Hypertext")
      mybase.setProp("value","")
      center = false
    end sub

    sub setCenter(value)
      center = value
    end sub
    
    sub value(data)
       if center then
         mybase.setProp("value","<Center>" & data & "</Center>")
       else
         mybase.setProp("value",data)
       end if
    end sub
    
    function check(byref globalContext)
      check=true
    end function

end class

class AutoSubmit: inherits component

    Sub init
      mybase.SetTypeName("Hypertext")
      mybase.setProp("value","")
    end sub
    
    sub setNextStep(value)
      nextstep = value
    end sub
    
    function check(byref globalContext)
      check=true
    end function
    
    function render
        set renderedData = buffer.create
        set localrenderData = buffer.create
        localrenderData.write("<script type='text/javascript'>function setOperationStep(step){jQuery('input[name=step]').attr( 'value', step );}</script>")
        localrenderData.write("<script> setOperationStep(""" & nextstep & """);$(""form.xml-dialog-form"").submit();</script>")
        mybase.visible(false)
        mybase.setProp("value",localrenderData.getValue)
        mybase.render
        renderedData.write(datafromrender)
        set myFakeButton = New SimpleButton
            myFakeButton.init
            myFakeButton.visible(false)
            
        renderedData.write(myFakeButton.render)
        render = renderedData.getValue
    end function
    
end class

class CTimer: inherits component

    Sub init
      mybase.SetTypeName("Hypertext")
      mybase.setProp("value","")
    end sub
    
    sub setTimer(value,vtime)
      nextstep = value
      timeout = vtime
    end sub
    
    function check(byref globalContext)
      check=true
    end function
    
    function render
        set renderedData = buffer.create
        set localrenderData = buffer.create
        localrenderData.write("<script type='text/javascript'>function setOperationStep(step){jQuery('input[name=step]').attr( 'value', step );}</script>")
        localrenderData.write("<script type='text/javascript'>var varTime = setTimeout(function(){ setOperationStep(""" & nextstep & """);window.clearTimeout(varTime);$(""form.xml-dialog-form"").submit();}," & timeout & ");</script>")
        mybase.visible(false)
        mybase.setProp("value",localrenderData.getValue)
        mybase.render
        renderedData.write(datafromrender)
        set myFakeButton = New SimpleButton
            myFakeButton.init
            myFakeButton.visible(false)         
        renderedData.write(myFakeButton.render)
        render = renderedData.getValue
        
     end function
    
end class

class Refresh: inherits component

    Sub init
      mybase.SetTypeName("Hypertext")
      mybase.setProp("value","")
    end sub
    
    function check(byref globalContext)
      check=true
    end function

    function render
      set renderedData = buffer.create
      set localrenderData = buffer.create
      if urlGUID="" then
        localrenderData.write("<script type='text/javascript'>location.reload();</script>")
      else
        localrenderData.write("<script type='text/javascript'>window.location.href=""home.vdom#"&urlGUID&""";location.reload();</script>")
      end if
      mybase.visible(false)
      mybase.setProp("value",localrenderData.getValue)
      mybase.render
      renderedData.write(datafromrender)
      set myFakeButton = New SimpleButton
            myFakeButton.init
            myFakeButton.visible(false)         
        renderedData.write(myFakeButton.render)
        render = renderedData.getValue
    end function
    
end class


class PdfViewer: inherits component

    Sub init
      mybase.SetTypeName("Hypertext")
      mybase.setProp("value","")
    end sub
    
    sub setHeight(value)
      heightViewer = value
    end sub
    
    sub setFileGuid(value)
      FileGuid = value
    end sub
    
    function check(byref globalContext)
      check=true
    end function
    
    function render
        set renderedData = buffer.create
        mybase.setProp("value","<iframe src=""/get_pdf?id=" & FileGuid & """ width=""770"" height=""" & heightViewer & """ align=""middle""></iframe> ")
        mybase.render
        renderedData.write(datafromrender)
        render = renderedData.getValue
    end function
    
end class

class BtnGroup: inherits component
    
    Sub init    
      buttonType="text"
      buttons = Dictionary
      buttonsHint = Dictionary
      buttonsOrder = Array
      mybase.SetTypeName("Hypertext")
      mybase.setProp("value","")
      showProgress = false
    end sub
    
    sub setToPicture
        buttonType="picture"
    end sub
    
    sub setToText
        buttonType="text"
    end sub
    
    sub addBtn(btnLegend,StepToGo)
        if instr(lcase(btnLegend),".png")<>0 or instr(lcase(btnLegend),".jpg")<>0 or instr(lcase(btnLegend),".gif")<>0 then
          buttonType="picture"
        else
          buttonType="text"
        end if
        buttons(btnLegend)=StepToGo
        buttonsHint(btnLegend)=""
        AppendToArray(buttonsOrder,btnLegend)
    end sub
    
    sub SetHint(btnLegend,dataHint)
      buttonsHint(btnLegend)=dataHint
    end sub
    
    function check(byref globalContext)
      check=true
    end function
    
    function render
        set renderedData = buffer.create
        set html = buffer.create
        if showProgress then
          if sessiondictionary("ProgressID")<>"" then
            progressIndicatorStr = "showProgress();getProgress('"+sessiondictionary("ProgressID")+"');"
          else
            progressIndicatorStr = ""
          end if
        else
          progressIndicatorStr = ""
        end if
        
        html.write ("<script type='text/javascript'>function setOperationStep(step){jQuery('input[name=step]').attr( 'value', step );"+progressIndicatorStr+"}</script>")
        html.write ("<table width=""100%"" align=""center"" style=""padding:0px; margin:0px;""><tr>")
        host = system.application_hosts
        nbBtn = len(buttonsOrder)
        i=1
        for each elt in buttonsOrder
            nextstep = buttons(elt)
            if buttonType="text" then
                action = "'setOperationStep(""" & nextstep & """)'"
                html.write ("<td align=""center""><input style=""width:" & cstr(btnSize) & "px;"" type=""submit"" name=""" & nextstep & """ value=""" & elt & """ onclick=" & action & "></td>")
            else
                action = "'setOperationStep(""" & nextstep & """);$(""form.xml-dialog-form"").submit();'"
                if elt in buttonsHint then
                  select case i
                    case 1 and nbBtn<>1
                      buttonsHintPos = "right"
                    case i=nbBtn
                      buttonsHintPos = "left"
                    case else
                      buttonsHintPos = "top"
                  end select
                  html.write ("<td align=""center""><div class=""hint--"&buttonsHintPos&" hint--info"" data-hint="""&buttonsHint(elt)&"""><img style='cursor: pointer;' src='//" & host(0) & resources.public_link(elt) & "' name=""" & nextstep & """ onclick=" & action & "></div></td>")
                else
                  html.write ("<td align=""center""><img style='cursor: pointer;' src='//" & host(0) & resources.public_link(elt) & "' name=""" & nextstep & """ onclick=" & action & "></td>")
                end if
            end if
            i=i+1
        next
        html.write ("</tr></table>")
        set myFakeButton = New SimpleButton
            myFakeButton.init
            myFakeButton.visible(false)
            
        renderedData.write(myFakeButton.render)
        mybase.setProp("value",html.getValue)
        mybase.render
        renderedData.write(datafromrender)
        render = renderedData.getValue
        
    end function
    
end class

Function MakeTableRow( Cell1, Cell2 )
    MakeTableRow = "<tr><td class=""rigthCell"">"&Cell1&"</td><td>"&Cell2&"</td></tr>"
End Function

'Tab SpecificComponent
class TabInputText

  dim label
  dim name
  dim value
  dim defaultValue
  dim className
  dim incell
  dim full
  dim isCheckOk
  dim ruleList
  dim errTxt
  dim Hint
  dim isDate
  
  Sub Class_Initialize()
    incell = false
    full = false
    isCheckOk = True
    ruleList = Array
    isDate = false
  end sub
  
  sub addRule(ruleData)
    AppendToArray(ruleList,ruleData)
  end sub
  
  sub setToDate
    isDate=true
  end sub
  
  function check(byref globalContext)
    check=true
    set checkExpToEval = New Evaluator
    set checkExpToEval.EvaluatorContext = globalContext
    
    'set value received
    checkExpToEval.Eval("dim('me')")
    checkExpToEval.Eval("affect('me',"&name&")")
    
    result = ""
    errTxt = ""
    i=0
    for each rule in ruleList
      if TypeName(rule)="dictionary" then
        for each subrule in rule
          result = checkExpToEval.Eval(subrule)
        next
      else
        result = checkExpToEval.Eval(rule)
      end if
      if lcase(result)<>"ok" then
          check = false
          isCheckOk = false
          i=i+1
      logger ("The rule : " & tojson(rule) & " on <b>" & name & "</b> FAILED with result : {"&result&"}")
          errTxt = errTxt & "err " & cstr(i) & ": " & result & ";"
          'showGrowl(result)
      end if
    next
  end function

  function render
      className = SimpleIf(isCheckOk,"", "alert ")
      className = className & SimpleIf(isDate,"date ", "")
    
      if hint="" then
        htmldata = "<div class='inputFullWidth'><input class='inputFullWidth" & SimpleIf(isEmpty(className),"", " " & className) & "' type='text' name='"&name&"' value='" & SimpleIf(isEmpty(value),defaultValue, value ) & "'>" & "<span style='font-family: tahoma; font-size: xx-small; color: crimson;' >" & errTxt & "</span></div>"
      else
        htmldata = "<div class='inputFullWidth hint--top hint--info' data-hint="""&Hint&"""><input class='inputFullWidth" & SimpleIf(isEmpty(className),"", " " & className) & "' type='text' name='"&name&"' value='" & SimpleIf(isEmpty(value),defaultValue, value ) & "'><span style='font-family: tahoma; font-size: xx-small; color: crimson;' >" & errTxt & "</span></div>"
      end if
      if incell then
        render = htmldata
      elseif full then
        render = "<tr><td colspan='2'>" & label & htmldata & "</td></tr>"
      else
        render = MakeTableRow (label, htmldata)
      end if
  end function
  
end class

class TabPicture
  
  dim incell
  dim picname
  dim align
  dim nextstep
  
  Sub Class_Initialize()
    incell = false
    align = "middle"
  end sub
  
  function check(byref globalContext)
    check=true
  end function
  
  function render
    host = system.application_hosts
    action = "'setOperationStep(""" & nextstep & """);$(""form.xml-dialog-form"").submit();'"
    if incell then
      render = "<img " & SimpleIf(nextstep<>"","style='cursor: pointer;'","") & " src='//" & host(0) & resources.public_link(picname) & "'  " & SimpleIf(nextstep<>"","onclick=" & action,"") & " >"
    else
      render = "<tr><td align='" & align & "' colspan='2' ><img " & SimpleIf(nextstep<>"","style='cursor: pointer;'","") & " src='//" & host(0) & resources.public_link(picname) & "' " & SimpleIf(nextstep<>"","onclick=" & action,"") & "  ></td></tr>"
    end if
  end function
end class

class TabText
  
  dim incell
  dim align
  dim value
  dim overflow
  dim height
  
  Sub Class_Initialize()
    incell = false
    align = "middle"
    overflow = false
  end sub
  
  function check(byref globalContext)
    check=true
  end function
  
  function render
    
    if incell then
      render = value
    else
      if overflow then
        styleData="overflow: auto;"
      end if
      if cstr(height)<>"" then
        styleData=styleData & "height: "&cstr(height)&"px;"
      end if
      render = "<tr><td align='" & align & "' colspan='2' ><div style='"&styleData&"'>"&value&"</div></td></tr>"
    end if
    
  end function

end class

class TabPdfViewer

  dim height
  dim incell
  dim fileguid
  dim msg
  dim pdfTemplate
  dim ratio
  
  Sub Class_Initialize()
    incell = false
    height = "500"
    fileguid = ""
    msg = ""
    modeSession = false
    ratio = 0.75
  end sub
  
  function check(byref globalContext)
    check=true
  end function
  
  sub setPageStatus(obj)
    set pageStatus = obj
    modeSession = true
  end sub
  
  function getFileGuid
  
      i=0
      getFileGuid = false
      
      nodes = page_status.selected_nodes
      
      for each node in nodes
        i=i+1
        set nodeSelected = node
      next
    
      Select case i

        case 0
          msg = "Vous devez sélectionner un fichier PDF en premier lieu !"
          if SessionDictionary("fileguid")<>"" then
            msg = ""
            fileguid = SessionDictionary("fileguid")
            getFileGuid = true
          end if
        case 1
          if TypeName(nodeSelected)="Nothing" then
            msg = "Le fichier selectionné n'existe plus ! Veuillez recharger la page"
          elseif nodeSelected.mimetype = "folder" then
            if SessionDictionary("fileguid")<>"" then
              msg = ""
              fileguid = SessionDictionary("fileguid")
              getFileGuid = true
            else
              msg = "Vous avez sélectionné un répertoire et non un fichier de type PDF"
            end if
          elseif nodeSelected.mimetype = "application/pdf" then
            fileguid = nodeSelected.guid
            getFileGuid = true
          else
            msg = "Seul les fichiers PDF sont visibles"
          end if
      case else
        if SessionDictionary("fileguid")<>"" then
            msg = ""
            fileguid = SessionDictionary("fileguid")
            getFileGuid = true 
        else
            msg = "Vous ne pouvez visualiser qu'un seul fichier à la fois"
        end if
      end select
  
  end function
  
  function render
  
    dim datatoshow
    dim nodes
    i=0
    
    if fileguid="" then
    
        nodes = page_status.selected_nodes
   
        for each node in nodes
            i=i+1
            set nodeSelected = node
        next
      
            
        Select case i
            case 0
                if SessionDictionary("fileguid")<>"" then
                    fileguid = SessionDictionary("fileguid")
                    datatoshow = "<iframe src=""/get_pdf?id=" & SessionDictionary("fileguid") & """ width=""100%"" height=""" & height & """ align=""middle""></iframe>"
                else
                    datatoshow = "Vous devez sélectionner un fichier PDF en premier lieu !"
                end if
            case 1
                if nodeSelected.mimetype = "folder" then
                    if SessionDictionary("fileguid")<>"" then
                      datatoshow = "<iframe src=""/get_pdf?id=" & SessionDictionary("fileguid") & """ width=""100%"" height=""" & height & """ align=""middle""></iframe>"
                    else
                      datatoshow = "Vous avez sélectionné un répertoire et non un fichier de type PDF"
                    end if
                elseif nodeSelected.mimetype = "application/pdf" or instr(lcase(nodeSelected.name),".pdf")<>0 then
                    fileguid = nodeSelected.guid
                    'if pdfTemplate then
                    ' datatoshow = "<iframe src=""/get_pdf?id=" & nodeSelected.guid & """ height=""100%"" width=""" & height & """ align=""middle""></iframe>"
                    'else
                      datatoshow = "<iframe src=""/get_pdf?id=" & nodeSelected.guid & """ width=""100%"" height=""" & height & """ align=""middle""></iframe>"
                    'end if  
                else
                     datatoshow = "Seul les fichiers PDF sont visibles"
                end if
           case else
                if SessionDictionary("fileguid")<>"" then
                    fileguid = SessionDictionary("fileguid")
                    datatoshow = "<iframe src=""/get_pdf?id=" & SessionDictionary("fileguid") & """ width=""100%"" height=""" & height & """ align=""middle""></iframe>"
                else
                    datatoshow = "Vous ne pouvez visualiser qu'un seul fichier à la fois"
                end if
                    
        end select
    else
        datatoshow = "Erreur, l'identifiant du fichier est absent. >>> fileguid = " & fileguid
    end if
  
    if incell or pdfTemplate then
        render = datatoshow
    else
        render = "<tr><td align='" & align & "' colspan='2' >"& datatoshow &"</td></tr>"
    end if
  
  end function

end class

function count_folders(folder) 
  result = 0
  for each file in folder.child_nodes
    if file.mimetype = "folder" then
      result = result + 1
    end if
  next
  count_folders = result
end function

Function drill(folder, indent, byref optionsArray,byRef optionsDict,profondeur,depth)
  if profondeur=depth then 
    exit function
  end if
  value = folder.guid
  if value = "" then value = "00000000-0000-0000-0000-000000000000"
  AppendToArray (optionsArray,indent & "--" & folder.name)
  optionsDict (cstr(len(optionsArray)))=folder.guid
  indent = replace(indent, "`", " ")
  nodes = folder.child_nodes
  count = count_folders(folder)
  index = 1
  for each node in nodes 
    if node.mimetype = "folder" then
      if index = count then
        drill(node, indent & "  `", optionsArray, optionsDict,profondeur+1,depth)
      else 
        drill(node, indent & "  |", optionsArray, optionsDict,profondeur+1,depth)
      end if
      index = index + 1
    end if
  next

End Function


class TabDropDown

  dim optionlist
  dim name
  dim label
  dim selectedValue
  dim incell
  dim full
  dim isCheckOk
  dim ruleList
  dim errTxt
  dim Hint
  dim Editable
  dim Size
  dim bSubmit
  dim nextStep
  dim multiple
  dim optionListDict
  dim isOptionList
  
  Sub Class_Initialize()
    optionlist = array
    incell = false
    full = false
    isCheckOk = true
    ruleList = Array
    Size = 1
    bSubmit = false
    nextStep = "Exit"
    multiple = false
    optionListDict = Dictionary
    isOptionList = false
  end sub
  
  sub autoSubmit(nxtStep)
    bSubmit = true
    nextStep = nxtStep
  end sub
  
  sub addRule(ruleData)
    AppendToArray(ruleList,ruleData)
  end sub
  
  function check(byref globalContext)
    check=true
    set checkExpToEval = New Evaluator
    set checkExpToEval.EvaluatorContext = globalContext
    
    'set value received
    checkExpToEval.Eval("dim('me')")
    checkExpToEval.Eval("affect('me',"&name&")")
    
    result = ""
    errTxt = ""
    i=0
    for each rule in ruleList
      if TypeName(rule)="dictionary" then
        for each subrule in rule
          result = checkExpToEval.Eval(subrule)
        next
      else
        result = checkExpToEval.Eval(rule)
      end if
      if lcase(result)<>"ok" then
          check = false
          isCheckOk = false
          i=i+1
      logger ("The rule : " & tojson(rule) & " on <b>" & name & "</b> FAILED with result : {"&result&"}") 
          errTxt = errTxt & "err " & cstr(i) & ": " & result & ";"
          'showGrowl(result)
      end if
    next
  end function
  
  sub addOption(optionName)
    AppendToArray(optionlist,optionname)  
  end sub
  
  sub buildFromPath(pathStart,depth)
    set folder = ProShare.get_by_path(pathStart) 
    drill(folder, "", optionlist,optionListDict,0,depth)
    isOptionList = true
  end sub
  
  sub buildFromUsers(groupName)
    if groupName="" then
      groupName="Everyone"
    end if
    try
      UsersList = ProAdmin.groups(groupName)(0).get_users
    catch
      UsersList = Array
    end try
    For each user in UsersList
      AppendToArray(optionlist,user.name)
      optionListDict (cstr(len(optionlist))) = user.guid
      isOptionList = true
    next
  end sub
  
  sub loadfromdb
    if "_" & name & "_" in dbdictionary then
      optionlist = asjson(dbdictionary("_"&lcase(name)&"_"))
    else
      dbdictionary("_" & name & "_")=tojson(array())
    end if
  end sub
  
  sub loadfrompath(path,type)
    pathOK = true
    try
      set localNode = Proshare.get_by_path(path)
      pathType = localNode.mimetype
    catch
      pathOK = false
    end try
    
    if pathOK and Typename(localNode)<>"Nothing" then
      if pathType = "folder" then
        allChildNodes = localNode.child_nodes
        select case lcase(type)
          case "all"
            for each node in allChildNodes 
              AppendToArray(optionlist,node.name)
            next
          case "folders"
            for each node in allChildNodes
              if node.mimetype = "folder" then
                AppendToArray(optionlist,node.name)
              end if
            next
          case "files"
            for each node in allChildNodes
              if node.mimetype <> "folder" then
                AppendToArray(optionlist,node.name)
              end if
            next
          case else
            AppendToArray(optionlist,"Invalid type")
        end select
     else
       AppendToArray(optionlist,"Path is not a folder")
     end if
   else
     AppendToArray(optionlist,"Invalid path")
   end if
  
  end sub
  
  sub setEditable
    Editable = true
  end sub
  
  sub sort
    'if typename(optionlist)="array" then
    '  if len(optionlist)>0 then
        optionlist = ArraySort(optionlist)
    '  end if
    'end if
  end sub
  
  function render
    
    className = SimpleIf(isCheckOk,"", "alert")
    
    set renderdata = buffer.create
    if bSubmit then
      action = "onclick='setOperationStep(""" & nextstep & """);$(""form.xml-dialog-form"").submit();'"
    else
      action = ""
    end if
    renderdata.write("<div "&action&" class='inputFullWidth" & SimpleIf(Hint="","'"," hint--top hint--info' data-hint="""&Hint&"""") & ">")
    renderdata.write("<select "&SimpleIf(multiple,"multiple","")&" Size='"&cstr(Size)&"' class='inputFullWidth' " & SimpleIf(isEmpty(className),"", " " & className) & "' name='"&name&"'>")
    i=1
    for each elt in optionlist
      if isOptionList then
        renderdata.write("<option value='"&optionListDict(cstr(i))&"' " & SimpleIf(selectedValue=optionListDict(cstr(i)), "selected", "" )&">" & elt & "</option>")
        i=i+1
      else
        renderdata.write("<option value='"&elt&"' " & SimpleIf(selectedValue=elt, "selected", "" )&">" & elt & "</option>")
      end if  
    next
    renderdata.write("</select>" & "<span style='font-family: tahoma; font-size: xx-small; color: crimson;' >" & errTxt & "</span>")
    renderdata.write("</div>")
    codeEditBefore = ""
    codeEditAfter = ""
    if Editable then
      host = system.application_hosts
      args = xml_dialog.get_answer
      renderdata.write ("<script type='text/javascript'>function setOperationStep(step){jQuery('input[name=step]').attr( 'value', step );}</script>")
      'renderdata.write ("<script type='text/javascript'>function setInternalOperationStep(step){jQuery('input[name=internalstep]').attr( 'value', step );}</script>")
      'action = "'setInternalOperationStep(""" & replace(ToJson(Dictionary("step","manageDropDown","data",name,"laststep",mainStep)),"""","""""") & """);setOperationStep(""Internal"");$(""form.xml-dialog-form"").submit();'"
      sessionDictionary("_internalData" & "_" & normalize(name))=Dictionary("step","manageDropDown","data",name,"laststep",mainStep)
      action = "'setOperationStep(""Internal>" & normalize(name) & """);$(""form.xml-dialog-form"").submit();'"
      codeEditBefore = "<div class='hint--top hint--info' data-hint='Editer le champ'><img style='cursor: pointer;' src='//" & host(0) & resources.public_link("crayon.png") & "' onclick=" & action & ">&nbsp;&nbsp;"
      codeEditAfter = "</div>"
    end if
    
    if incell then
      render = codeEditBefore & renderdata.getvalue & codeEditAfter
    elseif full then
      render = "<tr><td colspan='2'>" & codeEditBefore & label & renderdata.getvalue & codeEditAfter & "</td></tr>"
    else
      render = MakeTableRow (codeEditBefore & label & codeEditAfter,renderdata.getvalue)
    end if
    
  end function

end class

class TabFieldSet 

  dim optionlist
  dim name
  dim label
  dim selectedValueList
  dim args
  dim isCheckOk
  dim ruleList
  dim stateList
  dim height
  dim identifier

  Sub Class_Initialize()
    optionlist = array
    selectedValueList = array
    incell = false
    args = xml_dialog.get_answer
    isCheckOk = dictionary
    stateList = dictionary
    ruleList = Array
    height = 130
    identifier = ""
  end sub
  
  sub addOption(optionName)
    AppendToArray(optionlist,optionname)
    isCheckOk(optionName) = true
    stateList(optionName) = false
  end sub
  
  sub setOptionState(optionName,state)
    if state then
        stateList(optionName) = true
    else
        stateList(optionName) = false
    end if
  end sub
  
  sub addRule(ruleData)
    AppendToArray(ruleList)
  end sub
  
  function check(byref globalContext)
    check=true
  end function
  
  sub loadfromdb
    if "_" & name & "_" in dbdictionary then
      optionlist = asjson(dbdictionary("_" & lcase(name) & "_"))
      for each elt in optionlist
        isCheckOk(elt) = true
      next
    else
      dbdictionary("_" & name & "_")=tojson(array())
    end if
  end sub
  
  sub sort
    optionlist = ArraySort(optionlist)
  end sub
 
  function render
    set renderdata = buffer.create
    renderdata.write("<tr><td colspan='2'><fieldset><legend>"&label&"</legend>")
    renderdata.write("<div style='width: 99%; overflow: auto; height: "&cstr(height)&"px;'>")
    for each elt in optionlist
      renderdata.write("<div><input type='checkbox' name='" & SimpleIf(identifier="",normalize(elt),identifier+"_"+normalize(elt)) & "'")
      if normalize(elt) in args then
        if args(normalize(elt))="on" then
          renderdata.write("checked")
        end if
      elseif stateList(elt) then
          renderdata.write("checked")
      end if
      renderdata.write("><span " & SimpleIf(isCheckOk(elt),"", "class='alert'") & ">" & elt & "</span></div>")
    next
    renderdata.write("</div>")
    renderdata.write("</fieldset></td></tr>")
    render = renderdata.getvalue
  end function

end class

class TabTextArea

  dim label
  dim name
  dim value
  dim defaultValue
  dim className
  dim incell
  dim full
  dim isCheckOk
  dim ruleList
  dim errTxt
  dim Hint
  dim row
  dim editor
  
  Sub Class_Initialize()
    incell = false
    full = false
    isCheckOk = true
    ruleList = Array
    row = 3
    editor = false
  end sub

  sub addRule(ruleData)
    AppendToArray(ruleList,ruleData)
  end sub
  
  function check(byref globalContext)
    check=true
    set checkExpToEval = New Evaluator
    set checkExpToEval.EvaluatorContext = globalContext
    
    'set value received
    checkExpToEval.Eval("dim('me')")
    checkExpToEval.Eval("affect('me',"&name&")")
    
    result = ""
    errTxt = ""
    i=0
    for each rule in ruleList
      if TypeName(rule)="dictionary" then
        for each subrule in rule
          result = checkExpToEval.Eval(subrule)
        next
      else
        result = checkExpToEval.Eval(rule)
      end if
      if lcase(result)<>"ok" then
          check = false
          isCheckOk = false
          i=i+1
          logger ("The rule : " & tojson(rule) & " on <b>" & name & "</b> FAILED with result : {"&result&"}") 
          errTxt = errTxt & "err " & cstr(i) & ": " & result & ";"
          'showGrowl(result)
      end if
    next
  end function

  function render
    
    className = SimpleIf(isCheckOk,"", "alert")
    editoroptions = SimpleIf(editor," spellcheck='false' wrap='off'", "")
    host = system.application_hosts
    
    set txtAreaCode = Buffer.create
    'if editor then
    '  txtAreaCode.write("<script src=""//tinymce.cachefly.net/4.0/tinymce.min.js""></script>")
    'end if
    txtAreaCode.write("<div class='inputFullWidth'" & SimpleIf(Hint="","","class='hint--top hint--info' data-hint='"&Hint&"'") & ">")
    txtAreaCode.write("<textarea class='inputFullWidth" & SimpleIf(isEmpty(className),"", " " & className) & "' "+editoroptions+" rows='"&cstr(row)&"' type='text' name='"&name&"'>"&SimpleIf( isEmpty( value ), defaultValue, value )&"</textarea><span style='font-family: tahoma; font-size: xx-small; color: crimson;' >" & errTxt & "</span>")
    txtAreaCode.write("</div>")
    'if editor then
    '  txtAreaCode.write("<script>tinymce.init({selector:'textarea'});</script>")
    'end if
    if incell then
      render = txtAreaCode.getValue
    elseif full then
      render = "<tr><td colspan='2'>"&label&txtAreaCode.getValue&"</td></tr>"
    else
      render = MakeTableRow (label, txtAreaCode.getValue)
    end if
  end function

end class

class TabHistoryList

  dim label
  dim name
  dim defaultValue
  dim className
  dim incell
  dim full
  dim height
  
  Sub Class_Initialize()
    incell = false
    full = true
    height = "500"
  end sub

  
  function check(byref globalContext)
    check=true
  end function
  
  sub addComment(type,msg)
  
    try
      jsonInternalData = asjson(defaultValue)
    catch
      jsonInternalData = Dictionary
    end try
    
    set user = ProAdmin.current_user
    
    newComment         = Dictionary
    newComment("who")  = user.name
    newComment("date") = now
    newComment("msg")  = EscapeStr(msg)
    
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
    guid = md5(cstr(now))
    jsonInternalData(guid)=newComment
    defaultValue = tojson(jsonInternalData)
 
  end sub

  function render
  
    host = system.application_hosts
    try
      jsonInternalData = asjson(defaultValue)
    catch
      jsonInternalData = Dictionary
    end try
    
    set HistoryListCodeCode = Buffer.create
    HistoryListCodeCode.write("<div style='display:none;'><textarea name='"&name&"'>" & defaultValue & "</textarea></div>")
    HistoryListCodeCode.write("<style type=""text/css"">div#wrapdisctable{overflow-y: scroll; width:100%; height:"&cstr(height)&"px}div.msgheader{clear: left;}div.msgcomment{clear: left;}table#disctable td{border-bottom: 1px dotted #dbdbdb;}table#disctable tr:last-child td{border-style: none;}table#disctable tr td{padding-bottom: 15px;}.msgheader{font-size: 12px;color: #999999;}</style>")
    HistoryListCodeCode.write("<script type=""text/javascript"">jQuery(""table#disctable tr"").sortElements(function(a,b){return (new Date($(a).attr(""id"")))<(new Date($(b).attr(""id"")))?1:-1});</script>")
    
    if typename(jsonInternalData)="Dictionary" then
      if len(jsonInternalData)=0 then
        HistoryListCodeCode.write("<div style=""clear:left"">Aucun historique disponible</div>")
      else
        HistoryListCodeCode.write("<div id=""wrapdisctable""><table id=""disctable"" width=""100%"">")
        for each elt in jsonInternalData
          m = jsonInternalData(elt)
          d = CDate( m("date" ) )
          part2 = Hour(d) & ":" & Minute(d) & ":" & Second(d)
          part11 = Month(d) & "/" & Day(d) & "/" & Year(d) & " " & part2
          part12 = Day(d) & "." & Month(d) & "." & Year(d) & " " & part2
          
          HistoryListCodeCode.write("<tr id=""" & part11 & """><td><div class=""msgheader""><img src='//" & host(0) & resources.public_link(m("type")) & "'>" & m("who") & " (" & part12 & ")</div>")
          HistoryListCodeCode.write("<div class=""msgcomment"">" & m("msg") & "</div></td></tr>")
          
        next
        HistoryListCodeCode.write("</table></div>")
      end if
    end if
    
    if incell then
      render = HistoryListCodeCode.getValue
    elseif full then
      render = "<tr><td colspan='2'>"&HistoryListCodeCode.getValue&"</td></tr>"
    else
      render = MakeTableRow ("Historique", HistoryListCodeCode.getValue)
    end if
  end function

end class


class TabButton

  dim size
  dim incell
  dim nextstep
  dim align
  dim label
  dim visible

  Sub Class_Initialize()
    incell    = false
    size      = 70
    nextstep  = "Exit"
    align     = "center"
    label     = "button"
    visible   = true
  end sub
  
  function check(byref globalContext)
    check=true
  end function

  function render
  
    if visible then
      action = "'setOperationStep(""" & nextstep & """)'"
      if incell then
          render = "<input style=""width:" & cstr(Size) & "px;"" type=""submit"" value=""" & label & """ onclick=" & action & ">"
      else
          render = "<tr><td align='" & align & "' colspan='2' ><input style=""width:" & cstr(Size) & "px;"" type=""submit"" value=""" & label & """ onclick=" & action & "><td></tr>"
      end if
    else
      render = ""
    end if
  end function

end class

class TabCell

  dim nbCell
  dim size
  dim align
  dim valign
  
  Sub Class_Initialize()
    size = Array
    align = Array
    valign = Array
  end sub
  
  sub CreateCell(nb)
    nbCell = nb
    for i = 1 to nb
      AppendToArray(size,cstr(cint(100/nb)) & "%")
      AppendToArray(align,"center")
      AppendToArray(valign,"middle")
    next
  end sub
  
end class


'Tab components container
class TabContainer

  dim Component
  dim ComponentsOrder
  dim TabName
  dim args
  dim pdfTemplate
  
  Sub Class_Initialize()
    Component = Dictionary
    ComponentsOrder = Array
    args = xml_dialog.get_answer
    pdfTemplate = false
  end sub
  
  sub addComponent(id,componentType)
    typeFound = true
    
    select case lcase(componentType)
    
      case "inputtext"
        set Component(id)= New TabInputText
        Component(id).name = normalize(id)
        if normalize(id) in args then
          component(id).defaultValue = args(normalize(id))
        end if
        
      case "group"
        set Component(id)= New TabCell
            
      case "picture"
        set Component(id)= New TabPicture
  
      case "text"
        set Component(id)= New TabText
 
      case "pdfviewer"
        set Component(id)= New TabPdfViewer
            Component(id).pdfTemplate = pdfTemplate
      
      case "dropdown"
        set Component(id)= New TabDropDown
        Component(id).name = normalize(id)
        if normalize(id) in args then
          component(id).selectedValue = args(normalize(id))
        end if
        
      case "fieldset"
        set Component(id)= New TabFieldSet
        Component(id).name = normalize(id)
      
      case "textarea"
        set Component(id)= New TabTextArea
        Component(id).name = normalize(id)
        if normalize(id) in args then
          component(id).defaultValue = args(normalize(id))
        end if
        
      case "button"
        set Component(id)= New TabButton
      
      case "history"
        set Component(id)= New TabHistoryList
        Component(id).name = normalize(id)
        if normalize(id) in args then
          component(id).defaultValue = args(normalize(id))
        end if
      
      case else
        typeFound = false
    end select
    
    if typefound then
      appendToArray(ComponentsOrder,id)
    end if
  
  end sub
  
  function check(byref globalContext)
    check=true
    for each elt in ComponentsOrder
      try
        check = CBool(check) and CBool(Component(elt).check(globalContext))
      catch
        logger ("Check Tab Component : " & elt & " FAIL no method Check for class " & TypeName(Component(elt)))
      end try
    next
  end function
  
  function render
  
    dim iscell
    iscell = false
    
    set html = buffer.create
    html.write("<div id="""&normalize(TabName)&""">")
    
    if PdfTemplate then
      html.write("<table id='pdfTemplate' width=""100%"" cellpadding=""0""><tr><td id='pdf' width='{0}' >{pdf}</td><td valign='top'>")
    end if
    html.write("<table width=""100%"" cellpadding=""1""><thead><tr><td width=""30%""></td><td width=""70%""></td></tr></thead><tbody>" )    
    

    'render components
    i=0
    for each elt in ComponentsOrder
      if pdfTemplate and lcase(TypeName(Component(elt)))="tabpdfviewer" then
            memPdfViewerRender = Component(elt).render
            memSize = Component(elt).height
            memRatio = Component(elt).ratio
      else
        if iscell then
          if i=Component(saveelt).nbCell then
            html.write("</tr></table></td></tr>")
            if lcase(TypeName(Component(elt)))="tabcell" then
              saveelt = elt
              i=0
              html.write("<tr><td colspan='2'><table width='100%'><tr>")
            else
              try
                html.write(Component(elt).render)
              catch
                logger ("Oups ! This component can not be rendered : " & elt)
              end try
              iscell = false
              i=0
            end if  
          else
            html.write("<td width='" & Component(saveelt).size(i) & "' align='" & Component(saveelt).align(i) & "' valign='" & Component(saveelt).valign(i) & "'>")
            try
              Component(elt).incell = true
              html.write(Component(elt).render)
            catch
            end try
            html.write("</td>")
            i=i+1
          end if
        else
          if lcase(TypeName(Component(elt)))="tabcell" then
            saveelt = elt
            iscell = true
            i=0
            html.write("<tr><td colspan='2'><table width='100%'><tr>")
          else
            html.write(Component(elt).render)          
          end if
        end if
      end if
    next
    if i<>0 then
      html.write("</tr></table></td></tr>")
    end if
    html.write( "</tbody></table>")
    
    if pdfTemplate then
      html.write("</td></tr></table>")
    end if
    html.write( "</div>" )
    if pdfTemplate then
      datatorender = replace(html.getvalue,"{pdf}",replace(replace(memPdfViewerRender,"<tr>",""),"</tr>",""))
      render = replace(datatorender,"{0}",cstr(cint(memSize*memRatio)))
    else
      render = html.getvalue
    end if

    
  end function
  
end class

class TabGroup: inherits component 

    Sub init   
      tab       = Dictionary
      tabsorder = Array
      mybase.SetTypeName("Hypertext")
      mybase.setProp("value","")
    end sub
    
    sub addTab(TabName)
      set tab(TabName)= New TabContainer
      tab(TabName).TabName=TabName
      AppendToArray(tabsorder,TabName)
    end sub
    
    function check(byref globalContext)
      check=true
      for each elt in tabsorder
        'Try
          check = Cbool(check) and Cbool(tab(elt).check(globalContext))
        'catch
        '  logger ("FAIL to Check TAB("&elt&")")
        'end try
      next
    end function
    
    function render
      set html = buffer.create
      'Tab style'
      html.write( "<style type=""text/css"" media=""screen"">.rigthCell {text-align: right;} .inputFullWidth {width: 100%;} .alert {background-color: rgba(255, 5, 5, 0.2);} div.customTabs background: #333;padding: 1em;}div.customContainer {margin: auto; width: 90%; margin-bottom: 10px;} ul.customTabNavigation {list-style: none; margin: 0; padding: 0;} ul.customTabNavigation li {display: inline;} ul.customTabNavigation li a {background-color: #666;padding: 3px 9px; color: #D1D1D1; text-decoration: none;} ul.customTabNavigation li a.selected, ul.customTabNavigation li a.selected:hover {background: #FFF;color: #000;} ul.customTabNavigation li a:hover {background: #ccc;background: #ccc; color: #000;} ul.customTabNavigation li a:focus {outline: 0;} div.customTabs div {padding: 5px; margin-top: 3px; border: 1px solid #FFF; background: #FFF;}</style>" )
      'Tab script'
      html.write( "<script type=""text/javascript"">$(function(){var tabContainers = $('div.customTabs > div');tabContainers.hide().filter(':first').show();$('div.customTabs ul.customTabNavigation a').click(function(){tabContainers.hide();tabContainers.filter(this.hash).show();$('div.customTabs ul.customTabNavigation a').removeClass('selected');$(this).addClass('selected');return false;}).filter(':first').click();});</script>") 
      html.write ("<script type='text/javascript'>function setMemTab(tabname){jQuery('input[name=tabselected]').attr( 'value', tabname );}</script>")
      html.write( "<div class=""customTabs"">" )
    
      'Tab buttons'
      html.write( "<ul class=""customTabNavigation"">")
      for each elt in tabsorder
        html.write( "<li><a id='a"&normalize(elt)&"' href='#" & normalize(elt) & "' onclick=""setMemTab('" & normalize(elt) & "')"" >" & elt & "</a></li>")
        
      next
      html.write( "</ul>" )
      
      for each elt in tabsorder
        html.write(tab(elt).render)
      next
      args = xml_dialog.get_answer
      
      if args("tabselected")<>"" then
        html.write ("<script type='text/javascript'>")
        for each elt in tabsorder
          if normalize(elt)=args("tabselected") then
            html.write ("jQuery('div[id="&normalize(elt)&"]').attr( 'style', 'display: block;' );")
            html.write ("jQuery('a[id=a"&normalize(elt)&"]').attr( 'class', 'selected' );")
          else
            html.write ("jQuery('div[id="&normalize(elt)&"]').attr( 'style', 'display: none;' );")
            html.write ("jQuery('a[id=a"&normalize(elt)&"]').attr( 'class', '' );")
          end if
        next
        html.write ("setMemTab('"&args("tabselected")&"');")
        html.write ("</script>")
      end if
    
      mybase.setProp("value",html.getValue)
      mybase.render
      render = datafromrender
    
    end function

end class

class XMLScreen

    dim component
    dim compOrder
    dim DropDownSearchable
    dim Title
    dim Width
    dim Height
    dim NextStep
    dim args
    
    Sub Class_Initialize( )
        DropDownSearchable = false
        component = Dictionary
        compOrder = Array
        Title = "Screen" 
        Width = 400
        Height = 150
        NextStep="Exit"
    end
    
    function getNextStep
      getNextStep=NextStep
    end function
    
    sub SetNextStep(NextStepName)
      NextStep=NextStepName
    end sub
    
    sub addComponent(Name,ComponentType)
        typeFound = true
        select case lcase(ComponentType)
            case "heading"
                set component(Name) = New Heading
                    
            case "textbox"
                set component(Name) = New TextBox
                    
            case "textarea"
                set component(Name) = New TextArea
            
            case "pdfviewer"
                set component(Name) = New PdfViewer
                
            case "dropdown"
                set component(Name) = New DropDown
                
            case "checkbox"
                set component(Name) = New CheckBox
            
            case "simplebutton"
                set component(Name) = New SimpleButton
            
            case "hypertext"
                set component(Name) = New Hypertext
                    
            case "text"
                set component(Name) = New Text
            
            case "btngroup"
                set component(Name) = New BtnGroup
                    
            case "tabgroup"
                set component(Name) = New TabGroup
            
            case "autosubmit"
                set component(Name) = New AutoSubmit
            
            case "timer"
                set component(Name) = New CTimer
                
            case "refresh"
                set component(Name) = New Refresh
			
			case "livesearch"
                set component(Name) = New LiveSearch
				
			case "richtextarea"
                set component(Name) = New RichTextArea	

			case "fileupload"
                set component(Name) = New FileUpload

            case else
                typeFound = false
        end select
        
        if typeFound then
            AppendToArray(compOrder,Name)
            component(Name).init
            component(Name).SetName(normalize(name))
            args = xml_dialog.get_answer
            if normalize(name) in args then
              try
                component(Name).defaultValue(args(normalize(name)))
              catch
              end try
            end if
        end if
        
    end sub
    
    sub SetDropDownSearchable
        DropDownSearchable = true
    end sub
    
    sub UnSetDropDownSearchable
        DropDownSearchable = false
    end sub
    
    sub SetVisibility(ArrayOfComponent,value)
        for each elt in ArrayOfComponent
            component(Name).visible(value)
        next
    end sub
    
    sub SetDisability(ArrayOfComponent,value)
        for each elt in ArrayOfComponent
            component(Name).disabled(value)
        next
    end sub
    
    function check(byref globalContext)
      check = true
      for each elt in compOrder
        Try
          check = CBool(check) and CBool(component(elt).check(globalContext))
        catch
          Logger ("FAIL to Check Component :" & elt)
        end try
      next
    end function
    
    function render
        host = system.application_hosts
        'Add Script to make drowndown searchable
        if DropDownSearchable then
            set component("--DropDownSearchable") = New Hypertext
                component("--DropDownSearchable").init
                component("--DropDownSearchable").SetName("--DropDownSearchable")
                AppendToArray(compOrder,"--DropDownSearchable")
                set html = buffer.create
                html.write("<script type=""text/javascript"" src='//" & host(0) & resources.public_link("searchit.js") & "'></script>")
                html.write("<script type=""text/javascript"">$(document).ready(function() {$(""select"").searchit({ textFieldClass: 'searchbox' });});</script>")
                html.write("<style type=""text/css""> body {}.searchbox {border:1px solid #456879;border-radius:6px;height: 22px;width: 100%;margin-top: 5px;}</style>")
                component("--DropDownSearchable").setProp("value",html.getValue)      
        end if
        'Add CSS to manage hint
        
        set renderBuffer = Buffer.create
        for each elt in compOrder
            renderBuffer.write(component(elt).render)
        next
        render = renderBuffer.getValue
    end function
    
    function componentCount
        componentCount = len(compOrder)
    end function
    
end class

Class XMLDialogBuilder

    dim Screen
    dim templateScreen
    dim internalStep
    dim globalContext
    
    Sub Class_Initialize( )
        Screen = Dictionary
        templateScreen = ""
    end
    
    Function wrappToCData( txt )
      wrappToCData = "<!--[CDATA["  & txt & "]]-->"
    End Function
    
    sub AddScreen(ScreenName)
        set Screen(ScreenName)=New XMLScreen
    end sub
    
    sub RemoveScreen(ScreenName)
        remove Screen(ScreenName)
    end sub
    
    function checkScreen(screenName)
        args = xml_dialog.get_answer
        set globalContext = New EvalContext
        for each var in args
          if isnumeric(var) then
            globalContext.setVariable(var,cint(args(var)))
          else
            globalContext.setVariable(var,args(var))
          end if
        next
        Screen(screenName).UnSetDropDownSearchable
        checkScreen = Cbool(Screen(screenName).check(globalContext))
      
    end function
    
    function CompileScreen(screenName)
        CompileScreen = Screen(screenName).render
    end function
    
    sub setCompiledScreen(compiledData)
        templateScreen = compiledData
    end sub
    
    sub SaveScreenToDB(screenName)
      DBDictionary("XMLScreen_" & screenName) = Screen(screenName).render
    end sub
    
    sub RemoveScreenFromDB(screenName)
      DBDictionary.remove("XMLScreen_" & screenName)
    end sub
    
    sub loadTemplate(screenName)
      templateScreen = DBDictionary("XMLScreen_" & screenName)
    end sub
    
    sub runInternalProcess
      args = xml_dialog.get_answer
      dim Title
      set xml_components = Buffer.create
      xml_footer = "</Components></VDOMFormContainer>"
      endOfInternalProcess = false
      EndMsg = ""
      
      if instr(args("step"),">")<>0 then
        mainStep = split(args("step"),">")(0)
        objSender = split(args("step"),">")(1)
        sessionDictionary("objSender") = objSender
      else
        mainStep = args("step")
        objSender = sessionDictionary("objSender")
      end if

      'logger ("_internalData" & "_" & objSender & " -- " &  TypeName(sessionDictionary("_internalData" & "_" & objSender)))
      if TypeName(sessionDictionary("_internalData" & "_" & objSender))="Dictionary" then
        set internalScreen = New XMLScreen
        select case sessionDictionary("_internalData" & "_" & objSender)("step")
          case "manageDropDown"
            if mainStep="Internal" then
              'Initialization of the Internal process
              sessionDictionary("_memLastScreenData")= args
              Title = "Gestion d'un champ de formulaire"
              internalScreen.Width = 400 
              internalScreen.Height = 170
              internalScreen.addComponent("msg","Hypertext")
              internalScreen.component("msg").setProp("Value","Gestion du champ : <b>" & sessionDictionary("_internalData" & "_" & objSender)("data") & "</b>.<br>Veuillez choisir une option :")
              internalScreen.addComponent("btns","btngroup")
              internalScreen.Component("btns").addBtn("Ajouter","Internal:AddValue")
              internalScreen.Component("btns").addBtn("Modifier","Internal:UpdateValue")
              internalScreen.Component("btns").addBtn("Supprimer","Internal:DeleteValue")
              internalScreen.Component("btns").addBtn("Retour","Internal:End")
            else
              subprocess = split(mainStep,":")
              select case subprocess(1)
                case "mainScreen"
                  Title = "Gestion d'un champ de formulaire"
                  internalScreen.Width = 400 
                  internalScreen.Height = 170
                  internalScreen.addComponent("msg","Hypertext")
                  internalScreen.component("msg").setProp("Value","Gestion du champ : <b>" & sessionDictionary("_internalData" & "_" & objSender)("data") & "</b>.<br>Veuillez choisir une option :")
                  internalScreen.addComponent("btns","btngroup")
                  internalScreen.Component("btns").addBtn("Ajouter","Internal:AddValue")
                  internalScreen.Component("btns").addBtn("Modifier","Internal:UpdateValue")
                  internalScreen.Component("btns").addBtn("Supprimer","Internal:DeleteValue")
                  internalScreen.Component("btns").addBtn("Retour","Internal:End")
                case "AddValue"
                  Title = "Ajouter une valeur"
                  internalScreen.Width = 400 
                  internalScreen.Height = 170
                  internalScreen.addComponent("newvalue","TextBox")
                  internalScreen.Component("newvalue").label("Nouvelle valeur :")
                  internalScreen.addComponent("btns","btngroup")
                  internalScreen.Component("btns").addBtn("Valider","Internal:SaveValue")
                  internalScreen.Component("btns").addBtn("Annuler","Internal:mainScreen")
                  
                case "SaveValue"
                  tmpValue = asjson(DbDictionary("_" & sessionDictionary("_internalData" & "_" & objSender)("data") & "_"))
                  AppendToArray(tmpValue,args("newvalue"))
                  DbDictionary("_" & sessionDictionary("_internalData" & "_" & objSender)("data") & "_")=toJson(tmpValue)
                  Title = "Mémorisation"
                  internalScreen.Width = 400 
                  internalScreen.Height = 150
                  internalScreen.addComponent("msg","Hypertext")
                  internalScreen.component("msg").setProp("Value","La valeur <b>" & args("newvalue") & "</b> a bien été ajouté au champ <b>" & sessionDictionary("_internalData" & "_" & objSender)("data") & "</b>.")                  
                  internalScreen.addComponent("btns","btngroup")
                  internalScreen.Component("btns").addBtn("Ajouter une autre","Internal:AddValue")
                  internalScreen.Component("btns").addBtn("Retour au menu","Internal:mainScreen")
                  internalScreen.Component("btns").addBtn("Terminer","Internal:End")
                  
                case "UpdateValue"
                  Title = "Modifier une valeur"
                  internalScreen.Width = 400 
                  internalScreen.Height = 250
                  internalScreen.addComponent("oldvalue","DropDown")
                  internalScreen.Component("oldvalue").label("La valeur :")
                  internalScreen.Component("oldvalue").setProp("options",ArraySort(asjson(DbDictionary("_" & sessionDictionary("_internalData" & "_" & objSender)("data") & "_"))))
                  internalScreen.addComponent("newvalue","TextBox")
                  internalScreen.Component("newvalue").label("Sera remplacée par :")
                  internalScreen.addComponent("btns","btngroup")
                  internalScreen.Component("btns").addBtn("Valider","Internal:SaveUpdate")
                  internalScreen.Component("btns").addBtn("Annuler","Internal:mainScreen")
                  
                case "SaveUpdate"
                  tmpArray = Array
                  for each elt in asjson(DbDictionary("_" & sessionDictionary("_internalData" & "_" & objSender)("data") & "_"))
                    if elt <> args("oldvalue") then
                      AppendToArray(tmpArray,elt)
                    else
                      AppendToArray(tmpArray,args("newvalue"))
                    end if
                  next
                  DbDictionary("_" & sessionDictionary("_internalData" & "_" & objSender)("data") & "_")=tojson(tmpArray)
                  
                  Title = "Mémorisation"
                  internalScreen.Width = 400 
                  internalScreen.Height = 180
                  internalScreen.addComponent("msg","Hypertext")
                  internalScreen.component("msg").setProp("Value","La valeur <b>" & args("oldvalue") & "</b> a bien été remplacée par la valeur <b>" & args("newvalue") & "</b>.")                  
                  internalScreen.addComponent("btns","btngroup")
                  internalScreen.Component("btns").addBtn("Retour au menu","Internal:mainScreen")
                  internalScreen.Component("btns").addBtn("Terminer","Internal:End")
                
                case "DeleteValue"
                  Title = "Supprimer une valeur"
                  internalScreen.Width = 400 
                  internalScreen.Height = 200
                  internalScreen.addComponent("thevalue","DropDown")
                  internalScreen.Component("thevalue").label("La valeur :")
                  internalScreen.Component("thevalue").setProp("options",ArraySort(asjson(DbDictionary("_" & sessionDictionary("_internalData" & "_" & objSender)("data") & "_"))))
                  internalScreen.addComponent("msg","Hypertext")
                  internalScreen.component("msg").setProp("Value","Va être supprimée.")
                  internalScreen.addComponent("btns","btngroup")
                  internalScreen.Component("btns").addBtn("Valider","Internal:SaveDelete")
                  internalScreen.Component("btns").addBtn("Annuler","Internal:mainScreen")
                  
                case "SaveDelete"
                  tmpArray = Array
                  for each elt in asjson(DbDictionary("_" & sessionDictionary("_internalData" & "_" & objSender)("data") & "_"))
                    if elt <> args("thevalue") then
                      AppendToArray(tmpArray,elt)
                    end if
                  next
                  DbDictionary("_" & sessionDictionary("_internalData" & "_" & objSender)("data") & "_")=tojson(tmpArray)
                  
                  Title = "Mémorisation"
                  internalScreen.Width = 400 
                  internalScreen.Height = 150
                  internalScreen.addComponent("msg","Hypertext")
                  internalScreen.component("msg").setProp("Value","La valeur <b>" & args("thevalue") & "</b> a bien été supprimée")                  
                  internalScreen.addComponent("btns","btngroup")
                  internalScreen.Component("btns").addBtn("Retour au menu","Internal:mainScreen")
                  internalScreen.Component("btns").addBtn("Terminer","Internal:End")
                  
                
                case "End"
                  EndMsg = "<center><br><br>Fin du traitement, vous allez retourner au formulaire.</center>"
                  endOfInternalProcess = true
             
                case else
                
              end select
            end if
          case else
          
        end select
        
        if endOfInternalProcess then
          Title ="Retour au formulaire"
          internalScreen.Height = 150
          internalScreen.addComponent("msg","Hypertext")
          internalScreen.component("msg").setProp("Value",EndMsg)
          for each key in sessionDictionary("_memLastScreenData")
            if (key <> "macros_id") and (key <> "macros_id") and (key <> "id") and (key <> "sender") then
              internalScreen.addComponent(key,"TextBox")
              internalScreen.component(key).visible(false)
              internalScreen.component(key).setProp("defaultValue",replace(sessionDictionary("_memLastScreenData")(key),"""","&quot;"))
            end if
          next
          internalScreen.addComponent("go","autosubmit")
          internalScreen.Component("go").SetNextStep(sessionDictionary("_internalData" & "_" & objSender)("laststep"))
        else
          
        end if
        
        'Adding some additionnal textbox to store the Step
        internalScreen.addComponent("step","TextBox")
        internalScreen.component("step").visible(false)
        
        'Adding some additionnal textbox to store MacroID value
        internalScreen.addComponent("macros_id","TextBox")
        internalScreen.component("macros_id").visible(false)
        internalScreen.component("macros_id").setProp("defaultValue",xml_dialog.get_macros_id())
        xml_dialog.set_size( internalScreen.Width,internalScreen.Height)
        xml_components.write(internalScreen.render)
      else
        set errScreen = New XMLScreen
        
        title = "Can not get internal sub process !"
        errScreen.Height = 230
        
        'Create the message to show
        errScreen.addComponent("msg","Hypertext")
        errScreen.component("msg").setProp("Value","Impossible to find the variable internalstep to run any internal process ! <br>Something is going wrong ! Check the step in your main macro.")
        
        errScreen.addComponent("btn","SimpleButton")
        errScreen.component("btn").setProp("value","OK")
        
        'Adding some additionnal textbox to store MacroID value
        errScreen.addComponent("macros_id","TextBox")
        errScreen.component("macros_id").visible(false)
        errScreen.component("macros_id").setProp("defaultValue",xml_dialog.get_macros_id())
        
        xml_dialog.set_size( errScreen.Width,errScreen.Height)
        xml_components.write(errScreen.render)
      
      end if
      
      xml_header = "<?xml version='1.0' encoding='utf-8'?><VDOMFormContainer><Properties><Property name='label'>" & wrappToCData(title) & "</Property></Properties><Components>" 
      xml_dialog.show_xml_form ( xml_header & xml_components.getvalue &  xml_footer  )
    
    end sub
    
    sub ProShareRefresh(itempos)
      xml_dialog.executeCallback("selectEltOfPath", "{ itemid : " & cstr(itempos) & "}")
    end sub
    
    sub showScreen(screenName)
    
      dim Title
      set xml_components = Buffer.create
      xml_footer = "</Components></VDOMFormContainer>"
 
      if screenName in Screen then
      
        host = system.application_hosts
        title = Screen(screenName).Title
        args = xml_dialog.get_answer
        
        'Adding some additionnal textbox to store the selected Tab
        Screen(screenName).addComponent("tabselected","TextBox")
        Screen(screenName).component("tabselected").visible(false)
        Screen(screenName).component("tabselected").setProp("defaultValue",args("tabselected"))
        
        'Adding some additionnal textbox to store Next Step
        Screen(screenName).addComponent("step","TextBox")
        Screen(screenName).component("step").visible(false)
        Screen(screenName).component("step").setProp("defaultValue",Screen(screenName).getNextStep)
        
        'Adding some additionnal textbox to store MacroID value        
        Screen(screenName).addComponent("macros_id","TextBox")
        Screen(screenName).component("macros_id").visible(false)
        Screen(screenName).component("macros_id").setProp("defaultValue",xml_dialog.get_macros_id())
        
        'Adding css for tooltip
        Screen(screenName).addComponent("--HINT_CSS","Hypertext")
        Screen(screenName).component("--HINT_CSS").visible(false)
        Screen(screenName).component("--HINT_CSS").setProp("Value","<style type=""text/css"" media=""screen"">.hint,[data-hint]{position:relative;display:inline-block}.hint:before,.hint:after,[data-hint]:before,[data-hint]:after{position:absolute;-webkit-transform:translate3d(0,0,0);-moz-transform:translate3d(0,0,0);transform:translate3d(0,0,0);visibility:hidden;opacity:0;z-index:1000000;pointer-events:none;-webkit-transition:.3s ease;-moz-transition:.3s ease;transition:.3s ease}.hint:hover:before,.hint:hover:after,.hint:focus:before,.hint:focus:after,[data-hint]:hover:before,[data-hint]:hover:after,[data-hint]:focus:before,[data-hint]:focus:after{visibility:visible;opacity:1}.hint:before,[data-hint]:before{content:'';position:absolute;background:transparent;border:6px solid transparent;z-index:1000001}.hint:after,[data-hint]:after{content:attr(data-hint);background:#383838;color:#fff;text-shadow:0 -1px 0 #000;padding:8px 10px;font-size:12px;line-height:12px;white-space:nowrap;box-shadow:4px 4px 8px rgba(0,0,0,.3)}.hint--top:before{border-top-color:#383838}.hint--bottom:before{border-bottom-color:#383838}.hint--left:before{border-left-color:#383838}.hint--right:before{border-right-color:#383838}.hint--top:before{margin-bottom:-12px}.hint--top:after{margin-left:-18px}.hint--top:before,.hint--top:after{bottom:100%;left:50%}.hint--top:hover:after,.hint--top:hover:before,.hint--top:focus:after,.hint--top:focus:before{-webkit-transform:translateY(-8px);-moz-transform:translateY(-8px);transform:translateY(-8px)}.hint--bottom:before{margin-top:-12px}.hint--bottom:after{margin-left:-18px}.hint--bottom:before,.hint--bottom:after{top:100%;left:50%}.hint--bottom:hover:after,.hint--bottom:hover:before,.hint--bottom:focus:after,.hint--bottom:focus:before{-webkit-transform:translateY(8px);-moz-transform:translateY(8px);transform:translateY(8px)}.hint--right:before{margin-left:-12px;margin-bottom:-6px}.hint--right:after{margin-bottom:-14px}.hint--right:before,.hint--right:after{left:100%;bottom:50%}.hint--right:hover:after,.hint--right:hover:before,.hint--right:focus:after,.hint--right:focus:before{-webkit-transform:translateX(8px);-moz-transform:translateX(8px);transform:translateX(8px)}.hint--left:before{margin-right:-12px;margin-bottom:-6px}.hint--left:after{margin-bottom:-14px}.hint--left:before,.hint--left:after{right:100%;bottom:50%}.hint--left:hover:after,.hint--left:hover:before,.hint--left:focus:after,.hint--left:focus:before{-webkit-transform:translateX(-8px);-moz-transform:translateX(-8px);transform:translateX(-8px)}.hint--error:after{background-color:#b34e4d;text-shadow:0 -1px 0 #592726}.hint--error.hint--top:before{border-top-color:#b34e4d}.hint--error.hint--bottom:before{border-bottom-color:#b34e4d}.hint--error.hint--left:before{border-left-color:#b34e4d}.hint--error.hint--right:before{border-right-color:#b34e4d}.hint--warning:after{background-color:#c09854;text-shadow:0 -1px 0 #6c5328}.hint--warning.hint--top:before{border-top-color:#c09854}.hint--warning.hint--bottom:before{border-bottom-color:#c09854}.hint--warning.hint--left:before{border-left-color:#c09854}.hint--warning.hint--right:before{border-right-color:#c09854}.hint--info:after{background-color:#3986ac;text-shadow:0 -1px 0 #193b4d}.hint--info.hint--top:before{border-top-color:#3986ac}.hint--info.hint--bottom:before{border-bottom-color:#3986ac}.hint--info.hint--left:before{border-left-color:#3986ac}.hint--info.hint--right:before{border-right-color:#3986ac}.hint--success:after{background-color:#458746;text-shadow:0 -1px 0 #1a321a}.hint--success.hint--top:before{border-top-color:#458746}.hint--success.hint--bottom:before{border-bottom-color:#458746}.hint--success.hint--left:before{border-left-color:#458746}.hint--success.hint--right:before{border-right-color:#458746}.hint--always:after,.hint--always:before{opacity:1;visibility:visible}.hint--always.hint--top:after,.hint--always.hint--top:before{-webkit-transform:translateY(-8px);-moz-transform:translateY(-8px);transform:translateY(-8px)}.hint--always.hint--bottom:after,.hint--always.hint--bottom:before{-webkit-transform:translateY(8px);-moz-transform:translateY(8px);transform:translateY(8px)}.hint--always.hint--left:after,.hint--always.hint--left:before{-webkit-transform:translateX(-8px);-moz-transform:translateX(-8px);transform:translateX(-8px)}.hint--always.hint--right:after,.hint--always.hint--right:before{-webkit-transform:translateX(8px);-moz-transform:translateX(8px);transform:translateX(8px)}.hint--rounded:after{border-radius:4px}.hint--bounce:before,.hint--bounce:after{-webkit-transition:opacity .3s ease,visibility .3s ease,-webkit-transform .3s cubic-bezier(0.71,1.7,.77,1.24);-moz-transition:opacity .3s ease,visibility .3s ease,-moz-transform .3s cubic-bezier(0.71,1.7,.77,1.24);transition:opacity .3s ease,visibility .3s ease,transform .3s cubic-bezier(0.71,1.7,.77,1.24)}</style>")

        'Adding progessbar
        Screen(screenName).addComponent("progressbar","Hypertext")
        Screen(screenName).component("progressbar").visible(false)
        'topLoaderSize = cstr((cint(Screen(screenName).height)/3.2))
        if Screen(screenName).height < 170 then
          topLoaderSize = cstr(Screen(screenName).height - 90)          
        else
          topLoaderSize = "100"
        end if
        'posLeft = cstr(cint((Screen(screenName).width - cint((Screen(screenName).height)/3.2))/2))
        posLeft = cstr(Screen(screenName).width - 125)
        'posTop = cstr(cint((Screen(screenName).height - 50 - cint((Screen(screenName).height)/3.2))/2))
        posTop = "28"
  
        set html = buffer.create
        logger tojson(host)
        html.write "<script src='//" & host(0) & resources.public_link("jquery.percentageloader-0.1.min.js") & "'></script>"
        html.write "<style type=""text/css"" media=""screen"" >" & resources.open("jquery.percentageloader-0.1.css").getvalue & "</style>"
        html.write "<script>"
        html.write "$('#o_" & Replace(xml_dialog.containerGUID,"-","_") & "').append( ""<div id='topLoader' style='position: absolute; top: "+posTop+"px; left: "+posLeft+"px;'></div>"" );"
        html.write "var $topLoader = $('#topLoader').percentageLoader({width: "+topLoaderSize+", height: "+topLoaderSize+" , controllable : false, progress : 0.0, onProgressUpdate : function(val) {"
        'html.write "      $topLoader.setValue(Math.round(val * 100.0));"
        html.write "    }});"
        'html.write "var topLoaderRunning = false;"
        'html.write "topLoaderRunning = true;"
        html.write "$topLoader.setProgress(0);"
        html.write "$topLoader.setValue('...');"
        html.write "$('#topLoader').css('visibility','hidden')"
        
        html.write "</script>"
        
        Screen(screenName).component("progressbar").setProp("Value",html.getValue)

        'Adding some script for e2VDOM management for tooltip
        Screen(screenName).addComponent("E2VDOM","Hypertext")
        Screen(screenName).component("E2VDOM").visible(false)
        set html = buffer.create
        html.write "<script type='text/javascript'>"
        html.write "e2vdomSV['DialogGUID'] = '#o_" & Replace(xml_dialog.containerGUID,"-","_") & "';"
        html.write "e2vdomSV['MacrosID'] = '" & xml_dialog.get_macros_id() & "';"
        
        'set callbackfunctions
        html.write resources.open("dialog_v2_callback.js").getvalue
        html.write "</script>"
        Screen(screenName).component("E2VDOM").setProp("Value",html.getValue)
        
        xml_dialog.set_size( Screen(screenName).Width,Screen(screenName).Height)
        xml_components.write(Screen(screenName).render)
      else
        set errScreen = New XMLScreen
        
        title = "No Screen found !"
        
        'Create the message to show
        errScreen.addComponent("msg","Hypertext")
        errScreen.component("msg").setProp("Value","Impossible to find the screen (<b>" & screenName & "</b>) in the list of screen created !")
        
        'Adding some additionnal textbox to store Next Step
        errScreen.addComponent("step","TextBox")
        errScreen.component("step").visible(false)
        errScreen.component("step").setProp("defaultValue",errScreen.getNextStep)
        
        'Adding some additionnal textbox to store MacroID value
        errScreen.addComponent("macros_id","TextBox")
        errScreen.component("macros_id").visible(false)
        errScreen.component("macros_id").setProp("defaultValue",xml_dialog.get_macros_id())
        
        xml_dialog.set_size( errScreen.Width,errScreen.Height)
        xml_components.write(errScreen.render)
      end if
      
      xml_header = "<?xml version='1.0' encoding='utf-8'?><VDOMFormContainer><Properties><Property name='label'>" & wrappToCData(title) & "</Property></Properties><Components>" 
      xml_dialog.show_xml_form ( xml_header & xml_components.getvalue & templateScreen &  xml_footer  )

    end sub
    
end class