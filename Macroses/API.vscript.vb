logger("begin of MEGA API macros yo!!")
'logger("event.myguid:--")
'logger(event.myguid)
logger("event.data:--")
logger(event.data)
args = event.data
logger ("Data > " & tojson(args))

Function IsEmptyArray( massive )
  if typename(massive)="Array" then
    if len(massive) <> 0 then
      isEmptyArray = False
    else
      isEmptyArray = True
    end if
  end if
End Function

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

function Secured(args)
    'args = event.data
    'Set wholePart = CreateObject("Chilkat_9_5_0.Mime")
    'wholePart.AddHeaderField "Description","The WHOLE XML Object"
    'wholePart.ContentType = "text/wholexml"
    set EACAnswer = buffer.create
    'wholePart.SetBody "" &_
    EACAnswer.write("  <WHOLEXML Content=""dynamic"" SessionToken="""" Auth=""internal"">")
    EACAnswer.write ("<EVENTS>")
    EACAnswer.write ("</EVENTS>")
    EACAnswer.write (" <VDOMXML>")
    EACAnswer.write (" <![CDATA[")
    
    
    
     eac_guid = args("additional")("dbkey")
    
      letter = asjson(dbdictionary(eac_guid))
      content = letter("content")
      userdate = Dictionary
      userdate("user_guid") = (Proadmin.currentuser().guid)
      userdate("dt") = Now
      AppendToArray(letter("readStatus"),userdate)
      dbdictionary(eac_guid) = tojson(letter)
      
      
      users_list = tojson(letter("readStatus"))
      res = ""
      
      for each access in letter("readStatus")
          try
              user = Appinmail.getUser(access)
          catch
              user = Dictionary
              user("user_login") = access("user_guid")
          end try
          
          dt = access("dt")
          res = res & dt & "    " & user("user_login") & "<br>"
      next
      
    ' user = Appinmail.currentUser()
    ' right = Appinmail.getAcl(user("guid"), eac_guid)
    
    
    EACAnswer.write ("<CONTAINER name=""container1"" backgroundcolor=""B5F7E9"" designcolor=""84A2F0"" top=""0"" height=""500"" width=""600"" left=""0"">")
    EACAnswer.write ("    <RICHTEXT name=""richtext2"" top=""20"" left=""200"" value=""SECURED mail""/>")
    'EACAnswer.write ("    <FORMBUTTON name=""formbutton1"" left=""270"" top=""100"" label=""OK"" width=""60"" height=""30""/>")
    EACAnswer.write ("    <HYPERTEXT name=""text1"" left=""15"" top=""50"" width=""500"" height=""400"">")
    EACAnswer.write ("        <Attribute Name=""htmlcode"">")
    EACAnswer.write ("            <![CDATA[" & content & "<br><br><br>]]]]><![CDATA[>")
    EACAnswer.write ("        </Attribute>")
    EACAnswer.write ("    </HYPERTEXT>")
    EACAnswer.write ("    <HYPERTEXT name=""text2"" left=""15"" top=""155"" width=""500"" height=""400"" value=""USERS"">")
    EACAnswer.write ("        <Attribute Name=""htmlcode"">")
    EACAnswer.write ("           <![CDATA[users_list:   <pre>" & res & "</pre><br><br><br>]]]]><![CDATA[>")
    EACAnswer.write ("        </Attribute>")
    EACAnswer.write ("    </HYPERTEXT>")
    EACAnswer.write ("</CONTAINER>")
    
    EACAnswer.write (" ]]>")
    
    EACAnswer.write (" </VDOMXML>")
    EACAnswer.write ("</WHOLEXML>")
    logger (EACAnswer.getValue)
    
    Secured = EACAnswer.getValue()
end function


function Antispy(args)

    set EACAnswer = buffer.create
   
    EACAnswer.write("  <WHOLEXML Content=""dynamic"" SessionToken="""" Auth=""internal"">")
    EACAnswer.write ("<EVENTS>")
    EACAnswer.write ("</EVENTS>")
    EACAnswer.write (" <VDOMXML>")
    EACAnswer.write (" <![CDATA[")
    
     eac_guid = args("additional")("dbkey")
     
     content = dbdictionary(eac_guid)
     logger(content)
    
    EACAnswer.write ("<CONTAINER name=""container1"" backgroundcolor=""B5F7E9"" designcolor=""84A2F0"" top=""0"" height=""500"" width=""600"" left=""0"">")
    EACAnswer.write ("    <RICHTEXT name=""richtext2"" top=""20"" left=""200"" value=""ANTYSPY mail""/>")
    'EACAnswer.write ("    <FORMBUTTON name=""formbutton1"" left=""270"" top=""100"" label=""OK"" width=""60"" height=""30""/>")
    EACAnswer.write ("    <HYPERTEXT name=""text1"" left=""15"" top=""50"" width=""500"" height=""400"">")
    EACAnswer.write ("        <Attribute Name=""htmlcode"">")
    EACAnswer.write ("            <![CDATA[" & content & "<br><br><br>]]]]><![CDATA[>")
    EACAnswer.write ("        </Attribute>")
    EACAnswer.write ("    </HYPERTEXT>")
    EACAnswer.write ("</CONTAINER>")
    
    EACAnswer.write (" ]]>")
    EACAnswer.write (" </VDOMXML>")
    EACAnswer.write ("</WHOLEXML>")
    logger (EACAnswer.getValue)
    
    Antispy = EACAnswer.getValue()
    
end function

result = ""

type = args("additional")("type")

if type = "secured" then
    result = Secured(args)
end if

if type = "antispy" then
    result = Antispy(args)
end if

   
response.write(result)
logger("end of API!")