class EACOLD

  Dim Email                      'The Data concerning the email
  Dim renderingMode              'Rendering mode binary/string
  Dim dynamic                    'is dynamic content true/false
  Dim Auth                       'authentication mode internal/external 
  Dim logincontainer             'The guid of the container where the login method is
  Dim loginmethod                'The name of method to log in
  Dim getcontainer               'The container where to call the method to get data
  Dim getmethod                  'The name of the method to get the data
  Dim postcontainer              'The container where to call the method to put data
  Dim postmethod                 'The name of the method to put the data
  Dim sessionToken               'sessionToken a unique identifier to specify the session
  Dim API_Server                 'The business logic serveur url
  Dim appID                      'The GUID of the VDOM App running the business logic
  Dim gdata                      'Data to be sent as a pattern to the serveur
  Dim pdata                      'Data to be sent as a pattern to the serveur
  Dim EventsData                 'The Json Structure that represent the events of the application
  Dim EACAppName                 'A friendly name for this EAC app
  Dim EACToken                   'A unique identifier for this EAC App
  Dim EACMethod                  'The method to use new/update/delete
  
  sub class_initialize
      
      Email                = Dictionary("subject","EAC email ...","sender","EAC ProShare Automation","recipients",Array)
      renderingMode        = "binary"
      dynamic              = true
      sessionToken         = ""
      Auth                 = "internal"
      appID                = "7f459762-e1ba-42d3-a0e1-e74beda2eb85"
      logincontainer       = "5073ff75-da99-44fb-a5d7-e44e5ab28598"
      loginmethod          = "login"
      getcontainer         = "5073ff75-da99-44fb-a5d7-e44e5ab28598"
      getmethod            = "call_macro"
      postcontainer        = "5073ff75-da99-44fb-a5d7-e44e5ab28598"
      postmethod           = "call_macro"
      EventsData           = ""
      EACMethod            = "new"
      gd = Dictionary
      gd("plugin_guid") = "5d525ea5-bb1d-4eab-b5c3-77a30b5ebe67"
      gd("async")       = false
      gd("name")        = "loaddata"
      gd("data")        = "myguid"
      gdata = tojson(gd)
      pd = Dictionary
      pd("plugin_guid") = "5d525ea5-bb1d-4eab-b5c3-77a30b5ebe67"
      pd("async")       = false
      pd("name")        = "loaddata"
      pd("data")        = ""
      pdata = tojson(pd)
    
                 
  end sub
  
  function addTo(name,emailuser)
    array_len = len(Email("recipients"))
    if array_len=0 then
      Email("recipients")=Array(name+" "+"<"+emailuser+">")
    else
      recipients = Email("recipients")
      Redim Preserve recipients(array_len+1)
      recipients(array_len+1) = name+" "+"<"+emailuser+">"
      Email("recipients") = recipients
    end if
  end function
  
  sub loadApp(appName)
  
    if DbDictionary("EAC_App_"+normalize(appName))<>"" then
      AppData = asjson(DbDictionary("EAC_App_"+normalize(appName)))
    end if
  
  end sub

  function newSession()
    sessionToken = generateguid
  end function

  function openSession(guid)
    sessionToken = guid
  end function
  
  sub getdata(data)
    if getmethod = "call_macro" then
      gd = Dictionary
      gd("plugin_guid") = "95533a74-bbe3-4962-9405-9aa96a4092cd"
      gd("async")       = false
      gd("name")        = "EAC_Call"
      gd("data")        = data
      gdata = tojson(gd)
    else
      gdata = tojson(data)
    end if
  end sub
  
  sub postdata(data)
    if getmethod = "call_macro" then
      pd = Dictionary
      pd("plugin_guid") = "95533a74-bbe3-4962-9405-9aa96a4092cd"
      pd("async")       = false
      pd("name")        = "EAC_Call"
      pd("data")        = data
      pdata = tojson(pd)
    else
      pdata = tojson(pd)
    end if
  end sub
  
  sub SetEvents(EvtData)
    if typename(EvtData)="Dictionary" then
      EventsData = EvtData
    else
      try
        EventsData = asjson(EvtData)
      catch
      end try
    end if
  end sub
  
  function wholexml(vdomxml)
  
      if dynamic then
          dmethod = "dynamic"
      else
          dmethod = "static"
      end if
      
      hosts = system.application_hosts
      set WHOLE_DATA = buffer.create
      WHOLE_DATA.write("<WHOLEXML Content="""+dmethod+""" SessionToken="""+sessionToken+""" Auth="""+Auth+""">")
      if Auth="internal" then
        WHOLE_DATA.write("	<API server="""+"127.0.0.1"" appID="""+appID+""">")
      else
        WHOLE_DATA.write("	<API server="""+"http://"+hosts(0)+""" appID="""+appID+""">")
      end if
      WHOLE_DATA.write("	  <LOGIN container="""+logincontainer+""" action="""+loginmethod+"""/>")
      WHOLE_DATA.write("	  <GET container="""+getcontainer+""" action="""+getmethod+""">")
      WHOLE_DATA.write("	    <PATTERN><![CDATA[")
      WHOLE_DATA.write(gdata)
      WHOLE_DATA.write("		  ]]>")
      WHOLE_DATA.write("	    </PATTERN>")
      WHOLE_DATA.write("	  </GET>")
      WHOLE_DATA.write("	  <POST container="""+postcontainer+""" action="""+postmethod+""">")
      WHOLE_DATA.write("	    <PATTERN><![CDATA[")
      WHOLE_DATA.write(pdata)
      WHOLE_DATA.write("		  ]]>")
      WHOLE_DATA.write("	   </PATTERN>")
      WHOLE_DATA.write("	  </POST>")
      WHOLE_DATA.write("	</API>")
      if typename(EventsData)="Dictionary" then
        WHOLE_DATA.write("	<EVENTS>")
        WHOLE_DATA.write("		<![CDATA[")
        WHOLE_DATA.write(tojson(EventsData))
        WHOLE_DATA.write("		]]>")
        WHOLE_DATA.write("	</EVENTS>")
      end if
      WHOLE_DATA.write("	<VDOMXML>")
      WHOLE_DATA.write("		<![CDATA[")
      WHOLE_DATA.write(replace(vdomxml,"]]>","]]]]><![CDATA[>"))
      WHOLE_DATA.write("		]]>")
      WHOLE_DATA.write("	</VDOMXML>")
      WHOLE_DATA.write("</WHOLEXML>")
      if renderingMode = "binary" then
          wholexml = WHOLE_DATA.getbinary
      else
          wholexml = WHOLE_DATA.getvalue
      end if
  end function
 
  function EACAnswer(vdomxml)
  
    Set EACData = Buffer.create
    
    EACData.write ("<WHOLEXML SessionToken="""+sessionToken+""">")
    if typename(EventsData)="Dictionary" then
     EACData.write("	<EVENTS>")
     EACData.write("		<![CDATA[")
     EACData.write(tojson(EventsData))
     EACData.write("		]]>")
     EACData.write("	</EVENTS>")
    else
     EACData.write("	<EVENTS></EVENTS>")
    end if
    EACData.write ("  <VDOMXML>")
    EACData.write ("	<![CDATA[")
    EACData.write(replace(vdomxml,"]]>","]]]]><![CDATA[>"))
    EACData.write("		]]>")
    EACData.write("	</VDOMXML>")
    EACData.write("</WHOLEXML>")

    EACAnswer = EACData.getValue
    
    'logger ("VDOMxml >" & vdomxml)
  
  end function
  
   '""""""""""""""""""""'""""""sendEACMessage"""""""""""""""""""""""""""""""""""
   
   
  function sendEACMessage(methodname,vdomxml,sett)
     
     if DbDictionary("EAC_App_"+normalize(EACAppName))<>"" then
       EACToken = DbDictionary("EAC_App_"+normalize(EACAppName))
     else
       EACToken = generateguid
       DbDictionary("EAC_App_"+normalize(EACAppName)) = EACToken
     end if
      select case methodname
        case "new","update"
          EACMethod = methodname
        case "delete"
          EACMethod = methodname
        case else
          EACMethod = "new"
          erase DbDictionary("EAC_App_"+normalize(EACAppName))
      end select
  
     Set myEACMessage = new mailmessage
         myEACMessage.addheader("EAC-Token",EACToken)
         myEACMessage.addheader("EAC-Method",EACMethod)
         myEACMessage.subject    = Email("subject")
         'myEACMessage.sender     = Email("sender")
         
          myEACMessage.sender = sett.user
        myEACMessage.recipients = "testtestalina@gmail.com"
         myEACMessage.body       = "This message is EAC one !"
         logger(myEACMessage.recipients)
       
        ' recipients              = ""
        ' for each emaildata in Email("recipients")
        '   recipients = recipients + emaildata
        '  next
       '  recipients = mid(recipients,1,len(recipients)-1)
       '  myEACMessage.recipients = recipients
         
         set EAC_WHOLEAttachement = new mailattachment
             if Email("subject")= "string" then
               EAC_WHOLEAttachement.contenttype = "text"
               EAC_WHOLEAttachement.contentsubtype = "wholexml"
             else
               EAC_WHOLEAttachement.contenttype = "wholexml"
             end if
             EAC_WHOLEAttachement.data        = wholexml(vdomxml)
             
             
         'renderingMode = "string"
         'logger ("wholexml >" & wholexml(vdomxml))

         myEACMessage.addattachment(EAC_WHOLEAttachement)
         'server.mailer.send(myEACMessage)
          server.mailer.send_via(myEACMessage, sett)

         sendEACMessage = Dictionary ("AppName",EACAppName,"EAC-Token",EACToken,"EAC-Method",EACMethod,"SessionToken",sessionToken)
  
  end function
end class