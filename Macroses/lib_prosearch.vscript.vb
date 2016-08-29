'Need lib_Utils to work 


function debug(log,line)

  if false then
    logger("{lib_prosearch," & cstr(line) & "}: " & log)
  end if

end function


Class ProSearchApp

  'local data
  dim host
  dim user
  dim password
  dim appSearch
  dim errMsg
  dim searchType
  
  'agent configuration
  dim agentName
  dim sourceName
  dim sourceIndexes
  dim kErrScriptError
  dim kErrNotNeededRules 
  dim kErrSearchSourceObjectAlreadyExist
  dim kErrAgentObjNotExists
  dim kErrSearchSourceObjNotExists
  
  Sub Class_Initialize( )
  
    host = "localhost"
    user = "prosuite"
    password = "prosuite"
    errMsg = ""
    kErrScriptError = 0
    kErrNotNeededRules = 5
    kErrSearchSourceObjectAlreadyExist = 12
    kErrAgentObjNotExists = 13
    kErrSearchSourceObjNotExists = 14
    agentName = ProSearchAgentName
    searchType = "Tasks"
    
  End sub
  
  Function CheckError( respData )
    if respData( 0 ) = "error" then
      if respData( 1 ) = kErrScriptError then
        errMsg = respData( 2 )
      end if
      CheckError = True
    else
      CheckError = False
    end if
  End Function
  
  Function connect
  
    connect = true
    ConnectionResult = ApplicationConnection( "ProSearch", host, user, password )

    Select Case ConnectionResult(0)
      Case 0
        connect = false
        errMsg = FormatString( "Impossible to find application {0}", "ProSearch")
      Case 1
        connect = false
        errMsg = FormatString( "Couldn't connect to server: {0}. Please, check host, login and password", host)
      Case 2
        connect = false
        errMsg = FormatString( "Couldn't connect to {0} application. Please check application login and password", "ProSearch" )
      Case 3
        debug("Succes to connect to ProSearch",76) 
        users = ProAdmin.users( "root" )
        if len( users ) <> 1 then
          connect = false
          errMsg = "Couldn't find root user"
          debug(errMsg,81)
        else
          set RootUser = users(0)
          set AppSearch = ConnectionResult( 1 )
          LoginData = Dictionary( "token", ProAdmin.accessTokenForUser( RootUser ) )
          answer = AsJson( AppSearch.Invoke( "login", ToJson( LoginData ) ) )
          if CheckError( answer ) then
            debug ("Impossible to log in ProSearch",88)
          end if
        end if
    End Select
  
  end function
  
  Function CreateAgent()

    if agentName<>"" then
      CreateAgent = true
    
      AgentData = Dictionary
      AgentData( "name" ) = agentName
      AgentData( "type" ) = normalize(agentName)
  
      metaData = Dictionary
      metaData( "hosts" ) = tojson( System.applicationHosts )
      AgentData( "meta_data" ) = metaData
    
      answer = AsJson( AppSearch.Invoke( "create_agent", ToJson( AgentData ) ) )
    
      if CheckError( answer ) then
        CreateAgent = false
        errMsg = answer(2)
      else
        DBDictionary( "agent_ProSearch_" & agentName ) = answer( 1 )
      end if
    else
      CreateAgent = false
      errMsg = "There is no agent name set up"
    end if
    
  End Function
  

  Function GetAgentId()
    if "agent_ProSearch_" & agentName in DBDictionary then
      GetAgentId = DBDictionary( "agent_ProSearch_" & agentName )
    else
      if CreateAgent() then
        GetAgentId = DBDictionary( "agent_ProSearch_" & agentName )
      else
        debug("Impossible to create Agent",131)
        ExitScript
      end if
    end if
  End Function
  
  Function FindSourceByName( sourceName, agentId )
    answer = AsJson( AppSearch.Invoke( "retrieve_agent_info", ToJson( Dictionary( "id", agentId ))))
    
    answer = answer( 1 )
    sources = answer( "sources" )
    for each source in sources
      if sourceName = sources( source ) then
        FindSourceByName = source
        Exit Function
      end if
    next
    debug("Couldn't find source",148)
    ExitScript
  End Function
  
  Function CreateSource()
  
    CreateSource = true
    SourceData = Dictionary
    SourceData( "name" ) = sourceName
    SourceData( "agent_id" ) = GetAgentId
    metaData = Dictionary
    metaData( "meta_fields" ) = ToJson( sourceIndexes )
    SourceData( "meta_data" ) = metaData
    Flag = True
    dim sourceId
    
    Do While Flag = True
      answer = AsJson( AppSearch.Invoke( "create_search_source", ToJson( SourceData )))
      debug("create_search_source : " & toJson(answer),166)
      if CheckError( answer ) then
        if answer( 1 ) = kErrAgentObjNotExists then
          DBDictionary.remove( "agent_ProSearch_" & cstr(agentName) )
          SourceData( "agent_id" ) = GetAgentId 
        elseif answer( 1 ) = kErrSearchSourceObjectAlreadyExist then
          sourceId = FindSourceByName( sourceName, SourceData( "agent_id" ))
          Flag = False
        else
          errMsg = answer( 2 )
          debug(errMsg,176)
          ExitScript
        end if
     else
       debug ("no error creating source",180)
       Flag = False
       sourceId=answer( 1 )
     end if
    loop
    DBDictionary("source_ProSearch_" & cstr(SourceData( "agent_id" )) & "_" & sourceName) = sourceId
    CreateSource = sourceId 
  End Function

  
  Function GetSourceId()
  
    kSourceId = "source_ProSearch_" & cstr(GetAgentId) & "_" & sourceName 
  
    if kSourceId in DBDictionary then
      GetSourceId = DBDictionary( kSourceId ) 
    else
      GetSourceId = CreateSource
    end if
    
  End Function
  
  Function send(data)
    answer = AsJson( AppSearch.Invoke( "create_index", tojson(data)))
    debug ("answer create_index :" & tojson(answer),207)
    if CheckError( answer ) then
      if answer(1) = kErrSearchSourceObjNotExists then
        debug ("Delete DBSource",210)
        kSourceId = "source_ProSearch_" & cstr(GetAgentId) & "_" & sourceName
        DBDictionary.remove(kSourceId)
        debug ("Create source GetSourceId",213)
        data("source_id")=GetSourceId
        answer = AsJson( AppSearch.Invoke( "create_index", tojson(data)))
        debug ("answer 2 :" & tojson(answer),216)
        if CheckError( answer ) then
          send = false
          errMsg = tojson(answer)
        else
          send = true
        end if
      else
        send = false
        errMsg = tojson(answer)
      end if
    else
      send = true
    end if
  end function
  
  function getData(filepath)
    'first lest get all sources for this agent
    RetrieveAgentInfoData = Dictionary
    RetrieveAgentInfoData( "id" ) = GetAgentId
    answer = AsJson(AppSearch.Invoke( "retrieve_agent_info", tojson(RetrieveAgentInfoData)))

    if CheckError( answer ) then
      getData = Empty
    else
      
      sourceList = answer(1)("sources")
      sourceArray = Array
          
      for each source in sourceList
        AppendToArray(sourceArray,source)
      next
      'search for the file in prosearch
      Query = Dictionary
      Query( "query" ) = filepath
      Query( "access" ) = "w"
      Query( "agent_id" ) = GetAgentId
      Query( "sources" ) = sourceArray
      answer = AsJson( AppSearch.Invoke( "retrieve_results", tojson(Query)))
      
      if CheckError( answer ) then
        getData = Empty
      else
        if len(answer(1))=1 then
          getData=answer(1)
        else
          getData = Empty
        end if
      end if      
    end if
  end function
  
  function delete(recordid)
  
    IndexData = Dictionary
    IndexData( "id" ) = recordid
    
    answer = AsJson( AppSearch.Invoke( "delete_index", tojson(IndexData)))
    if CheckError( answer ) then
      delete = false
      logger ("Error occured while deleting : "& recordid &" with answer :" & tojson(answer))
    else
      delete = true
    end if
    
  end function
  
end class